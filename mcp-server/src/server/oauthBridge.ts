import express, { Router } from "express";
import crypto from "crypto";

/**
 * OAuth 2.1 bridge that lets a Claude custom connector authenticate users
 * against Microsoft Entra ID.
 *
 * Why a bridge? Claude's remote-connector flow discovers an authorization
 * server, dynamically registers a client, and runs Authorization Code + PKCE.
 * Microsoft Entra ID does NOT support open Dynamic Client Registration, so this
 * server presents itself as the authorization server to Claude and brokers the
 * real flow to Entra using a single pre-registered Entra app (the one already
 * registered in the Fireball tenant).
 *
 * Token model: we request Microsoft Graph delegated scopes from Entra, so the
 * access tokens handed back to Claude are Graph-audience tokens. The /mcp
 * handler forwards them straight to Graph (token passthrough). This is the
 * simplest correct model for an internal connector; upgrade to On-Behalf-Of if
 * you later want our own API audience.
 */

export interface BridgeConfig {
  publicBaseUrl: string; // e.g. https://aimail.fireballz.ai
  tenantId: string;
  clientId: string;
  clientSecret: string;
  scopes: string[]; // Graph delegated scopes (offline_access added automatically)
}

interface PendingAuth {
  claudeRedirectUri: string;
  claudeState?: string;
  claudeCodeChallenge: string;
  claudeCodeChallengeMethod: string;
  entraCodeVerifier: string;
  createdAt: number;
}

interface IssuedCode {
  entraTokens: TokenResponse;
  claudeCodeChallenge: string;
  claudeCodeChallengeMethod: string;
  createdAt: number;
}

interface TokenResponse {
  access_token: string;
  refresh_token?: string;
  expires_in: number;
  token_type: string;
  scope?: string;
}

const TEN_MINUTES = 10 * 60 * 1000;

function base64url(buf: Buffer): string {
  return buf.toString("base64").replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

function verifyPkce(verifier: string, challenge: string, method: string): boolean {
  if (method === "plain") return verifier === challenge;
  const hash = base64url(crypto.createHash("sha256").update(verifier).digest());
  return hash === challenge;
}

export function createOAuthBridge(config: BridgeConfig): Router {
  const router = express.Router();
  const authority = `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0`;
  const callbackUrl = `${config.publicBaseUrl}/auth/callback`;
  const scopes = Array.from(new Set([...config.scopes, "offline_access", "openid", "profile"]));

  // Short-lived in-memory state. For multi-instance hosting, back these with a
  // shared store (e.g. Redis); a single Container App replica is fine to start.
  const pending = new Map<string, PendingAuth>();
  const issuedCodes = new Map<string, IssuedCode>();

  const sweep = () => {
    const now = Date.now();
    for (const [k, v] of pending) if (now - v.createdAt > TEN_MINUTES) pending.delete(k);
    for (const [k, v] of issuedCodes) if (now - v.createdAt > TEN_MINUTES) issuedCodes.delete(k);
  };

  // --- Authorization Server Metadata (RFC 8414) ---
  router.get("/.well-known/oauth-authorization-server", (_req, res) => {
    res.json({
      issuer: config.publicBaseUrl,
      authorization_endpoint: `${config.publicBaseUrl}/authorize`,
      token_endpoint: `${config.publicBaseUrl}/token`,
      registration_endpoint: `${config.publicBaseUrl}/register`,
      response_types_supported: ["code"],
      grant_types_supported: ["authorization_code", "refresh_token"],
      code_challenge_methods_supported: ["S256", "plain"],
      token_endpoint_auth_methods_supported: ["none"],
      scopes_supported: scopes,
    });
  });

  // --- Dynamic Client Registration shim (RFC 7591) ---
  // We accept any registration and echo back a public client bound to the
  // single Entra app behind this bridge.
  router.post("/register", express.json(), (req, res) => {
    const redirectUris = req.body?.redirect_uris || [];
    res.status(201).json({
      client_id: `claude-connector-${base64url(crypto.randomBytes(6))}`,
      client_id_issued_at: Math.floor(Date.now() / 1000),
      redirect_uris: redirectUris,
      token_endpoint_auth_method: "none",
      grant_types: ["authorization_code", "refresh_token"],
      response_types: ["code"],
    });
  });

  // --- Authorization endpoint: redirect the user to Entra ---
  router.get("/authorize", (req, res) => {
    sweep();
    const { redirect_uri, state, code_challenge, code_challenge_method } = req.query as Record<
      string,
      string
    >;
    if (!redirect_uri || !code_challenge) {
      res.status(400).send("Missing redirect_uri or code_challenge");
      return;
    }

    const proxyState = base64url(crypto.randomBytes(24));
    const entraVerifier = base64url(crypto.randomBytes(32));
    const entraChallenge = base64url(crypto.createHash("sha256").update(entraVerifier).digest());

    pending.set(proxyState, {
      claudeRedirectUri: redirect_uri,
      claudeState: state,
      claudeCodeChallenge: code_challenge,
      claudeCodeChallengeMethod: code_challenge_method || "S256",
      entraCodeVerifier: entraVerifier,
      createdAt: Date.now(),
    });

    const url = new URL(`${authority}/authorize`);
    url.searchParams.set("client_id", config.clientId);
    url.searchParams.set("response_type", "code");
    url.searchParams.set("redirect_uri", callbackUrl);
    url.searchParams.set("response_mode", "query");
    url.searchParams.set("scope", scopes.join(" "));
    url.searchParams.set("state", proxyState);
    url.searchParams.set("code_challenge", entraChallenge);
    url.searchParams.set("code_challenge_method", "S256");
    res.redirect(url.toString());
  });

  // --- Callback from Entra: exchange code, mint our own code for Claude ---
  router.get("/auth/callback", async (req, res) => {
    const { code, state, error, error_description } = req.query as Record<string, string>;
    if (error) {
      res.status(400).send(`Entra error: ${error} - ${error_description || ""}`);
      return;
    }
    const record = state ? pending.get(state) : undefined;
    if (!code || !record) {
      res.status(400).send("Invalid or expired authorization state");
      return;
    }
    pending.delete(state);

    try {
      const tokens = await exchangeEntraCode(authority, config, callbackUrl, code, record.entraCodeVerifier);
      const ourCode = base64url(crypto.randomBytes(24));
      issuedCodes.set(ourCode, {
        entraTokens: tokens,
        claudeCodeChallenge: record.claudeCodeChallenge,
        claudeCodeChallengeMethod: record.claudeCodeChallengeMethod,
        createdAt: Date.now(),
      });
      const redirect = new URL(record.claudeRedirectUri);
      redirect.searchParams.set("code", ourCode);
      if (record.claudeState) redirect.searchParams.set("state", record.claudeState);
      res.redirect(redirect.toString());
    } catch (e: any) {
      res.status(502).send(`Token exchange failed: ${e.message}`);
    }
  });

  // --- Token endpoint: authorization_code + refresh_token ---
  router.post("/token", express.urlencoded({ extended: true }), express.json(), async (req, res) => {
    sweep();
    const grantType = req.body?.grant_type;

    try {
      if (grantType === "authorization_code") {
        const { code, code_verifier } = req.body;
        const issued = code ? issuedCodes.get(code) : undefined;
        if (!issued) {
          res.status(400).json({ error: "invalid_grant", error_description: "Unknown or expired code" });
          return;
        }
        issuedCodes.delete(code);
        if (
          !code_verifier ||
          !verifyPkce(code_verifier, issued.claudeCodeChallenge, issued.claudeCodeChallengeMethod)
        ) {
          res.status(400).json({ error: "invalid_grant", error_description: "PKCE verification failed" });
          return;
        }
        res.json(issued.entraTokens);
        return;
      }

      if (grantType === "refresh_token") {
        const tokens = await refreshEntraToken(authority, config, req.body?.refresh_token);
        res.json(tokens);
        return;
      }

      res.status(400).json({ error: "unsupported_grant_type" });
    } catch (e: any) {
      res.status(502).json({ error: "server_error", error_description: e.message });
    }
  });

  return router;
}

async function exchangeEntraCode(
  authority: string,
  config: BridgeConfig,
  redirectUri: string,
  code: string,
  codeVerifier: string
): Promise<TokenResponse> {
  const body = new URLSearchParams({
    client_id: config.clientId,
    client_secret: config.clientSecret,
    grant_type: "authorization_code",
    code,
    redirect_uri: redirectUri,
    code_verifier: codeVerifier,
  });
  return entraToken(authority, body);
}

async function refreshEntraToken(
  authority: string,
  config: BridgeConfig,
  refreshToken?: string
): Promise<TokenResponse> {
  if (!refreshToken) throw new Error("Missing refresh_token");
  const body = new URLSearchParams({
    client_id: config.clientId,
    client_secret: config.clientSecret,
    grant_type: "refresh_token",
    refresh_token: refreshToken,
  });
  return entraToken(authority, body);
}

async function entraToken(authority: string, body: URLSearchParams): Promise<TokenResponse> {
  const resp = await fetch(`${authority}/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: body.toString(),
  });
  const data = (await resp.json()) as any;
  if (!resp.ok) {
    throw new Error(data.error_description || data.error || `HTTP ${resp.status}`);
  }
  return data as TokenResponse;
}
