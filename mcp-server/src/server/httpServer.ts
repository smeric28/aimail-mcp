import express from "express";
import cors from "cors";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";

import { buildMcpServer } from "./mcp.js";
import { createOAuthBridge } from "./oauthBridge.js";
import { MicrosoftProvider } from "../providers/microsoft.js";
import { MailProvider, AccountInfo } from "../types/mail.js";

export interface HttpServerConfig {
  port: number;
  publicBaseUrl: string;
  tenantId: string;
  clientId: string;
  clientSecret: string;
  graphScopes: string[];
  draftsFolderName: string;
}

function bearerToken(req: express.Request): string | undefined {
  const header = req.headers.authorization || "";
  const match = header.match(/^Bearer\s+(.+)$/i);
  return match?.[1];
}

export async function startHttpServer(config: HttpServerConfig): Promise<void> {
  const app = express();
  app.set("trust proxy", true);

  app.use(
    cors({
      origin: true,
      exposedHeaders: ["Mcp-Session-Id", "WWW-Authenticate"],
      allowedHeaders: ["Content-Type", "Authorization", "Mcp-Session-Id", "MCP-Protocol-Version"],
    })
  );

  app.get("/healthz", (_req, res) => res.json({ status: "ok", service: "aimail-mcp" }));

  // OAuth: authorization-server bridge to Entra (authorize/callback/token/register).
  app.use(
    createOAuthBridge({
      publicBaseUrl: config.publicBaseUrl,
      tenantId: config.tenantId,
      clientId: config.clientId,
      clientSecret: config.clientSecret,
      scopes: config.graphScopes,
    })
  );

  // OAuth 2.0 Protected Resource Metadata (RFC 9728) — tells Claude where to authenticate.
  const resourceMetadata = {
    resource: `${config.publicBaseUrl}/mcp`,
    authorization_servers: [config.publicBaseUrl],
    scopes_supported: config.graphScopes,
    bearer_methods_supported: ["header"],
  };
  app.get("/.well-known/oauth-protected-resource", (_req, res) => res.json(resourceMetadata));
  app.get("/.well-known/oauth-protected-resource/mcp", (_req, res) => res.json(resourceMetadata));

  const challenge = (res: express.Response) => {
    res
      .status(401)
      .set(
        "WWW-Authenticate",
        `Bearer resource_metadata="${config.publicBaseUrl}/.well-known/oauth-protected-resource"`
      )
      .json({ error: "unauthorized", error_description: "Microsoft sign-in required." });
  };

  // The MCP endpoint. Stateless: a fresh server+transport per request, bound to
  // the caller's per-user Graph token.
  app.post("/mcp", express.json({ limit: "25mb" }), async (req, res) => {
    const token = bearerToken(req);
    if (!token) return challenge(res);

    const account: AccountInfo = {
      id: "me",
      name: "Signed-in Microsoft 365 user",
      email: "",
      type: "outlook",
    };

    const resolveProvider = async ({ mailbox }: { mailbox?: string }): Promise<MailProvider> =>
      MicrosoftProvider.fromAccessToken(token, config.draftsFolderName, mailbox);

    const server = buildMcpServer({
      resolveProvider,
      listAccounts: () => [account],
      draftsFolderName: config.draftsFolderName,
    });
    const transport = new StreamableHTTPServerTransport({ sessionIdGenerator: undefined });

    res.on("close", () => {
      transport.close();
      server.close();
    });

    try {
      await server.connect(transport);
      await transport.handleRequest(req, res, req.body);
    } catch (err: any) {
      if (!res.headersSent) {
        res.status(500).json({
          jsonrpc: "2.0",
          error: { code: -32603, message: `Internal error: ${err.message}` },
          id: null,
        });
      }
    }
  });

  // Stateless transport does not use server-initiated SSE or session deletion.
  const methodNotAllowed = (_req: express.Request, res: express.Response) =>
    res.status(405).json({
      jsonrpc: "2.0",
      error: { code: -32000, message: "Method not allowed. Use POST." },
      id: null,
    });
  app.get("/mcp", methodNotAllowed);
  app.delete("/mcp", methodNotAllowed);

  await new Promise<void>((resolve) => {
    app.listen(config.port, () => {
      console.error(`AI Mail MCP remote server listening on :${config.port}`);
      console.error(`Public base URL: ${config.publicBaseUrl}`);
      console.error(`MCP endpoint:    ${config.publicBaseUrl}/mcp`);
      resolve();
    });
  });
}
