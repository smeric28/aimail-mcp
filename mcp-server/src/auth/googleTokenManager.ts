import { google } from "googleapis";
import { OAuth2Client } from "google-auth-library";
import fs from "fs";
import path from "path";
import os from "os";
import http from "http";
import url from "url";
import { exec } from "child_process";

interface GoogleAuthConfig {
  clientId: string;
  clientSecret: string;
  redirectUri: string;
  scopes: string[];
}

export class GoogleTokenManager {
  private oauth2Client: OAuth2Client;
  private config: GoogleAuthConfig;
  private tokenPath: string;

  constructor(config: GoogleAuthConfig) {
    this.config = config;
    this.oauth2Client = new google.auth.OAuth2(
      config.clientId,
      config.clientSecret,
      config.redirectUri
    );

    const cacheDir = path.join(os.homedir(), ".fbi-mcp-o365-cache");
    if (!fs.existsSync(cacheDir)) {
      fs.mkdirSync(cacheDir, { recursive: true });
    }
    this.tokenPath = path.join(cacheDir, "google-tokens.json");

    if (fs.existsSync(this.tokenPath)) {
      const tokens = JSON.parse(fs.readFileSync(this.tokenPath, "utf-8"));
      this.oauth2Client.setCredentials(tokens);
    }
  }

  private async startRedirectListener(): Promise<string> {
    const redirectUrl = new URL(this.config.redirectUri);
    const port = parseInt(redirectUrl.port) || 3000;

    return new Promise((resolve, reject) => {
      const server = http.createServer(async (req, res) => {
        const parsedUrl = url.parse(req.url || "", true);
        const code = parsedUrl.query.code as string;

        if (code) {
          res.writeHead(200, { "Content-Type": "text/html" });
          res.end(`
            <html>
              <head><title>Authentication Successful</title></head>
              <body style="font-family: sans-serif; background: #0f172a; color: #f8fafc; display: flex; align-items: center; justify-content: center; height: 100vh; margin: 0;">
                <div style="background: #1e293b; padding: 2.5rem; border-radius: 1rem; text-align: center; border: 1px solid #334155;">
                  <h1 style="color: #38bdf8;">Login Successful</h1>
                  <p>Google authentication complete. You can close this window.</p>
                </div>
                <script>setTimeout(() => window.close(), 3000);</script>
              </body>
            </html>
          `);
          server.close();
          resolve(code);
        } else {
          res.end("Error: No code found");
          reject(new Error("No code found"));
        }
      });

      server.listen(port);
    });
  }

  async getClient(): Promise<any> {
    // Check if we have valid tokens
    const tokens = this.oauth2Client.credentials;
    if (!tokens || !tokens.refresh_token) {
      await this.authenticate();
    } else {
      // Check if expired
      const expiryDate = tokens.expiry_date || 0;
      const isExpired = expiryDate <= Date.now() + 10000; // 10s buffer

      if (isExpired) {
        try {
          const { credentials } = await this.oauth2Client.refreshAccessToken();
          this.saveTokens(credentials);
        } catch (err) {
          console.error("Failed to refresh Google token, re-authenticating...");
          await this.authenticate();
        }
      }
    }

    return google.gmail({ version: "v1", auth: this.oauth2Client });
  }

  private async authenticate() {
    const authUrl = this.oauth2Client.generateAuthUrl({
      access_type: "offline",
      scope: this.config.scopes,
      prompt: "consent",
    });

    console.error(`\n🔐 Google Authentication Required: ${authUrl}\n`);

    let command = "";
    if (process.platform === "win32") {
      command = `start "" "${authUrl}"`;
    } else if (process.platform === "darwin") {
      command = `open "${authUrl}"`;
    } else {
      command = `xdg-open "${authUrl}" || sensible-browser "${authUrl}"`;
    }

    exec(command);

    const code = await this.startRedirectListener();
    const { tokens } = await this.oauth2Client.getToken(code);
    this.oauth2Client.setCredentials(tokens);
    this.saveTokens(tokens);
  }

  private saveTokens(tokens: any) {
    // Merge tokens if we already have some (to keep refresh_token if it's not returned)
    let currentTokens = {};
    if (fs.existsSync(this.tokenPath)) {
      currentTokens = JSON.parse(fs.readFileSync(this.tokenPath, "utf-8"));
    }
    const mergedTokens = { ...currentTokens, ...tokens };
    fs.writeFileSync(this.tokenPath, JSON.stringify(mergedTokens, null, 2));
  }
}
