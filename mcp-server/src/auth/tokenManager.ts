import {
  ConfidentialClientApplication,
  Configuration,
  AuthenticationResult,
  ICachePlugin,
  TokenCacheContext,
} from "@azure/msal-node";
import { Client } from "@microsoft/microsoft-graph-client";
import fs from "fs";
import path from "path";
import os from "os";
import http from "http";
import url from "url";
import { exec } from "child_process";

interface AuthConfig {
  clientId: string;
  tenantId: string;
  clientSecret?: string;
  redirectUri: string;
  scopes: string[];
}

interface CachedToken {
  accessToken: string;
  refreshToken?: string; // This is actually the homeAccountId in this implementation
  expiresOn: number;
}

export class TokenManager {
  private msalClient: ConfidentialClientApplication;
  private config: AuthConfig;
  private msalCachePath: string;

  constructor(config: AuthConfig) {
    this.config = config;

    // Use a stable, OS-level cache directory
    const cacheDir = path.join(os.homedir(), ".fbi-mcp-o365-cache");
    if (!fs.existsSync(cacheDir)) {
      console.error(`Creating cache directory at: ${cacheDir}`);
      fs.mkdirSync(cacheDir, { recursive: true });
    }
    this.msalCachePath = path.join(cacheDir, "msal-cache.json");

    const cachePlugin: ICachePlugin = {
      beforeCacheAccess: async (cacheContext: TokenCacheContext) => {
        if (fs.existsSync(this.msalCachePath)) {
          try {
            const cacheData = fs.readFileSync(this.msalCachePath, "utf-8");
            cacheContext.tokenCache.deserialize(cacheData);
          } catch (err) {
            console.error("Failed to deserialize MSAL cache:", err);
          }
        }
      },
      afterCacheAccess: async (cacheContext: TokenCacheContext) => {
        if (cacheContext.cacheHasChanged) {
          try {
            const cacheData = cacheContext.tokenCache.serialize();
            fs.writeFileSync(this.msalCachePath, cacheData);
          } catch (err) {
            console.error("Failed to serialize MSAL cache:", err);
          }
        }
      },
    };

    const msalConfig: Configuration = {
      auth: {
        clientId: config.clientId,
        authority: `https://login.microsoftonline.com/${config.tenantId}`,
        clientSecret: config.clientSecret,
      },
      cache: {
        cachePlugin,
      },
    };
    this.msalClient = new ConfidentialClientApplication(msalConfig);
    console.error(`Token cache management initialized at: ${this.msalCachePath}`);
  }

  private async startRedirectListener(): Promise<{ code: string; state: string }> {
    const redirectUrl = new URL(this.config.redirectUri);
    const port = parseInt(redirectUrl.port) || 3000;

    return new Promise((resolve, reject) => {
      const server = http.createServer((req, res) => {
        const parsedUrl = url.parse(req.url || "", true);
        const query = parsedUrl.query;
        const code = query.code as string;
        const state = query.state as string;

        if (code) {
          res.writeHead(200, { "Content-Type": "text/html" });
          res.end(`
            <html>
              <head>
                <title>Authentication Successful</title>
                <style>
                  body { 
                    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; 
                    background: #0f172a; 
                    color: #f8fafc; 
                    display: flex; 
                    align-items: center; 
                    justify-content: center; 
                    height: 100vh; 
                    margin: 0; 
                  }
                  .card { 
                    background: #1e293b; 
                    padding: 2.5rem; 
                    border-radius: 1rem; 
                    box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.3), 0 10px 10px -5px rgba(0, 0, 0, 0.1); 
                    text-align: center; 
                    max-width: 450px;
                    border: 1px solid #334155;
                  }
                  h1 { color: #38bdf8; margin-top: 0; font-weight: 700; }
                  p { color: #94a3b8; line-height: 1.6; font-size: 1.1rem; }
                  .icon { 
                    display: inline-flex;
                    align-items: center;
                    justify-content: center;
                    width: 64px;
                    height: 64px;
                    background: #064e3b;
                    color: #10b981;
                    border-radius: 50%;
                    font-size: 32px;
                    margin-bottom: 1.5rem;
                  }
                  .footer { margin-top: 2rem; font-size: 0.875rem; color: #64748b; }
                </style>
              </head>
              <body>
                <div class="card">
                  <div class="icon">✓</div>
                  <h1>Login Successful</h1>
                  <p>Authentication complete. You can close this window and return to your application.</p>
                  <div class="footer">This window will attempt to close automatically.</div>
                </div>
                <script>
                  setTimeout(() => {
                    window.close();
                    // Fallback for browsers that don't allow script-initiated close
                    document.querySelector('.footer').innerText = "Safe to close this window.";
                  }, 3000);
                </script>
              </body>
            </html>
          `);
          server.close();
          resolve({ code, state: state || "" });
        } else {
          res.writeHead(400, { "Content-Type": "text/plain" });
          res.end("Authorization code not found in redirect.");
          reject(new Error("Authorization code not found in redirect."));
        }
      });

      server.on("error", (err: any) => {
        if (err.code === "EADDRINUSE") {
          console.error(`Port ${port} is in use. Assuming auth is in progress or stalled.`);
          reject(new Error(`Port ${port} is already in use. Cannot start auth listener.`));
        } else {
          reject(err);
        }
      });

      server.listen(port, () => {
        // Timeout after 5 minutes to prevent hanging
        setTimeout(() => {
          if (server.listening) {
            console.error("Auth server timed out. Closing.");
            server.close();
          }
        }, 300000);
      });
    });
  }

  async getToken(): Promise<string> {
    try {
      // 1. Try Silent Refresh using the internal MSAL cache
      const accounts = await this.msalClient.getTokenCache().getAllAccounts();

      if (accounts.length > 0) {
        try {
          // console.error("Attempting silent refresh with cached account...");
          const result = await this.msalClient.acquireTokenSilent({
            account: accounts[0],
            scopes: this.config.scopes,
          });
          if (result) {
            // console.error("Silent refresh successful.");
            return result.accessToken;
          }
        } catch (refreshErr) {
          // console.error("Silent refresh failed, will attempt interactive flow if needed.");
        }
      }
    } catch (err) {
      console.error("Error checking MSAL cache:", err);
    }

    // 2. Interactive Flow (Authorization Code)
    console.error("\n🔐 Session expired or not found. Action Required: Please complete login in the browser.");

    try {
      const authCodeUrlParameters = {
        scopes: this.config.scopes,
        redirectUri: this.config.redirectUri,
      };

      const authUrl = await this.msalClient.getAuthCodeUrl(authCodeUrlParameters);
      console.error(`Opening: ${authUrl}\n`);

      let command = "";
      if (process.platform === "win32") {
        command = `start "" "${authUrl}"`;
      } else if (process.platform === "darwin") {
        command = `open "${authUrl}"`;
      } else {
        command = `xdg-open "${authUrl}" || sensible-browser "${authUrl}" || echo "Please open this URL manually: ${authUrl}"`;
      }

      exec(command, (err: any) => {
        if (err) {
          console.error("Failed to open browser. Please open the URL manually:", authUrl);
        }
      });

      const { code } = await this.startRedirectListener();

      const tokenRequest = {
        code: code,
        scopes: this.config.scopes,
        redirectUri: this.config.redirectUri,
      };

      const response = await this.msalClient.acquireTokenByCode(tokenRequest);

      if (response) {
        console.error("Token acquired successfully.");
        return response.accessToken;
      } else {
        throw new Error("Failed to acquire token: response is null");
      }
    } catch (error) {
      throw new Error(`Failed to acquire token: ${error}`);
    }
  }

  async getClient(): Promise<Client> {
    const token = await this.getToken();
    return Client.init({
      authProvider: (done) => {
        done(null, token);
      },
    });
  }
}