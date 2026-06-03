import dotenv from "dotenv";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
dotenv.config({ path: path.resolve(__dirname, "../../.env") });

import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";

import { ProviderManager, MultiAccountConfig } from "./auth/providerManager.js";
import { buildMcpServer } from "./server/mcp.js";
import { startHttpServer } from "./server/httpServer.js";
import { GRAPH_DELEGATED_SCOPES } from "./config/scopes.js";

// Load multi-account config (used by the local/stdio transport).
let multiConfig: MultiAccountConfig;

if (process.env.ACCOUNTS_CONFIG) {
  try {
    multiConfig = JSON.parse(process.env.ACCOUNTS_CONFIG);
  } catch (err) {
    console.error("Failed to parse ACCOUNTS_CONFIG env var. Using defaults.");
    multiConfig = getDefaultConfig();
  }
} else {
  multiConfig = getDefaultConfig();
}

function getDefaultConfig(): MultiAccountConfig {
  return {
    accounts: [{ id: "default", name: "Default Outlook", type: "outlook" }],
    common: {
      outlook: {
        clientId: process.env.AZURE_CLIENT_ID || "",
        tenantId: process.env.AZURE_TENANT_ID || "organizations",
        clientSecret: process.env.AZURE_CLIENT_SECRET || "",
        redirectUri: process.env.REDIRECT_URI || "http://localhost:3000/callback",
      },
      gmail: {
        clientId: process.env.GOOGLE_CLIENT_ID || "",
        clientSecret: process.env.GOOGLE_CLIENT_SECRET || "",
        redirectUri: process.env.REDIRECT_URI || "http://localhost:3000/callback",
      },
      draftsFolderName: process.env.DRAFTS_FOLDER_NAME || "aiDrafts",
    },
  };
}

async function startStdio() {
  const manager = new ProviderManager(multiConfig);
  const server = buildMcpServer({
    resolveProvider: async ({ accountId, mailbox }) =>
      (await manager.getProvider(accountId, mailbox)).provider,
    listAccounts: () => manager.getAccounts(),
    draftsFolderName: multiConfig.common.draftsFolderName,
  });

  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("FBI-MCP-O365 Server started (stdio, multi-account)");
}

async function startHttp() {
  const publicBaseUrl = (process.env.PUBLIC_BASE_URL || `http://localhost:${process.env.PORT || 8080}`).replace(
    /\/+$/,
    ""
  );
  const clientId = process.env.AZURE_CLIENT_ID || "";
  const clientSecret = process.env.AZURE_CLIENT_SECRET || "";
  const tenantId = process.env.AZURE_TENANT_ID || "organizations";

  if (!clientId || !clientSecret) {
    console.error(
      "ERROR: AZURE_CLIENT_ID and AZURE_CLIENT_SECRET are required for the remote (http) transport."
    );
    process.exit(1);
  }

  await startHttpServer({
    port: Number(process.env.PORT) || 8080,
    publicBaseUrl,
    tenantId,
    clientId,
    clientSecret,
    graphScopes: GRAPH_DELEGATED_SCOPES,
    draftsFolderName: process.env.DRAFTS_FOLDER_NAME || "aiDrafts",
  });
}

const transport = (process.env.MCP_TRANSPORT || "stdio").toLowerCase();
const main = transport === "http" ? startHttp : startStdio;
main().catch((err) => {
  console.error(err);
  process.exit(1);
});
