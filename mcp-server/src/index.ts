import dotenv from "dotenv";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
dotenv.config({ path: path.resolve(__dirname, "../../.env") });

import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";

import { ProviderManager, MultiAccountConfig } from "./auth/providerManager.js";
import { searchEmailsTool } from "./tools/searchEmails.js";
import { getEmailTool } from "./tools/getEmail.js";
import { createDraftTool } from "./tools/createDraft.js";
import { updateDraftTool } from "./tools/updateDraft.js";
import { deleteDraftTool } from "./tools/deleteDraft.js";
import { moveEmailTool } from "./tools/moveEmail.js";
import { listFoldersTool } from "./tools/listFolders.js";

// Load config
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
    accounts: [
      { id: "default", name: "Default Outlook", type: "outlook" }
    ],
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
    }
  };
}

async function main() {
  const manager = new ProviderManager(multiConfig);
  
  const server = new Server(
    { name: "aimail-mcp", version: "2.0.0" },
    { capabilities: { tools: {} } }
  );

  // Handle tool listing
  server.setRequestHandler(ListToolsRequestSchema, async () => {
    return {
      tools: [
        {
          name: "list_accounts",
          description: "List configured email accounts.",
          inputSchema: { type: "object", properties: {} }
        },
        searchEmailsTool,
        getEmailTool,
        createDraftTool,
        updateDraftTool,
        deleteDraftTool,
        moveEmailTool,
        listFoldersTool,
      ],
    };
  });

  // Handle tool calls
  server.setRequestHandler(CallToolRequestSchema, async (request) => {
    const { name, arguments: args } = request.params;
    const accountId = args?.accountId as string | undefined;
    const mailbox = args?.mailbox as string | undefined;

    if (name === "list_accounts") {
      return {
        content: [{ type: "text", text: JSON.stringify(manager.getAccounts(), null, 2) }]
      };
    }

    try {
      const { provider } = await manager.getProvider(accountId, mailbox);

      switch (name) {
        case "search_emails":
          return await searchEmailsTool.handler(provider, args || {}, multiConfig.common);
        case "get_email":
          return await getEmailTool.handler(provider, args || {});
        case "create_draft":
          return await createDraftTool.handler(provider, args || {}, multiConfig.common);
        case "update_draft":
          return await updateDraftTool.handler(provider, args || {});
        case "delete_draft":
          return await deleteDraftTool.handler(provider, args || {});
        case "move_email":
          return await moveEmailTool.handler(provider, args || {});
        case "list_mail_folders":
          return await listFoldersTool.handler(provider, args || {});
        default:
          return {
            content: [{ type: "text", text: `Unknown tool: ${name}` }],
            isError: true,
          };
      }
    } catch (error: any) {
      return {
        content: [{ type: "text", text: `Error: ${error.message}` }],
        isError: true,
      };
    }
  });

  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("AI Mail MCP Server started with Multi-Account support");
}

main().catch(console.error);
