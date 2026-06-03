import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";

import { MailProvider, AccountInfo } from "../types/mail.js";
import { toolDefinitions, toolsByName } from "../tools/registry.js";

export interface McpServerDeps {
  // Resolve the concrete provider for a tool call from its account-routing args.
  resolveProvider(args: {
    accountId?: string;
    mailbox?: string;
  }): Promise<MailProvider>;
  // List accounts available to this session (for the list_accounts tool).
  listAccounts(): AccountInfo[] | Promise<AccountInfo[]>;
  draftsFolderName: string;
}

const SERVER_INFO = { name: "aimail-mcp", version: "3.0.0" };

/**
 * Build a fully-wired MCP Server. The transport (stdio or HTTP) is attached by
 * the caller. Tool listing and dispatch come from the shared registry, so both
 * transports expose identical capabilities.
 */
export function buildMcpServer(deps: McpServerDeps): Server {
  const server = new Server(SERVER_INFO, { capabilities: { tools: {} } });

  server.setRequestHandler(ListToolsRequestSchema, async () => ({
    tools: [
      {
        name: "list_accounts",
        description: "List the email/Microsoft 365 accounts available to you.",
        inputSchema: { type: "object", properties: {} },
      },
      ...toolDefinitions(),
    ],
  }));

  // Cast: our ToolResult matches the standard CallTool result shape, but the
  // SDK's union also includes a task-augmented variant that confuses inference.
  server.setRequestHandler(CallToolRequestSchema, (async (request: any) => {
    const { name, arguments: rawArgs } = request.params;
    const args = (rawArgs || {}) as Record<string, unknown>;

    if (name === "list_accounts") {
      const accounts = await deps.listAccounts();
      return { content: [{ type: "text", text: JSON.stringify(accounts, null, 2) }] };
    }

    const tool = toolsByName.get(name);
    if (!tool) {
      return { content: [{ type: "text", text: `Unknown tool: ${name}` }], isError: true };
    }

    try {
      const provider = await deps.resolveProvider({
        accountId: args.accountId as string | undefined,
        mailbox: args.mailbox as string | undefined,
      });
      return await tool.run(provider, args, { draftsFolderName: deps.draftsFolderName });
    } catch (error: any) {
      return { content: [{ type: "text", text: `Error: ${error.message}` }], isError: true };
    }
  }) as any);

  return server;
}
