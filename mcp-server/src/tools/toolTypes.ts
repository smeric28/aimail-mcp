import { MailProvider } from "../types/mail.js";

export interface ToolContext {
  draftsFolderName: string;
}

export interface ToolResult {
  content: Array<{ type: "text"; text: string }>;
  isError?: boolean;
}

export interface RegisteredTool {
  name: string;
  description: string;
  inputSchema: {
    type: "object";
    properties: Record<string, any>;
    required?: string[];
  };
  run: (
    provider: MailProvider,
    args: Record<string, unknown>,
    ctx: ToolContext
  ) => Promise<ToolResult>;
}

export function ok(data: unknown): ToolResult {
  return {
    content: [
      {
        type: "text",
        text: typeof data === "string" ? data : JSON.stringify(data, null, 2),
      },
    ],
  };
}

export function fail(message: string): ToolResult {
  return { content: [{ type: "text", text: `Error: ${message}` }], isError: true };
}

// Standard account-routing fields shared by every tool. The transport layer
// resolves these into a concrete provider before run() is called.
export const accountFields = {
  accountId: { type: "string", description: "Optional account ID (see list_accounts)" },
  mailbox: { type: "string", description: "Optional shared mailbox email to act on" },
};
