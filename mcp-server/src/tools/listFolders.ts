import { MailProvider } from "../types/mail.js";

interface Tool {
  name: string;
  description: string;
  inputSchema: {
    type: "object";
    properties: {
      accountId: { type: "string"; description: "Optional account ID" };
      mailbox: { type: "string"; description: "Optional shared mailbox email" };
    };
  };
  handler: (
    provider: MailProvider,
    args: Record<string, unknown>
  ) => Promise<{
    content: Array<{ type: "text"; text: string }>;
    isError?: boolean;
  }>;
}

export const listFoldersTool: Tool = {
  name: "list_mail_folders",
  description: "List available mail folders/labels in the account.",
  inputSchema: {
    type: "object",
    properties: {
      accountId: {
        type: "string",
        description: "Optional account ID",
      },
      mailbox: {
        type: "string",
        description: "Optional shared mailbox email",
      },
    },
  },
  async handler(provider: MailProvider, args: Record<string, unknown>) {
    try {
      const folders = await provider.listFolders();

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(folders, null, 2),
          },
        ],
      };
    } catch (err: any) {
      return {
        content: [
          {
            type: "text",
            text: `Error listing folders: ${err.message}`,
          },
        ],
        isError: true,
      };
    }
  },
};
