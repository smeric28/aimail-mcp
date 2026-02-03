import { MailProvider } from "../types/mail.js";

interface Tool {
  name: string;
  description: string;
  inputSchema: {
    type: "object";
    properties: {
      emailId: { type: "string"; description: "ID of the email to move" };
      folder: { type: "string"; description: "Target folder name or ID" };
      accountId: { type: "string"; description: "Optional account ID" };
      mailbox: { type: "string"; description: "Optional shared mailbox email" };
    };
    required: ["emailId", "folder"];
  };
  handler: (
    provider: MailProvider,
    args: Record<string, unknown>
  ) => Promise<{
    content: Array<{ type: "text"; text: string }>;
    isError?: boolean;
  }>;
}

export const moveEmailTool: Tool = {
  name: "move_email",
  description: "Move an email to a different folder.",
  inputSchema: {
    type: "object",
    properties: {
      emailId: { type: "string", description: "ID of the email to move" },
      folder: { type: "string", description: "Target folder name or ID" },
      accountId: { type: "string", description: "Optional account ID" },
      mailbox: { type: "string", description: "Optional shared mailbox email" },
    },
    required: ["emailId", "folder"],
  },
  async handler(provider: MailProvider, args: Record<string, unknown>) {
    const emailId = args.emailId as string;
    const folder = args.folder as string;

    try {
      // For move, we might need to resolve folder name to ID if it's a name
      let folderId = folder;
      const folders = await provider.listFolders();
      const targetFolder = folders.find(f => f.displayName.toLowerCase() === folder.toLowerCase() || f.id === folder);
      if (targetFolder) {
        folderId = targetFolder.id;
      }

      await provider.moveEmail(emailId, folderId);

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify({ success: true, message: `Moved to ${folderId}.` }),
          },
        ],
      };
    } catch (err: any) {
      return {
        content: [
          {
            type: "text",
            text: `Error moving email: ${err.message}`,
          },
        ],
        isError: true,
      };
    }
  },
};
