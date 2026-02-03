import { MailProvider } from "../types/mail.js";

interface Tool {
  name: string;
  description: string;
  inputSchema: {
    type: "object";
    properties: {
      draftId: { type: "string"; description: "ID of the draft to delete" };
      accountId: { type: "string"; description: "Optional account ID" };
      mailbox: { type: "string"; description: "Optional shared mailbox email" };
    };
    required: ["draftId"];
  };
  handler: (
    provider: MailProvider,
    args: Record<string, unknown>
  ) => Promise<{
    content: Array<{ type: "text"; text: string }>;
    isError?: boolean;
  }>;
}

export const deleteDraftTool: Tool = {
  name: "delete_draft",
  description: "Delete an email draft.",
  inputSchema: {
    type: "object",
    properties: {
      draftId: {
        type: "string",
        description: "ID of the draft to delete",
      },
      accountId: {
        type: "string",
        description: "Optional account ID",
      },
      mailbox: {
        type: "string",
        description: "Optional shared mailbox email",
      },
    },
    required: ["draftId"],
  },
  async handler(provider: MailProvider, args: Record<string, unknown>) {
    const draftId = args.draftId as string;

    try {
      await provider.deleteDraft(draftId);

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify({ success: true, message: "Draft deleted." }),
          },
        ],
      };
    } catch (err: any) {
      return {
        content: [
          {
            type: "text",
            text: `Error deleting draft: ${err.message}`,
          },
        ],
        isError: true,
      };
    }
  },
};