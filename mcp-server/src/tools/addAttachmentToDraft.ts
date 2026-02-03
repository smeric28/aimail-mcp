import { MailProvider } from "../types/mail.js";

interface Tool {
  name: string;
  description: string;
  inputSchema: {
    type: "object";
    properties: {
      draftId: { type: "string"; description: "ID of the draft" };
      name: { type: "string"; description: "Attachment filename" };
      contentType: { type: "string"; description: "MIME type" };
      contentBytes: { type: "string"; description: "Base64 encoded content" };
      accountId: { type: "string"; description: "Optional account ID" };
      mailbox: { type: "string"; description: "Optional shared mailbox email" };
    };
    required: ["draftId", "name", "contentType", "contentBytes"];
  };
  handler: (
    provider: MailProvider,
    args: Record<string, unknown>
  ) => Promise<{
    content: Array<{ type: "text"; text: string }>;
    isError?: boolean;
  }>;
}

export const addAttachmentToDraftTool: Tool = {
  name: "add_attachment_to_draft",
  description: "Add a file attachment to an existing draft.",
  inputSchema: {
    type: "object",
    properties: {
      draftId: { type: "string", description: "ID of the draft" },
      name: { type: "string", description: "Attachment filename" },
      contentType: { type: "string", description: "MIME type" },
      contentBytes: { type: "string", description: "Base64 encoded content" },
      accountId: { type: "string", description: "Optional account ID" },
      mailbox: { type: "string", description: "Optional shared mailbox email" },
    },
    required: ["draftId", "name", "contentType", "contentBytes"],
  },
  async handler(provider: MailProvider, args: Record<string, unknown>) {
    const { draftId, name, contentType, contentBytes } = args;

    try {
      await provider.addAttachmentToDraft(draftId as string, {
        name: name as string,
        contentType: contentType as string,
        contentBytes: contentBytes as string,
      });

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify({ success: true, message: "Attachment added." }),
          },
        ],
      };
    } catch (err: any) {
      return {
        content: [
          {
            type: "text",
            text: `Error adding attachment: ${err.message}`,
          },
        ],
        isError: true,
      };
    }
  },
};
