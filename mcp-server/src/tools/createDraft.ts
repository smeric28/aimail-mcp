import { MailProvider } from "../types/mail.js";

interface Tool {
  name: string;
  description: string;
  inputSchema: {
    type: "object";
    properties: {
      to: { type: "string"; description: "Recipient email address" };
      subject: { type: "string"; description: "Email subject" };
      body: { type: "string"; description: "Email body" };
      bodyType: { type: "string"; enum: ["text", "HTML"]; description: "Body format" };
      originalEmailId: { type: "string"; description: "ID of email to reply to (optional)" };
      accountId: { type: "string"; description: "Optional account ID" };
      mailbox: { type: "string"; description: "Optional shared mailbox email" };
    };
    required: ["to", "subject", "body"];
  };
  handler: (
    provider: MailProvider,
    args: Record<string, unknown>,
    config: { draftsFolderName: string }
  ) => Promise<{
    content: Array<{ type: "text"; text: string }>;
    isError?: boolean;
  }>;
}

export const createDraftTool: Tool = {
  name: "create_draft",
  description:
    "Create an email draft for manual review. " +
    "Drafts are NOT sent automatically. " +
    "This is the ONLY write operation permitted by this MCP server.",
  inputSchema: {
    type: "object",
    properties: {
      to: { type: "string", description: "Recipient email address" },
      subject: { type: "string", description: "Email subject" },
      body: { type: "string", description: "Email body" },
      bodyType: { type: "string", enum: ["text", "HTML"], description: "Body format" },
      originalEmailId: { type: "string", description: "ID of email to reply to (optional)" },
      accountId: { type: "string", description: "Optional account ID" },
      mailbox: { type: "string", description: "Optional shared mailbox email" },
    },
    required: ["to", "subject", "body"],
  },
  async handler(provider: MailProvider, args: Record<string, unknown>) {
    const to = args.to as string;
    const subject = args.subject as string;
    const body = args.body as string;
    const bodyType = (args.bodyType as "text" | "HTML") || "text";
    const originalEmailId = args.originalEmailId as string | undefined;

    try {
      const draftId = await provider.createDraft({
        to,
        subject,
        body,
        bodyType,
        originalEmailId,
      });

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify({
              success: true,
              draftId,
              message: "Draft created successfully. Review and send manually.",
            }, null, 2),
          },
        ],
      };
    } catch (err: any) {
      return {
        content: [
          {
            type: "text",
            text: `Error creating draft: ${err.message}`,
          },
        ],
        isError: true,
      };
    }
  },
};