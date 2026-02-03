import { MailProvider } from "../types/mail.js";

interface Tool {
  name: string;
  description: string;
  inputSchema: {
    type: "object";
    properties: {
      draftId: { type: "string"; description: "ID of the draft to update" };
      subject: { type: "string"; description: "New subject (optional)" };
      body: { type: "string"; description: "New body content (optional)" };
      bodyType: { type: "string"; enum: ["text", "HTML"]; description: "Body format (optional)" };
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

export const updateDraftTool: Tool = {
  name: "update_draft",
  description: "Update an existing email draft.",
  inputSchema: {
    type: "object",
    properties: {
      draftId: { type: "string", description: "ID of the draft to update" },
      subject: { type: "string", description: "New subject (optional)" },
      body: { type: "string", description: "New body content (optional)" },
      bodyType: { type: "string", enum: ["text", "HTML"], description: "Body format (optional)" },
      accountId: { type: "string", description: "Optional account ID" },
      mailbox: { type: "string", description: "Optional shared mailbox email" },
    },
    required: ["draftId"],
  },
  async handler(provider: MailProvider, args: Record<string, unknown>) {
    const draftId = args.draftId as string;
    const { subject, body, bodyType } = args;

    try {
      await provider.updateDraft(draftId, { 
        subject: subject as string, 
        body: body as string, 
        bodyType: bodyType as "text" | "HTML" 
      });

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify({ success: true, message: "Draft updated." }),
          },
        ],
      };
    } catch (err: any) {
      return {
        content: [
          {
            type: "text",
            text: `Error updating draft: ${err.message}`,
          },
        ],
        isError: true,
      };
    }
  },
};