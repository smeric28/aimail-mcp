import { MailProvider } from "../types/mail.js";

interface Tool {
  name: string;
  description: string;
  inputSchema: {
    type: "object";
    properties: {
      emailId: { type: "string"; description: "Email ID to retrieve" };
      accountId: { type: "string"; description: "Optional account ID" };
      mailbox: { type: "string"; description: "Optional shared mailbox email" };
    };
    required: ["emailId"];
  };
  handler: (
    provider: MailProvider,
    args: Record<string, unknown>
  ) => Promise<{
    content: Array<{ type: "text"; text: string }>;
    isError?: boolean;
  }>;
}

export const getEmailTool: Tool = {
  name: "get_email",
  description: "Get full details of a specific email including body content.",
  inputSchema: {
    type: "object",
    properties: {
      emailId: {
        type: "string",
        description: "Email ID to retrieve",
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
    required: ["emailId"],
  },
  async handler(provider: MailProvider, args: Record<string, unknown>) {
    const emailId = args.emailId as string;

    try {
      const email = await provider.getEmail(emailId);

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(email, null, 2),
          },
        ],
      };
    } catch (err: any) {
      return {
        content: [
          {
            type: "text",
            text: `Error retrieving email: ${err.message}`,
          },
        ],
        isError: true,
      };
    }
  },
};