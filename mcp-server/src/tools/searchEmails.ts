import { MailProvider } from "../types/mail.js";

interface Tool {
  name: string;
  description: string;
  inputSchema: {
    type: "object";
    properties: {
      query: { type: "string"; description: "Search query" };
      folder: { type: "string"; description: string };
      top: { type: "number"; description: "Max results (default: 10)" };
      accountId: { type: "string"; description: "Optional account ID to use" };
    };
    required: string[];
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

export const searchEmailsTool: Tool = {
  name: "search_emails",
  description:
    "Search emails using natural language or provider-specific filters. " +
    "Read-only operation - returns matching emails without modification.",
  inputSchema: {
    type: "object",
    properties: {
      query: {
        type: "string",
        description: "Search query",
      },
      folder: {
        type: "string",
        description: "Folder to search (default: inbox or INBOX)",
      },
      top: {
        type: "number",
        description: "Max results (default: 10)",
      },
      accountId: {
        type: "string",
        description: "Optional account ID to use",
      },
    },
    required: ["query"],
  },
  async handler(provider: MailProvider, args: Record<string, unknown>, config: { draftsFolderName: string }) {
    const query = args.query as string;
    const folder = (args.folder as string) || "inbox";
    const top = Math.min(Number(args.top) || 10, 50);

    try {
      const results = await provider.searchEmails(query, folder, top);

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(results, null, 2),
          },
        ],
      };
    } catch (err: any) {
      return {
        content: [
          {
            type: "text",
            text: `Error searching emails: ${err.message}`,
          },
        ],
        isError: true,
      };
    }
  },
};

export function parseSearchQuery(query: string): { type: 'filter' | 'search', value: string } {
  const lowerQuery = query.toLowerCase();
  const parts = lowerQuery.split(" ");

  if (lowerQuery.includes("is:unread") || (parts.includes("unread") && !parts.includes("week"))) {
    return { type: 'filter', value: "isRead eq false" };
  }

  if (parts.includes("unread") && parts.includes("week")) {
    const weekAgo = new Date();
    weekAgo.setDate(weekAgo.getDate() - 7);
    return { type: 'filter', value: `receivedDateTime gt ${weekAgo.toISOString()} and isRead eq false` };
  }

  if (parts.includes("attachment")) {
    return { type: 'filter', value: "hasAttachments eq true" };
  }

  // Use KQL for text searches
  if (lowerQuery.includes("from:")) {
    return { type: 'search', value: query };
  }

  if (parts.includes("from")) {
    const fromIdx = parts.indexOf("from") + 1;
    if (fromIdx < parts.length) {
      const name = parts.slice(fromIdx).join(" ");
      return { type: 'search', value: `from:${name}` };
    }
  }

  // Fallback: KQL Search for general text
  return { type: 'search', value: query };
}