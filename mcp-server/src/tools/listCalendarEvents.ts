import { MailProvider } from "../types/mail.js";

interface Tool {
  name: string;
  description: string;
  inputSchema: {
    type: "object";
    properties: Record<string, any>;
    required: string[];
  };
  handler: (
    provider: MailProvider,
    args: Record<string, unknown>
  ) => Promise<{
    content: Array<{ type: "text"; text: string }>;
    isError?: boolean;
  }>;
}

export const listCalendarEventsTool: Tool = {
  name: "list_calendar_events",
  description:
    "List calendar events in a time window. start/end must be ISO-8601 datetimes (UTC). " +
    "Returns subject, times, attendees, location, and webLink.",
  inputSchema: {
    type: "object",
    properties: {
      start: { type: "string", description: "ISO-8601 start of window (UTC), e.g. 2026-04-13T00:00:00Z" },
      end: { type: "string", description: "ISO-8601 end of window (UTC)" },
      top: { type: "number", description: "Max events to return (default 50, max 100)" },
      accountId: { type: "string", description: "Optional account ID" },
      mailbox: { type: "string", description: "Optional shared mailbox email" },
    },
    required: ["start", "end"],
  },
  async handler(provider, args) {
    try {
      const events = await provider.listCalendarEvents(
        args.start as string,
        args.end as string,
        args.top as number | undefined
      );
      return {
        content: [
          { type: "text", text: JSON.stringify({ count: events.length, events }, null, 2) },
        ],
      };
    } catch (err: any) {
      return {
        content: [{ type: "text", text: `Error listing calendar events: ${err.message}` }],
        isError: true,
      };
    }
  },
};
