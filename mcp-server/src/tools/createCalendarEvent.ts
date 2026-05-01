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

export const createCalendarEventTool: Tool = {
  name: "create_calendar_event",
  description:
    "Create a calendar event / meeting invite. Sends invitations to attendees immediately. " +
    "start/end must be ISO-8601 local datetime (e.g. '2026-04-15T14:00:00'). Supply timeZone if not UTC.",
  inputSchema: {
    type: "object",
    properties: {
      subject: { type: "string", description: "Event title" },
      start: { type: "string", description: "ISO-8601 start datetime, e.g. 2026-04-15T14:00:00" },
      end: { type: "string", description: "ISO-8601 end datetime" },
      timeZone: { type: "string", description: "IANA/Windows timezone, e.g. 'Pacific Standard Time' or 'UTC'. Defaults to UTC." },
      attendees: {
        type: "array",
        items: { type: "string" },
        description: "Attendee email addresses",
      },
      body: { type: "string", description: "Event description / agenda" },
      bodyType: { type: "string", enum: ["text", "HTML"], description: "Body format (default HTML)" },
      location: { type: "string", description: "Physical location (optional)" },
      isOnlineMeeting: { type: "boolean", description: "Add a Teams meeting link" },
      accountId: { type: "string", description: "Optional account ID" },
      mailbox: { type: "string", description: "Optional shared mailbox email" },
    },
    required: ["subject", "start", "end"],
  },
  async handler(provider, args) {
    try {
      const event = await provider.createCalendarEvent({
        subject: args.subject as string,
        start: args.start as string,
        end: args.end as string,
        timeZone: args.timeZone as string | undefined,
        attendees: args.attendees as string[] | undefined,
        body: args.body as string | undefined,
        bodyType: args.bodyType as "text" | "HTML" | undefined,
        location: args.location as string | undefined,
        isOnlineMeeting: args.isOnlineMeeting as boolean | undefined,
      });

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify({ success: true, event }, null, 2),
          },
        ],
      };
    } catch (err: any) {
      return {
        content: [{ type: "text", text: `Error creating calendar event: ${err.message}` }],
        isError: true,
      };
    }
  },
};
