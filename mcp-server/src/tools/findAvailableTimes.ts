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

export const findAvailableTimesTool: Tool = {
  name: "find_available_times",
  description:
    "Find meeting time slots that work for you and optional attendees using Microsoft Graph findMeetingTimes. " +
    "Returns up to 10 candidate slots ranked by confidence.",
  inputSchema: {
    type: "object",
    properties: {
      durationMinutes: { type: "number", description: "Meeting length in minutes" },
      windowStart: { type: "string", description: "ISO-8601 earliest start, e.g. 2026-04-14T09:00:00" },
      windowEnd: { type: "string", description: "ISO-8601 latest end" },
      attendees: {
        type: "array",
        items: { type: "string" },
        description: "Attendee emails whose free/busy should be checked",
      },
      timeZone: { type: "string", description: "Timezone for windowStart/windowEnd. Default UTC." },
      accountId: { type: "string", description: "Optional account ID" },
      mailbox: { type: "string", description: "Optional shared mailbox email" },
    },
    required: ["durationMinutes", "windowStart", "windowEnd"],
  },
  async handler(provider, args) {
    try {
      const slots = await provider.findAvailableTimes({
        durationMinutes: args.durationMinutes as number,
        windowStart: args.windowStart as string,
        windowEnd: args.windowEnd as string,
        attendees: args.attendees as string[] | undefined,
        timeZone: args.timeZone as string | undefined,
      });
      return {
        content: [
          { type: "text", text: JSON.stringify({ count: slots.length, slots }, null, 2) },
        ],
      };
    } catch (err: any) {
      return {
        content: [{ type: "text", text: `Error finding available times: ${err.message}` }],
        isError: true,
      };
    }
  },
};
