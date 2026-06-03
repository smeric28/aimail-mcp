import { RegisteredTool, ok, fail, accountFields } from "./toolTypes.js";

export const calendarActionTools: RegisteredTool[] = [
  {
    name: "update_calendar_event",
    description:
      "Update an existing calendar event (time, subject, location, body, or attendees). Updates are sent to attendees.",
    inputSchema: {
      type: "object",
      properties: {
        id: { type: "string", description: "Event ID" },
        subject: { type: "string" },
        start: { type: "string", description: "ISO-8601 start datetime" },
        end: { type: "string", description: "ISO-8601 end datetime" },
        timeZone: { type: "string", description: "IANA/Windows timezone (default UTC)" },
        location: { type: "string" },
        body: { type: "string" },
        bodyType: { type: "string", enum: ["text", "HTML"] },
        attendees: { type: "array", items: { type: "string" }, description: "Replaces the attendee list" },
        ...accountFields,
      },
      required: ["id"],
    },
    async run(provider, args) {
      if (!provider.updateCalendarEvent) return fail("Calendar updates are not supported for this account.");
      const event = await provider.updateCalendarEvent(args.id as string, {
        subject: args.subject as string | undefined,
        start: args.start as string | undefined,
        end: args.end as string | undefined,
        timeZone: args.timeZone as string | undefined,
        location: args.location as string | undefined,
        body: args.body as string | undefined,
        bodyType: args.bodyType as "text" | "HTML" | undefined,
        attendees: args.attendees as string[] | undefined,
      });
      return ok({ success: true, event });
    },
  },
  {
    name: "cancel_calendar_event",
    description: "Cancel an event you organize and notify attendees. Use delete_calendar_event to remove without notifying.",
    inputSchema: {
      type: "object",
      properties: {
        id: { type: "string", description: "Event ID" },
        comment: { type: "string", description: "Optional cancellation note to attendees" },
        ...accountFields,
      },
      required: ["id"],
    },
    async run(provider, args) {
      if (!provider.cancelCalendarEvent) return fail("Cancel is not supported for this account.");
      await provider.cancelCalendarEvent(args.id as string, args.comment as string | undefined);
      return ok({ success: true, message: "Event cancelled and attendees notified." });
    },
  },
  {
    name: "delete_calendar_event",
    description: "Delete an event from the calendar (no cancellation notice sent).",
    inputSchema: {
      type: "object",
      properties: {
        id: { type: "string", description: "Event ID" },
        ...accountFields,
      },
      required: ["id"],
    },
    async run(provider, args) {
      if (!provider.deleteCalendarEvent) return fail("Delete is not supported for this account.");
      await provider.deleteCalendarEvent(args.id as string);
      return ok({ success: true });
    },
  },
  {
    name: "respond_to_calendar_event",
    description: "Accept, decline, or tentatively accept a meeting invitation you received.",
    inputSchema: {
      type: "object",
      properties: {
        id: { type: "string", description: "Event ID" },
        response: {
          type: "string",
          enum: ["accept", "decline", "tentativelyAccept"],
          description: "Your response",
        },
        comment: { type: "string", description: "Optional message to the organizer" },
        ...accountFields,
      },
      required: ["id", "response"],
    },
    async run(provider, args) {
      if (!provider.respondToCalendarEvent) return fail("Responding is not supported for this account.");
      await provider.respondToCalendarEvent(
        args.id as string,
        args.response as "accept" | "decline" | "tentativelyAccept",
        args.comment as string | undefined
      );
      return ok({ success: true, response: args.response });
    },
  },
];
