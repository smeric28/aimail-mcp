import { RegisteredTool, ok, fail, accountFields } from "./toolTypes.js";

export const mailActionTools: RegisteredTool[] = [
  {
    name: "send_email",
    description:
      "Compose and SEND an email immediately (not a draft). Use create_draft if the user wants to review first.",
    inputSchema: {
      type: "object",
      properties: {
        to: { type: "string", description: "Recipient(s), comma-separated" },
        cc: { type: "string", description: "CC recipient(s), comma-separated" },
        bcc: { type: "string", description: "BCC recipient(s), comma-separated" },
        subject: { type: "string", description: "Subject line" },
        body: { type: "string", description: "Message body" },
        bodyType: { type: "string", enum: ["text", "HTML"], description: "Body format (default HTML)" },
        ...accountFields,
      },
      required: ["to", "subject", "body"],
    },
    async run(provider, args) {
      if (!provider.sendEmail) return fail("Sending mail is not supported for this account.");
      await provider.sendEmail({
        to: args.to as string,
        cc: args.cc as string | undefined,
        bcc: args.bcc as string | undefined,
        subject: args.subject as string,
        body: args.body as string,
        bodyType: args.bodyType as "text" | "HTML" | undefined,
      });
      return ok({ success: true, message: "Email sent." });
    },
  },
  {
    name: "send_draft",
    description: "Send an existing draft message by its ID.",
    inputSchema: {
      type: "object",
      properties: {
        draftId: { type: "string", description: "ID of the draft to send" },
        ...accountFields,
      },
      required: ["draftId"],
    },
    async run(provider, args) {
      if (!provider.sendDraft) return fail("Sending drafts is not supported for this account.");
      await provider.sendDraft(args.draftId as string);
      return ok({ success: true, message: "Draft sent." });
    },
  },
  {
    name: "reply_to_email",
    description:
      "Reply (or reply-all) to an email. Set send=true to send immediately, otherwise a draft reply is created and its ID returned.",
    inputSchema: {
      type: "object",
      properties: {
        id: { type: "string", description: "ID of the message to reply to" },
        body: { type: "string", description: "Reply text" },
        bodyType: { type: "string", enum: ["text", "HTML"] },
        replyAll: { type: "boolean", description: "Reply to all recipients" },
        send: { type: "boolean", description: "Send now instead of creating a draft" },
        ...accountFields,
      },
      required: ["id", "body"],
    },
    async run(provider, args) {
      if (!provider.replyToEmail) return fail("Reply is not supported for this account.");
      const draftId = await provider.replyToEmail(args.id as string, {
        body: args.body as string,
        bodyType: args.bodyType as "text" | "HTML" | undefined,
        replyAll: args.replyAll as boolean | undefined,
        send: args.send as boolean | undefined,
      });
      return ok({ success: true, sent: !!args.send, draftId: draftId || undefined });
    },
  },
  {
    name: "forward_email",
    description:
      "Forward an email to new recipients. Set send=true to send immediately, otherwise a draft is created.",
    inputSchema: {
      type: "object",
      properties: {
        id: { type: "string", description: "ID of the message to forward" },
        to: { type: "string", description: "Recipient(s), comma-separated" },
        comment: { type: "string", description: "Optional note added above the forwarded message" },
        send: { type: "boolean", description: "Send now instead of creating a draft" },
        ...accountFields,
      },
      required: ["id", "to"],
    },
    async run(provider, args) {
      if (!provider.forwardEmail) return fail("Forward is not supported for this account.");
      const draftId = await provider.forwardEmail(args.id as string, {
        to: args.to as string,
        comment: args.comment as string | undefined,
        send: args.send as boolean | undefined,
      });
      return ok({ success: true, sent: !!args.send, draftId: draftId || undefined });
    },
  },
  {
    name: "set_email_read",
    description: "Mark an email as read or unread.",
    inputSchema: {
      type: "object",
      properties: {
        id: { type: "string", description: "Message ID" },
        isRead: { type: "boolean", description: "true = read, false = unread" },
        ...accountFields,
      },
      required: ["id", "isRead"],
    },
    async run(provider, args) {
      if (!provider.markEmailRead) return fail("This action is not supported for this account.");
      await provider.markEmailRead(args.id as string, args.isRead as boolean);
      return ok({ success: true });
    },
  },
  {
    name: "flag_email",
    description: "Flag or unflag an email for follow-up.",
    inputSchema: {
      type: "object",
      properties: {
        id: { type: "string", description: "Message ID" },
        flagged: { type: "boolean", description: "true = flag, false = clear flag" },
        ...accountFields,
      },
      required: ["id", "flagged"],
    },
    async run(provider, args) {
      if (!provider.flagEmail) return fail("This action is not supported for this account.");
      await provider.flagEmail(args.id as string, args.flagged as boolean);
      return ok({ success: true });
    },
  },
];
