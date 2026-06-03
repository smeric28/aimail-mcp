import { RegisteredTool, ok, fail, accountFields } from "./toolTypes.js";
import { hasTeams } from "../types/graph.js";

export const teamsTools: RegisteredTool[] = [
  {
    name: "list_chats",
    description: "List the signed-in user's recent Microsoft Teams chats (1:1 and group).",
    inputSchema: {
      type: "object",
      properties: {
        top: { type: "number", description: "Max chats (default 25)" },
        ...accountFields,
      },
    },
    async run(provider, args) {
      if (!hasTeams(provider)) return fail("Teams is not supported for this account.");
      return ok(await provider.listChats(Number(args.top) || undefined));
    },
  },
  {
    name: "list_chat_messages",
    description: "List recent messages in a Microsoft Teams chat.",
    inputSchema: {
      type: "object",
      properties: {
        chatId: { type: "string", description: "Chat ID (from list_chats)" },
        top: { type: "number", description: "Max messages (default 25)" },
        ...accountFields,
      },
      required: ["chatId"],
    },
    async run(provider, args) {
      if (!hasTeams(provider)) return fail("Teams is not supported for this account.");
      return ok(await provider.listChatMessages(args.chatId as string, Number(args.top) || undefined));
    },
  },
  {
    name: "send_chat_message",
    description: "Send a message to a Microsoft Teams chat.",
    inputSchema: {
      type: "object",
      properties: {
        chatId: { type: "string", description: "Chat ID (from list_chats)" },
        content: { type: "string", description: "Message text" },
        contentType: { type: "string", enum: ["text", "html"], description: "Message format (default text)" },
        ...accountFields,
      },
      required: ["chatId", "content"],
    },
    async run(provider, args) {
      if (!hasTeams(provider)) return fail("Teams is not supported for this account.");
      const message = await provider.sendChatMessage(
        args.chatId as string,
        args.content as string,
        (args.contentType as "text" | "html" | undefined) || "text"
      );
      return ok({ success: true, message });
    },
  },
];
