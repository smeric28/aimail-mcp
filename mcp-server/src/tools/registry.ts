import { MailProvider } from "../types/mail.js";
import { RegisteredTool, ToolContext } from "./toolTypes.js";

// Existing single-purpose tools (one object per file).
import { searchEmailsTool } from "./searchEmails.js";
import { getEmailTool } from "./getEmail.js";
import { createDraftTool } from "./createDraft.js";
import { updateDraftTool } from "./updateDraft.js";
import { deleteDraftTool } from "./deleteDraft.js";
import { moveEmailTool } from "./moveEmail.js";
import { listFoldersTool } from "./listFolders.js";
import { findAttachmentsTool } from "./findAttachments.js";
import { copyAttachmentTool } from "./copyAttachment.js";
import { addAttachmentToDraftTool } from "./addAttachmentToDraft.js";
import { createCalendarEventTool } from "./createCalendarEvent.js";
import { listCalendarEventsTool } from "./listCalendarEvents.js";
import { findAvailableTimesTool } from "./findAvailableTimes.js";

// New grouped tool modules.
import { mailActionTools } from "./mailActions.js";
import { calendarActionTools } from "./calendarActions.js";
import { fileTools } from "./files.js";
import { contactTools } from "./contacts.js";
import { taskTools } from "./tasks.js";
import { teamsTools } from "./teams.js";

// Adapt a legacy tool object (with a `handler`) into a RegisteredTool. The
// handlers have slightly different arities; the cast lets us call them all the
// same way, passing the context as the optional third argument.
function adapt(tool: {
  name: string;
  description: string;
  inputSchema: any;
  handler: (...a: any[]) => Promise<any>;
}): RegisteredTool {
  return {
    name: tool.name,
    description: tool.description,
    inputSchema: tool.inputSchema,
    run: (provider: MailProvider, args: Record<string, unknown>, ctx: ToolContext) =>
      tool.handler(provider, args, ctx),
  };
}

const legacyTools: RegisteredTool[] = [
  searchEmailsTool,
  getEmailTool,
  createDraftTool,
  updateDraftTool,
  deleteDraftTool,
  moveEmailTool,
  listFoldersTool,
  findAttachmentsTool,
  copyAttachmentTool,
  addAttachmentToDraftTool,
  createCalendarEventTool,
  listCalendarEventsTool,
  findAvailableTimesTool,
].map(adapt);

export const allTools: RegisteredTool[] = [
  ...legacyTools,
  ...mailActionTools,
  ...calendarActionTools,
  ...fileTools,
  ...contactTools,
  ...taskTools,
  ...teamsTools,
];

export const toolsByName: Map<string, RegisteredTool> = new Map(
  allTools.map((t) => [t.name, t])
);

// MCP tool list payload (name/description/inputSchema only).
export function toolDefinitions() {
  return allTools.map((t) => ({
    name: t.name,
    description: t.description,
    inputSchema: t.inputSchema,
  }));
}
