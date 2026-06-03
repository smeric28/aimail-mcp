import { RegisteredTool, ok, fail, accountFields } from "./toolTypes.js";
import { hasFiles } from "../types/graph.js";

const driveField = {
  driveId: {
    type: "string",
    description: "Optional drive ID (a SharePoint document library). Omit for the user's OneDrive.",
  },
};

export const fileTools: RegisteredTool[] = [
  {
    name: "list_files",
    description: "List files and folders in OneDrive or a SharePoint document library at an optional path.",
    inputSchema: {
      type: "object",
      properties: {
        path: { type: "string", description: "Folder path relative to the drive root, e.g. 'Documents/Reports'. Omit for root." },
        top: { type: "number", description: "Max items (default 50)" },
        ...driveField,
        ...accountFields,
      },
    },
    async run(provider, args) {
      if (!hasFiles(provider)) return fail("File access is not supported for this account.");
      const items = await provider.listFiles({
        path: args.path as string | undefined,
        driveId: args.driveId as string | undefined,
        top: Number(args.top) || undefined,
      });
      return ok(items);
    },
  },
  {
    name: "search_files",
    description: "Search OneDrive / SharePoint for files and folders matching a query.",
    inputSchema: {
      type: "object",
      properties: {
        query: { type: "string", description: "Search text" },
        top: { type: "number", description: "Max results (default 25)" },
        ...driveField,
        ...accountFields,
      },
      required: ["query"],
    },
    async run(provider, args) {
      if (!hasFiles(provider)) return fail("File access is not supported for this account.");
      const items = await provider.searchFiles(args.query as string, {
        driveId: args.driveId as string | undefined,
        top: Number(args.top) || undefined,
      });
      return ok(items);
    },
  },
  {
    name: "read_file",
    description: "Read a file's content from OneDrive / SharePoint. Returns base64-encoded bytes plus metadata.",
    inputSchema: {
      type: "object",
      properties: {
        itemId: { type: "string", description: "Drive item ID (from list_files or search_files)" },
        ...driveField,
        ...accountFields,
      },
      required: ["itemId"],
    },
    async run(provider, args) {
      if (!hasFiles(provider)) return fail("File access is not supported for this account.");
      const file = await provider.readFileContent(args.itemId as string, args.driveId as string | undefined);
      return ok(file);
    },
  },
  {
    name: "upload_file",
    description: "Upload (create or overwrite) a file in OneDrive / SharePoint. Content must be base64-encoded. For files under 4 MB.",
    inputSchema: {
      type: "object",
      properties: {
        name: { type: "string", description: "File name including extension" },
        contentBytes: { type: "string", description: "Base64-encoded file content" },
        parentPath: { type: "string", description: "Destination folder path relative to root (optional)" },
        parentItemId: { type: "string", description: "Destination folder item ID (optional, alternative to parentPath)" },
        ...driveField,
        ...accountFields,
      },
      required: ["name", "contentBytes"],
    },
    async run(provider, args) {
      if (!hasFiles(provider)) return fail("File access is not supported for this account.");
      const item = await provider.uploadFile({
        name: args.name as string,
        contentBytes: args.contentBytes as string,
        parentPath: args.parentPath as string | undefined,
        parentItemId: args.parentItemId as string | undefined,
        driveId: args.driveId as string | undefined,
      });
      return ok({ success: true, item });
    },
  },
  {
    name: "create_folder",
    description: "Create a new folder in OneDrive / SharePoint.",
    inputSchema: {
      type: "object",
      properties: {
        name: { type: "string", description: "New folder name" },
        parentItemId: { type: "string", description: "Parent folder item ID (optional, defaults to root)" },
        ...driveField,
        ...accountFields,
      },
      required: ["name"],
    },
    async run(provider, args) {
      if (!hasFiles(provider)) return fail("File access is not supported for this account.");
      const item = await provider.createFolder({
        name: args.name as string,
        parentItemId: args.parentItemId as string | undefined,
        driveId: args.driveId as string | undefined,
      });
      return ok({ success: true, item });
    },
  },
  {
    name: "delete_file",
    description: "Delete a file or folder from OneDrive / SharePoint.",
    inputSchema: {
      type: "object",
      properties: {
        itemId: { type: "string", description: "Drive item ID" },
        ...driveField,
        ...accountFields,
      },
      required: ["itemId"],
    },
    async run(provider, args) {
      if (!hasFiles(provider)) return fail("File access is not supported for this account.");
      await provider.deleteFile(args.itemId as string, args.driveId as string | undefined);
      return ok({ success: true });
    },
  },
  {
    name: "share_file",
    description: "Create a sharing link for a file or folder. Scope 'organization' (default) or 'anonymous'; type 'view' or 'edit'.",
    inputSchema: {
      type: "object",
      properties: {
        itemId: { type: "string", description: "Drive item ID" },
        type: { type: "string", enum: ["view", "edit"], description: "Link permission (default view)" },
        scope: { type: "string", enum: ["anonymous", "organization"], description: "Link audience (default organization)" },
        ...driveField,
        ...accountFields,
      },
      required: ["itemId"],
    },
    async run(provider, args) {
      if (!hasFiles(provider)) return fail("File access is not supported for this account.");
      const link = await provider.createShareLink({
        itemId: args.itemId as string,
        driveId: args.driveId as string | undefined,
        type: args.type as "view" | "edit" | undefined,
        scope: args.scope as "anonymous" | "organization" | undefined,
      });
      return ok(link);
    },
  },
  {
    name: "list_sharepoint_sites",
    description: "Find SharePoint sites in the organization. Use a site's drives to read/write its document libraries.",
    inputSchema: {
      type: "object",
      properties: {
        query: { type: "string", description: "Optional search text; omit to list available sites" },
        ...accountFields,
      },
    },
    async run(provider, args) {
      if (!hasFiles(provider)) return fail("SharePoint access is not supported for this account.");
      const sites = await provider.listSharePointSites(args.query as string | undefined);
      return ok(sites);
    },
  },
  {
    name: "list_site_drives",
    description: "List the document libraries (drives) of a SharePoint site. Pass a returned drive ID to the file tools.",
    inputSchema: {
      type: "object",
      properties: {
        siteId: { type: "string", description: "SharePoint site ID (from list_sharepoint_sites)" },
        ...accountFields,
      },
      required: ["siteId"],
    },
    async run(provider, args) {
      if (!hasFiles(provider)) return fail("SharePoint access is not supported for this account.");
      const drives = await provider.listSiteDrives(args.siteId as string);
      return ok(drives);
    },
  },
];
