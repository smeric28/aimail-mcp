import { RegisteredTool, ok, fail, accountFields } from "./toolTypes.js";
import { hasContacts } from "../types/graph.js";

export const contactTools: RegisteredTool[] = [
  {
    name: "list_contacts",
    description: "List the user's Outlook contacts.",
    inputSchema: {
      type: "object",
      properties: {
        top: { type: "number", description: "Max results (default 50)" },
        ...accountFields,
      },
    },
    async run(provider, args) {
      if (!hasContacts(provider)) return fail("Contacts are not supported for this account.");
      return ok(await provider.listContacts(Number(args.top) || undefined));
    },
  },
  {
    name: "search_contacts",
    description: "Search Outlook contacts by name.",
    inputSchema: {
      type: "object",
      properties: {
        query: { type: "string", description: "Name to search for" },
        top: { type: "number", description: "Max results (default 25)" },
        ...accountFields,
      },
      required: ["query"],
    },
    async run(provider, args) {
      if (!hasContacts(provider)) return fail("Contacts are not supported for this account.");
      return ok(await provider.searchContacts(args.query as string, Number(args.top) || undefined));
    },
  },
  {
    name: "create_contact",
    description: "Create a new Outlook contact.",
    inputSchema: {
      type: "object",
      properties: {
        givenName: { type: "string" },
        surname: { type: "string" },
        displayName: { type: "string" },
        emails: { type: "array", items: { type: "string" }, description: "Email addresses" },
        phones: { type: "array", items: { type: "string" }, description: "Phone numbers" },
        company: { type: "string" },
        jobTitle: { type: "string" },
        ...accountFields,
      },
    },
    async run(provider, args) {
      if (!hasContacts(provider)) return fail("Contacts are not supported for this account.");
      const contact = await provider.createContact({
        givenName: args.givenName as string | undefined,
        surname: args.surname as string | undefined,
        displayName: args.displayName as string | undefined,
        emails: args.emails as string[] | undefined,
        phones: args.phones as string[] | undefined,
        company: args.company as string | undefined,
        jobTitle: args.jobTitle as string | undefined,
      });
      return ok({ success: true, contact });
    },
  },
  {
    name: "update_contact",
    description: "Update fields on an existing Outlook contact.",
    inputSchema: {
      type: "object",
      properties: {
        id: { type: "string", description: "Contact ID" },
        givenName: { type: "string" },
        surname: { type: "string" },
        displayName: { type: "string" },
        emails: { type: "array", items: { type: "string" } },
        phones: { type: "array", items: { type: "string" } },
        company: { type: "string" },
        jobTitle: { type: "string" },
        ...accountFields,
      },
      required: ["id"],
    },
    async run(provider, args) {
      if (!hasContacts(provider)) return fail("Contacts are not supported for this account.");
      const { id, accountId, mailbox, ...changes } = args as Record<string, unknown>;
      const contact = await provider.updateContact(id as string, changes as any);
      return ok({ success: true, contact });
    },
  },
  {
    name: "delete_contact",
    description: "Delete an Outlook contact.",
    inputSchema: {
      type: "object",
      properties: {
        id: { type: "string", description: "Contact ID" },
        ...accountFields,
      },
      required: ["id"],
    },
    async run(provider, args) {
      if (!hasContacts(provider)) return fail("Contacts are not supported for this account.");
      await provider.deleteContact(args.id as string);
      return ok({ success: true });
    },
  },
];
