/**
 * Microsoft Graph delegated scopes for the full Microsoft 365 connector surface
 * (mail, calendar, contacts, OneDrive/SharePoint files, and To Do tasks) with
 * read AND write access. These are the scopes the Entra app registration must
 * be granted (with admin consent for the Fireball tenant), and the scopes the
 * OAuth bridge requests on behalf of each user.
 *
 * `offline_access`, `openid`, and `profile` are added automatically by the
 * OAuth bridge and are not listed here.
 */
export const GRAPH_DELEGATED_SCOPES = [
  "User.Read",

  // Mail (read/write, send, shared mailboxes)
  "Mail.ReadWrite",
  "Mail.ReadWrite.Shared",
  "Mail.Send",
  "Mail.Send.Shared",

  // Calendar (read/write, shared calendars, invites)
  "Calendars.ReadWrite",
  "Calendars.ReadWrite.Shared",

  // Contacts
  "Contacts.ReadWrite",

  // Files: OneDrive + SharePoint document libraries
  "Files.ReadWrite.All",
  "Sites.ReadWrite.All",

  // Microsoft To Do tasks
  "Tasks.ReadWrite",

  // Teams chat (read + send)
  "Chat.ReadWrite",

  // Directory lookups for find-times / people
  "People.Read",
];

/** Scopes formatted as fully-qualified Graph resource URIs (for MSAL/.default-style requests). */
export function graphScopeUris(): string[] {
  return GRAPH_DELEGATED_SCOPES.map((s) =>
    s.includes("/") ? s : `https://graph.microsoft.com/${s}`
  );
}
