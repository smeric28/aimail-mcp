// Entity types and capability interfaces for the broader Microsoft 365 / Graph
// surface (OneDrive + SharePoint files, contacts, and To Do tasks). These are
// kept separate from the core MailProvider so providers can opt in to the
// capabilities they support. The Microsoft provider implements all of them;
// the Google provider implements only MailProvider.

export interface DriveItem {
  id: string;
  name: string;
  isFolder: boolean;
  size?: number;
  webUrl?: string;
  lastModified?: string;
  mimeType?: string;
  parentPath?: string;
  driveId?: string;
}

export interface SharePointSite {
  id: string;
  displayName: string;
  webUrl?: string;
}

export interface ShareLink {
  url: string;
  type: string;
  scope: string;
  expiresOn?: string;
}

export interface ContactCard {
  id: string;
  displayName?: string;
  givenName?: string;
  surname?: string;
  emails: string[];
  phones?: string[];
  company?: string;
  jobTitle?: string;
}

export interface TaskList {
  id: string;
  displayName: string;
  isDefault?: boolean;
}

export interface TaskItem {
  id: string;
  title: string;
  status: string; // notStarted | inProgress | completed | ...
  dueDateTime?: string;
  reminderDateTime?: string;
  body?: string;
  importance?: string;
  listId?: string;
}

export interface ChatSummary {
  id: string;
  topic?: string;
  chatType?: string; // oneOnOne | group | meeting
  members?: string[];
  lastUpdated?: string;
  webUrl?: string;
}

export interface ChatMessage {
  id: string;
  from?: string;
  createdDateTime?: string;
  content?: string;
  contentType?: string;
}

// --- Capability interfaces (Microsoft-only today) ---

export interface FilesCapable {
  // OneDrive / SharePoint document library operations.
  listFiles(opts?: { path?: string; driveId?: string; top?: number }): Promise<DriveItem[]>;
  searchFiles(query: string, opts?: { driveId?: string; top?: number }): Promise<DriveItem[]>;
  readFileContent(itemId: string, driveId?: string): Promise<{ name: string; contentType: string; contentBytes: string }>;
  uploadFile(opts: {
    name: string;
    contentBytes: string; // base64
    parentPath?: string;
    parentItemId?: string;
    driveId?: string;
  }): Promise<DriveItem>;
  createFolder(opts: { name: string; parentItemId?: string; driveId?: string }): Promise<DriveItem>;
  deleteFile(itemId: string, driveId?: string): Promise<void>;
  createShareLink(opts: {
    itemId: string;
    driveId?: string;
    type?: "view" | "edit";
    scope?: "anonymous" | "organization";
  }): Promise<ShareLink>;
  listSharePointSites(query?: string): Promise<SharePointSite[]>;
  listSiteDrives(siteId: string): Promise<Array<{ id: string; name: string }>>;
}

export interface ContactsCapable {
  listContacts(top?: number): Promise<ContactCard[]>;
  searchContacts(query: string, top?: number): Promise<ContactCard[]>;
  createContact(contact: {
    givenName?: string;
    surname?: string;
    displayName?: string;
    emails?: string[];
    phones?: string[];
    company?: string;
    jobTitle?: string;
  }): Promise<ContactCard>;
  updateContact(id: string, changes: Partial<{
    givenName: string;
    surname: string;
    displayName: string;
    emails: string[];
    phones: string[];
    company: string;
    jobTitle: string;
  }>): Promise<ContactCard>;
  deleteContact(id: string): Promise<void>;
}

export interface TasksCapable {
  listTaskLists(): Promise<TaskList[]>;
  listTasks(listId: string, opts?: { includeCompleted?: boolean; top?: number }): Promise<TaskItem[]>;
  createTask(listId: string, task: {
    title: string;
    body?: string;
    dueDateTime?: string;
    reminderDateTime?: string;
    importance?: "low" | "normal" | "high";
  }): Promise<TaskItem>;
  updateTask(listId: string, taskId: string, changes: Partial<{
    title: string;
    body: string;
    dueDateTime: string;
    status: string;
    importance: string;
  }>): Promise<TaskItem>;
  completeTask(listId: string, taskId: string): Promise<TaskItem>;
  deleteTask(listId: string, taskId: string): Promise<void>;
}

export interface TeamsCapable {
  listChats(top?: number): Promise<ChatSummary[]>;
  listChatMessages(chatId: string, top?: number): Promise<ChatMessage[]>;
  sendChatMessage(chatId: string, content: string, contentType?: "text" | "html"): Promise<ChatMessage>;
}

// Runtime capability guards used by the tool router so we can return a clean
// "not supported for this account" message instead of a crash.
export function hasFiles(p: unknown): p is FilesCapable {
  return typeof (p as FilesCapable)?.listFiles === "function";
}
export function hasContacts(p: unknown): p is ContactsCapable {
  return typeof (p as ContactsCapable)?.listContacts === "function";
}
export function hasTasks(p: unknown): p is TasksCapable {
  return typeof (p as TasksCapable)?.listTaskLists === "function";
}
export function hasTeams(p: unknown): p is TeamsCapable {
  return typeof (p as TeamsCapable)?.listChats === "function";
}
