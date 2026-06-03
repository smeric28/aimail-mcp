export interface MailMessage {
  id: string;
  subject: string;
  from: string;
  date: string;
  preview: string;
  hasAttachments: boolean;
  body?: string;
  bodyType?: "text" | "HTML";
}

export interface MailAttachment {
  id: string;
  name: string;
  contentType: string;
  size: number;
  contentBytes?: string; // base64
}

export interface MailFolder {
  id: string;
  displayName: string;
}

export interface CalendarEvent {
  id: string;
  subject: string;
  start: string;
  end: string;
  organizer?: string;
  attendees?: Array<{ email: string; name?: string; response?: string }>;
  location?: string;
  body?: string;
  bodyType?: "text" | "HTML";
  webLink?: string;
  isOnlineMeeting?: boolean;
  onlineMeetingUrl?: string;
}

export interface TimeSlot {
  start: string;
  end: string;
  confidence?: number;
}

export interface MailProvider {
  searchEmails(query: string, folder?: string, top?: number): Promise<MailMessage[]>;
  getEmail(id: string): Promise<MailMessage>;
  createDraft(draft: {
    to: string;
    subject: string;
    body: string;
    bodyType?: "text" | "HTML";
    originalEmailId?: string;
  }): Promise<string>;
  updateDraft(id: string, draft: {
    subject?: string;
    body?: string;
    bodyType?: "text" | "HTML";
  }): Promise<void>;
  deleteDraft(id: string): Promise<void>;
  moveEmail(id: string, folderId: string): Promise<void>;
  listFolders(): Promise<MailFolder[]>;
  getAttachments(messageId: string): Promise<MailAttachment[]>;
  getAttachmentContent(messageId: string, attachmentId: string): Promise<string>;
  addAttachmentToDraft(draftId: string, attachment: {
    name: string;
    contentType: string;
    contentBytes: string;
  }): Promise<void>;

  listCalendarEvents(start: string, end: string, top?: number): Promise<CalendarEvent[]>;
  createCalendarEvent(event: {
    subject: string;
    start: string;
    end: string;
    timeZone?: string;
    attendees?: string[];
    body?: string;
    bodyType?: "text" | "HTML";
    location?: string;
    isOnlineMeeting?: boolean;
  }): Promise<CalendarEvent>;
  findAvailableTimes(params: {
    durationMinutes: number;
    windowStart: string;
    windowEnd: string;
    attendees?: string[];
    timeZone?: string;
  }): Promise<TimeSlot[]>;

  // --- Optional mail write actions (implemented by Microsoft; Gmail may omit) ---
  sendEmail?(message: {
    to: string;
    cc?: string;
    bcc?: string;
    subject: string;
    body: string;
    bodyType?: "text" | "HTML";
    saveToSentItems?: boolean;
  }): Promise<void>;
  sendDraft?(draftId: string): Promise<void>;
  replyToEmail?(id: string, opts: {
    body: string;
    bodyType?: "text" | "HTML";
    replyAll?: boolean;
    send?: boolean;
  }): Promise<string | void>;
  forwardEmail?(id: string, opts: {
    to: string;
    comment?: string;
    send?: boolean;
  }): Promise<string | void>;
  markEmailRead?(id: string, isRead: boolean): Promise<void>;
  flagEmail?(id: string, flagged: boolean): Promise<void>;

  // --- Optional calendar write actions ---
  updateCalendarEvent?(id: string, changes: {
    subject?: string;
    start?: string;
    end?: string;
    timeZone?: string;
    location?: string;
    body?: string;
    bodyType?: "text" | "HTML";
    attendees?: string[];
  }): Promise<CalendarEvent>;
  cancelCalendarEvent?(id: string, comment?: string): Promise<void>;
  deleteCalendarEvent?(id: string): Promise<void>;
  respondToCalendarEvent?(id: string, response: "accept" | "decline" | "tentativelyAccept", comment?: string): Promise<void>;
  getCalendarEvent?(id: string): Promise<CalendarEvent>;
}

export type ProviderType = "outlook" | "gmail";

export interface AccountInfo {
  id: string;
  name: string;
  email: string;
  type: ProviderType;
}
