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
}

export type ProviderType = "outlook" | "gmail";

export interface AccountInfo {
  id: string;
  name: string;
  email: string;
  type: ProviderType;
}
