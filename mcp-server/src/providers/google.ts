import { gmail_v1 } from "googleapis";
import {
  MailProvider,
  MailMessage,
  MailFolder,
  MailAttachment,
} from "../types/mail.js";

export class GoogleProvider implements MailProvider {
  private gmail: gmail_v1.Gmail;
  private draftsLabelName: string;

  constructor(gmail: gmail_v1.Gmail, draftsLabelName: string) {
    this.gmail = gmail;
    this.draftsLabelName = draftsLabelName;
  }

  async searchEmails(
    query: string,
    folder: string = "INBOX",
    top: number = 10
  ): Promise<MailMessage[]> {
    const limit = Math.min(top, 50);
    
    // Map standard folders to Gmail labels
    let q = query;
    if (folder.toUpperCase() === "INBOX") q = `label:INBOX ${q}`;
    else if (folder.toUpperCase() === "SENTITEMS") q = `label:SENT ${q}`;
    else if (folder.toUpperCase() === "ARCHIVE") q = `-label:INBOX -label:TRASH -label:SPAM ${q}`;
    else q = `label:${folder} ${q}`;

    const res = await this.gmail.users.messages.list({
      userId: "me",
      q: q,
      maxResults: limit,
    });

    const messages = res.data.messages || [];
    const results: MailMessage[] = [];

    for (const msg of messages) {
      if (msg.id) {
        const fullMsg = await this.getEmail(msg.id);
        results.push(fullMsg);
      }
    }

    return results;
  }

  async getEmail(id: string): Promise<MailMessage> {
    const res = await this.gmail.users.messages.get({
      userId: "me",
      id: id,
      format: "full",
    });

    const msg = res.data;
    const headers = msg.payload?.headers || [];
    
    const subject = headers.find(h => h.name?.toLowerCase() === "subject")?.value || "No Subject";
    const from = headers.find(h => h.name?.toLowerCase() === "from")?.value || "Unknown";
    const date = headers.find(h => h.name?.toLowerCase() === "date")?.value || "";

    // Extract body and snippet
    let body = "";
    if (msg.payload?.parts) {
      // Simplistic body extraction
      const parts = msg.payload.parts;
      const textPart = parts.find(p => p.mimeType === "text/plain");
      const htmlPart = parts.find(p => p.mimeType === "text/html");
      
      if (htmlPart?.body?.data) {
        body = Buffer.from(htmlPart.body.data, "base64").toString();
      } else if (textPart?.body?.data) {
        body = Buffer.from(textPart.body.data, "base64").toString();
      }
    } else if (msg.payload?.body?.data) {
      body = Buffer.from(msg.payload.body.data, "base64").toString();
    }

    const hasAttachments = !!(msg.payload?.parts?.some(p => p.filename && p.filename.length > 0));

    return {
      id: msg.id!,
      subject,
      from,
      date: new Date(date).toLocaleString(),
      preview: msg.snippet || "",
      hasAttachments,
      body,
      bodyType: "HTML", // Assume HTML for now or check mimeType
    };
  }

  async createDraft(draft: {
    to: string;
    subject: string;
    body: string;
    bodyType?: "text" | "HTML";
    originalEmailId?: string;
  }): Promise<string> {
    // Gmail requires a raw MIME message for drafts
    const mimeMessage = this.buildMimeMessage(draft);
    const encodedMessage = Buffer.from(mimeMessage)
      .toString("base64")
      .replace(/\+/g, "-")
      .replace(/\//g, "_")
      .replace(/=+$/, "");

    const res = await this.gmail.users.drafts.create({
      userId: "me",
      requestBody: {
        message: {
          raw: encodedMessage,
          // We don't explicitly set the label here if we want it in the general Drafts, 
          // but the user asked for "aiDrafts" support.
        },
      },
    });

    // If we want to add a label to the message in the draft
    if (res.data.message?.id) {
        try {
            // Check if label exists, if so add it
            const labelsRes = await this.gmail.users.labels.list({ userId: "me" });
            const aiDraftsLabel = labelsRes.data.labels?.find(l => l.name === this.draftsLabelName);
            if (aiDraftsLabel) {
                await this.gmail.users.messages.batchModify({
                    userId: "me",
                    requestBody: {
                        ids: [res.data.message.id],
                        addLabelIds: [aiDraftsLabel.id!]
                    }
                });
            }
        } catch (err) {
            console.error("Failed to add AI drafts label to Gmail draft:", err);
        }
    }

    return res.data.id!;
  }

  private buildMimeMessage(draft: { to: string, subject: string, body: string, bodyType?: string, originalEmailId?: string }): string {
    const contentType = draft.bodyType === "HTML" ? "text/html" : "text/plain";
    const boundary = "boundary_" + Math.random().toString(36).substring(2);
    
    let message = [
      `To: ${draft.to}`,
      `Subject: ${draft.subject}`,
      `Content-Type: ${contentType}; charset=utf-8`,
      "MIME-Version: 1.0",
      "",
      draft.body
    ].join("\r\n");

    return message;
  }

  async updateDraft(
    id: string,
    draft: { subject?: string; body?: string; bodyType?: "text" | "HTML" }
  ): Promise<void> {
    // To update a draft in Gmail, you usually get the draft, then update its message
    const existingDraft = await this.gmail.users.drafts.get({ userId: "me", id });
    const headers = existingDraft.data.message?.payload?.headers || [];
    const to = headers.find(h => h.name === "To")?.value || "";
    
    const newDraftData = {
        to,
        subject: draft.subject || headers.find(h => h.name === "Subject")?.value || "",
        body: draft.body || "", // Ideally we'd extract the old body if not provided
        bodyType: draft.bodyType
    };

    const mimeMessage = this.buildMimeMessage(newDraftData);
    const encodedMessage = Buffer.from(mimeMessage)
      .toString("base64")
      .replace(/\+/g, "-")
      .replace(/\//g, "_")
      .replace(/=+$/, "");

    await this.gmail.users.drafts.update({
      userId: "me",
      id: id,
      requestBody: {
        message: {
          raw: encodedMessage,
        },
      },
    });
  }

  async deleteDraft(id: string): Promise<void> {
    await this.gmail.users.drafts.delete({ userId: "me", id });
  }

  async moveEmail(id: string, folderId: string): Promise<void> {
    // In Gmail, "moving" is adding a label and removing INBOX
    await this.gmail.users.messages.modify({
      userId: "me",
      id: id,
      requestBody: {
        addLabelIds: [folderId],
        removeLabelIds: ["INBOX"],
      },
    });
  }

  async listFolders(): Promise<MailFolder[]> {
    const res = await this.gmail.users.labels.list({ userId: "me" });
    return (res.data.labels || []).map(l => ({
      id: l.id!,
      displayName: l.name!,
    }));
  }

  async getAttachments(messageId: string): Promise<MailAttachment[]> {
    const res = await this.gmail.users.messages.get({ userId: "me", id: messageId });
    const parts = res.data.payload?.parts || [];
    const attachments: MailAttachment[] = [];

    for (const part of parts) {
      if (part.filename && part.filename.length > 0) {
        attachments.push({
          id: part.body?.attachmentId || part.partId!,
          name: part.filename,
          contentType: part.mimeType!,
          size: part.body?.size || 0,
        });
      }
    }
    return attachments;
  }

  async getAttachmentContent(
    messageId: string,
    attachmentId: string
  ): Promise<string> {
    const res = await this.gmail.users.messages.attachments.get({
      userId: "me",
      messageId: messageId,
      id: attachmentId,
    });
    return res.data.data!; // This is base64url encoded
  }

  async addAttachmentToDraft(
    draftId: string,
    attachment: { name: string; contentType: string; contentBytes: string }
  ): Promise<void> {
    // Gmail drafts are immutable once created, you have to recreate them with attachments.
    // For simplicity, this is often handled by building the MIME message with everything initially.
    // Implement if needed, but for now we'll throw a placeholder error or skip.
    throw new Error("Add attachment to draft not yet implemented for Gmail in this MVP.");
  }
}
