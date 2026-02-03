import { Client } from "@microsoft/microsoft-graph-client";
import {
  MailProvider,
  MailMessage,
  MailFolder,
  MailAttachment,
} from "../types/mail.js";

export class MicrosoftProvider implements MailProvider {
  private client: Client;
  private draftsFolderName: string;
  private userPrefix: string;

  constructor(client: Client, draftsFolderName: string, userEmail?: string) {
    this.client = client;
    this.draftsFolderName = draftsFolderName;
    this.userPrefix = userEmail ? `/users/${userEmail}` : "/me";
  }

  async searchEmails(
    query: string,
    folder: string = "inbox",
    top: number = 10
  ): Promise<MailMessage[]> {
    const limit = Math.min(top, 50);
    let folderId = folder;

    if (folder !== "inbox" && folder !== "sentitems" && folder !== "archive") {
      const folders = await this.client.api(`${this.userPrefix}/mailFolders`).get();
      const targetFolder = folders.value.find(
        (f: { displayName: string }) =>
          f.displayName.toLowerCase() === folder.toLowerCase()
      );
      if (targetFolder) {
        folderId = targetFolder.id;
      }
    }

    const queryStrategy = this.parseSearchQuery(query);
    const endpoint =
      folder === "inbox"
        ? `${this.userPrefix}/mailFolders/inbox/messages`
        : `${this.userPrefix}/mailFolders/${folderId}/messages`;

    let request = this.client
      .api(endpoint)
      .select("subject,from,receivedDateTime,bodyPreview,hasAttachments")
      .top(limit);

    if (queryStrategy.type === "filter") {
      request = request
        .filter(queryStrategy.value)
        .orderby("receivedDateTime desc");
    } else {
      request = request.search(`"${queryStrategy.value}"`);
    }

    try {
      const response = await request.get();
      return (
        response.value?.map((msg: any) => ({
          id: msg.id,
          subject: msg.subject,
          from: msg.from?.emailAddress?.name || msg.from?.emailAddress?.address || "Unknown",
          date: new Date(msg.receivedDateTime).toLocaleString(),
          preview: (msg.bodyPreview || "").substring(0, 200),
          hasAttachments: msg.hasAttachments,
        })) || []
      );
    } catch (err: any) {
      if (queryStrategy.type === "search") {
        const cleanQuery = queryStrategy.value.replace(/'/g, "''");
        const fallbackResponse = await this.client
          .api(endpoint)
          .filter(`contains(subject,'${cleanQuery}')`)
          .select("subject,from,receivedDateTime,bodyPreview,hasAttachments")
          .top(limit)
          .orderby("receivedDateTime desc")
          .get();

        return fallbackResponse.value.map((msg: any) => ({
          id: msg.id,
          subject: msg.subject,
          from: msg.from?.emailAddress?.name || msg.from?.emailAddress?.address || "Unknown",
          date: new Date(msg.receivedDateTime).toLocaleString(),
          preview: (msg.bodyPreview || "").substring(0, 200),
          hasAttachments: msg.hasAttachments,
        }));
      }
      throw err;
    }
  }

  async getEmail(id: string): Promise<MailMessage> {
    const msg = await this.client.api(`${this.userPrefix}/messages/${id}`).get();
    return {
      id: msg.id,
      subject: msg.subject,
      from: msg.from?.emailAddress?.name || msg.from?.emailAddress?.address || "Unknown",
      date: new Date(msg.receivedDateTime).toLocaleString(),
      preview: (msg.bodyPreview || "").substring(0, 200),
      hasAttachments: msg.hasAttachments,
      body: msg.body?.content,
      bodyType: msg.body?.contentType === "html" ? "HTML" : "text",
    };
  }

  async createDraft(draft: {
    to: string;
    subject: string;
    body: string;
    bodyType?: "text" | "HTML";
    originalEmailId?: string;
  }): Promise<string> {
    const folders = await this.client.api(`${this.userPrefix}/mailFolders`).get();
    const draftsFolder = folders.value.find(
      (f: { displayName: string }) =>
        f.displayName.toLowerCase() === this.draftsFolderName.toLowerCase()
    );

    if (!draftsFolder) {
      throw new Error(`"${this.draftsFolderName}" folder not found. Please create it first.`);
    }

    const recipients = draft.to.split(",").map((email) => ({
      emailAddress: { address: email.trim() },
    }));

    let payload: any = {
      subject: draft.subject,
      toRecipients: recipients,
      body: {
        contentType: draft.bodyType || "text",
        content: draft.body,
      },
      isDraft: true,
    };

    if (draft.originalEmailId) {
      const originalEmail = await this.client
        .api(`${this.userPrefix}/messages/${draft.originalEmailId}`)
        .get();
      payload.inReplyTo = originalEmail.id;
      if (!draft.subject) {
        payload.subject = `Re: ${originalEmail.subject}`;
      }
    }

    const result = await this.client
      .api(`${this.userPrefix}/mailFolders/${draftsFolder.id}/messages`)
      .post(payload);
    return result.id;
  }

  async updateDraft(
    id: string,
    draft: { subject?: string; body?: string; bodyType?: "text" | "HTML" }
  ): Promise<void> {
    const payload: any = {};
    if (draft.subject) payload.subject = draft.subject;
    if (draft.body) {
      payload.body = {
        contentType: draft.bodyType || "text",
        content: draft.body,
      };
    }
    await this.client.api(`${this.userPrefix}/messages/${id}`).patch(payload);
  }

  async deleteDraft(id: string): Promise<void> {
    await this.client.api(`${this.userPrefix}/messages/${id}`).delete();
  }

  async moveEmail(id: string, folderId: string): Promise<void> {
    await this.client.api(`${this.userPrefix}/messages/${id}/move`).post({
      destinationId: folderId,
    });
  }

  async listFolders(): Promise<MailFolder[]> {
    const response = await this.client.api(`${this.userPrefix}/mailFolders`).get();
    return response.value.map((f: any) => ({
      id: f.id,
      displayName: f.displayName,
    }));
  }

  async getAttachments(messageId: string): Promise<MailAttachment[]> {
    const response = await this.client
      .api(`${this.userPrefix}/messages/${messageId}/attachments`)
      .get();
    return response.value.map((a: any) => ({
      id: a.id,
      name: a.name,
      contentType: a.contentType,
      size: a.size,
    }));
  }

  async getAttachmentContent(
    messageId: string,
    attachmentId: string
  ): Promise<string> {
    const attachment = await this.client
      .api(`${this.userPrefix}/messages/${messageId}/attachments/${attachmentId}`)
      .get();
    return attachment.contentBytes; // MSAL returns base64 contentBytes for fileAttachment
  }

  async addAttachmentToDraft(
    draftId: string,
    attachment: { name: string; contentType: string; contentBytes: string }
  ): Promise<void> {
    await this.client.api(`${this.userPrefix}/messages/${draftId}/attachments`).post({
      "@odata.type": "#microsoft.graph.fileAttachment",
      name: attachment.name,
      contentType: attachment.contentType,
      contentBytes: attachment.contentBytes,
    });
  }

  private parseSearchQuery(query: string): {
    type: "filter" | "search";
    value: string;
  } {
    const lowerQuery = query.toLowerCase();
    const parts = lowerQuery.split(" ");

    if (
      lowerQuery.includes("is:unread") ||
      (parts.includes("unread") && !parts.includes("week"))
    ) {
      return { type: "filter", value: "isRead eq false" };
    }

    if (parts.includes("unread") && parts.includes("week")) {
      const weekAgo = new Date();
      weekAgo.setDate(weekAgo.getDate() - 7);
      return {
        type: "filter",
        value: `receivedDateTime gt ${weekAgo.toISOString()} and isRead eq false`,
      };
    }

    if (parts.includes("attachment")) {
      return { type: "filter", value: "hasAttachments eq true" };
    }

    if (lowerQuery.includes("from:")) {
      return { type: "search", value: query };
    }

    if (parts.includes("from")) {
      const fromIdx = parts.indexOf("from") + 1;
      if (fromIdx < parts.length) {
        const name = parts.slice(fromIdx).join(" ");
        return { type: "search", value: `from:${name}` };
      }
    }

    return { type: "search", value: query };
  }
}
