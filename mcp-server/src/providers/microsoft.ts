import { Client } from "@microsoft/microsoft-graph-client";
import {
  MailProvider,
  MailMessage,
  MailFolder,
  MailAttachment,
  CalendarEvent,
  TimeSlot,
} from "../types/mail.js";
import {
  FilesCapable,
  ContactsCapable,
  TasksCapable,
  DriveItem,
  SharePointSite,
  ShareLink,
  ContactCard,
  TaskList,
  TaskItem,
} from "../types/graph.js";

export class MicrosoftProvider
  implements MailProvider, FilesCapable, ContactsCapable, TasksCapable
{
  private client: Client;
  private draftsFolderName: string;
  private userPrefix: string;

  constructor(client: Client, draftsFolderName: string, userEmail?: string) {
    this.client = client;
    this.draftsFolderName = draftsFolderName;
    this.userPrefix = userEmail ? `/users/${userEmail}` : "/me";
  }

  /**
   * Build a provider directly from a per-user Graph access token. Used by the
   * remote HTTP transport, where each request carries the signed-in user's
   * delegated token instead of relying on a local MSAL cache.
   */
  static fromAccessToken(
    accessToken: string,
    draftsFolderName: string,
    userEmail?: string
  ): MicrosoftProvider {
    const client = Client.init({
      authProvider: (done) => done(null, accessToken),
    });
    return new MicrosoftProvider(client, draftsFolderName, userEmail);
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

  async listCalendarEvents(
    start: string,
    end: string,
    top: number = 50
  ): Promise<CalendarEvent[]> {
    const response = await this.client
      .api(`${this.userPrefix}/calendarView`)
      .query({ startDateTime: start, endDateTime: end })
      .select("id,subject,start,end,organizer,attendees,location,bodyPreview,webLink,isOnlineMeeting,onlineMeeting")
      .orderby("start/dateTime")
      .top(Math.min(top, 100))
      .get();

    return (response.value || []).map((e: any) => this.mapEvent(e));
  }

  async createCalendarEvent(event: {
    subject: string;
    start: string;
    end: string;
    timeZone?: string;
    attendees?: string[];
    body?: string;
    bodyType?: "text" | "HTML";
    location?: string;
    isOnlineMeeting?: boolean;
  }): Promise<CalendarEvent> {
    const tz = event.timeZone || "UTC";
    const payload: any = {
      subject: event.subject,
      start: { dateTime: event.start, timeZone: tz },
      end: { dateTime: event.end, timeZone: tz },
      body: {
        contentType: event.bodyType || "HTML",
        content: event.body || "",
      },
    };

    if (event.attendees && event.attendees.length > 0) {
      payload.attendees = event.attendees.map((email) => ({
        emailAddress: { address: email.trim() },
        type: "required",
      }));
    }

    if (event.location) {
      payload.location = { displayName: event.location };
    }

    if (event.isOnlineMeeting) {
      payload.isOnlineMeeting = true;
      payload.onlineMeetingProvider = "teamsForBusiness";
    }

    const created = await this.client
      .api(`${this.userPrefix}/events`)
      .post(payload);

    return this.mapEvent(created);
  }

  async findAvailableTimes(params: {
    durationMinutes: number;
    windowStart: string;
    windowEnd: string;
    attendees?: string[];
    timeZone?: string;
  }): Promise<TimeSlot[]> {
    const tz = params.timeZone || "UTC";
    const payload: any = {
      attendees: (params.attendees || []).map((email) => ({
        emailAddress: { address: email.trim() },
        type: "required",
      })),
      timeConstraint: {
        activityDomain: "work",
        timeSlots: [
          {
            start: { dateTime: params.windowStart, timeZone: tz },
            end: { dateTime: params.windowEnd, timeZone: tz },
          },
        ],
      },
      meetingDuration: `PT${params.durationMinutes}M`,
      maxCandidates: 10,
      isOrganizerOptional: false,
      returnSuggestionReasons: false,
      minimumAttendeePercentage: 100,
    };

    const response = await this.client
      .api(`${this.userPrefix}/findMeetingTimes`)
      .post(payload);

    return (response.meetingTimeSuggestions || []).map((s: any) => ({
      start: s.meetingTimeSlot?.start?.dateTime,
      end: s.meetingTimeSlot?.end?.dateTime,
      confidence: s.confidence,
    }));
  }

  // ===========================================================================
  // Mail write actions
  // ===========================================================================

  private toRecipients(value?: string) {
    if (!value) return [];
    return value
      .split(",")
      .map((e) => e.trim())
      .filter(Boolean)
      .map((address) => ({ emailAddress: { address } }));
  }

  async sendEmail(message: {
    to: string;
    cc?: string;
    bcc?: string;
    subject: string;
    body: string;
    bodyType?: "text" | "HTML";
    saveToSentItems?: boolean;
  }): Promise<void> {
    await this.client.api(`${this.userPrefix}/sendMail`).post({
      message: {
        subject: message.subject,
        body: { contentType: message.bodyType || "HTML", content: message.body },
        toRecipients: this.toRecipients(message.to),
        ccRecipients: this.toRecipients(message.cc),
        bccRecipients: this.toRecipients(message.bcc),
      },
      saveToSentItems: message.saveToSentItems !== false,
    });
  }

  async sendDraft(draftId: string): Promise<void> {
    await this.client.api(`${this.userPrefix}/messages/${draftId}/send`).post({});
  }

  async replyToEmail(
    id: string,
    opts: { body: string; bodyType?: "text" | "HTML"; replyAll?: boolean; send?: boolean }
  ): Promise<string | void> {
    const action = opts.replyAll ? "createReplyAll" : "createReply";
    if (opts.send) {
      const sendAction = opts.replyAll ? "replyAll" : "reply";
      await this.client.api(`${this.userPrefix}/messages/${id}/${sendAction}`).post({
        comment: opts.body,
      });
      return;
    }
    const draft = await this.client.api(`${this.userPrefix}/messages/${id}/${action}`).post({
      comment: opts.body,
    });
    return draft.id;
  }

  async forwardEmail(
    id: string,
    opts: { to: string; comment?: string; send?: boolean }
  ): Promise<string | void> {
    if (opts.send) {
      await this.client.api(`${this.userPrefix}/messages/${id}/forward`).post({
        comment: opts.comment || "",
        toRecipients: this.toRecipients(opts.to),
      });
      return;
    }
    const draft = await this.client.api(`${this.userPrefix}/messages/${id}/createForward`).post({
      comment: opts.comment || "",
      toRecipients: this.toRecipients(opts.to),
    });
    return draft.id;
  }

  async markEmailRead(id: string, isRead: boolean): Promise<void> {
    await this.client.api(`${this.userPrefix}/messages/${id}`).patch({ isRead });
  }

  async flagEmail(id: string, flagged: boolean): Promise<void> {
    await this.client.api(`${this.userPrefix}/messages/${id}`).patch({
      flag: { flagStatus: flagged ? "flagged" : "notFlagged" },
    });
  }

  // ===========================================================================
  // Calendar write actions
  // ===========================================================================

  async getCalendarEvent(id: string): Promise<CalendarEvent> {
    const e = await this.client.api(`${this.userPrefix}/events/${id}`).get();
    return this.mapEvent(e);
  }

  async updateCalendarEvent(
    id: string,
    changes: {
      subject?: string;
      start?: string;
      end?: string;
      timeZone?: string;
      location?: string;
      body?: string;
      bodyType?: "text" | "HTML";
      attendees?: string[];
    }
  ): Promise<CalendarEvent> {
    const tz = changes.timeZone || "UTC";
    const payload: any = {};
    if (changes.subject !== undefined) payload.subject = changes.subject;
    if (changes.start) payload.start = { dateTime: changes.start, timeZone: tz };
    if (changes.end) payload.end = { dateTime: changes.end, timeZone: tz };
    if (changes.location !== undefined) payload.location = { displayName: changes.location };
    if (changes.body !== undefined) {
      payload.body = { contentType: changes.bodyType || "HTML", content: changes.body };
    }
    if (changes.attendees) {
      payload.attendees = changes.attendees.map((email) => ({
        emailAddress: { address: email.trim() },
        type: "required",
      }));
    }
    const updated = await this.client.api(`${this.userPrefix}/events/${id}`).patch(payload);
    return this.mapEvent(updated);
  }

  async cancelCalendarEvent(id: string, comment?: string): Promise<void> {
    // cancel notifies attendees; only valid for events the user organizes.
    await this.client.api(`${this.userPrefix}/events/${id}/cancel`).post({
      comment: comment || "",
    });
  }

  async deleteCalendarEvent(id: string): Promise<void> {
    await this.client.api(`${this.userPrefix}/events/${id}`).delete();
  }

  async respondToCalendarEvent(
    id: string,
    response: "accept" | "decline" | "tentativelyAccept",
    comment?: string
  ): Promise<void> {
    await this.client.api(`${this.userPrefix}/events/${id}/${response}`).post({
      comment: comment || "",
      sendResponse: true,
    });
  }

  // ===========================================================================
  // Files (OneDrive + SharePoint)
  // ===========================================================================

  private driveRoot(driveId?: string): string {
    return driveId ? `/drives/${driveId}` : `${this.userPrefix}/drive`;
  }

  private mapDriveItem(i: any, driveId?: string): DriveItem {
    return {
      id: i.id,
      name: i.name,
      isFolder: !!i.folder,
      size: i.size,
      webUrl: i.webUrl,
      lastModified: i.lastModifiedDateTime,
      mimeType: i.file?.mimeType,
      parentPath: i.parentReference?.path,
      driveId: driveId || i.parentReference?.driveId,
    };
  }

  async listFiles(opts?: { path?: string; driveId?: string; top?: number }): Promise<DriveItem[]> {
    const base = this.driveRoot(opts?.driveId);
    const path = opts?.path?.replace(/^\/+|\/+$/g, "");
    const endpoint = path
      ? `${base}/root:/${encodeURIComponent(path).replace(/%2F/g, "/")}:/children`
      : `${base}/root/children`;
    const res = await this.client.api(endpoint).top(Math.min(opts?.top || 50, 200)).get();
    return (res.value || []).map((i: any) => this.mapDriveItem(i, opts?.driveId));
  }

  async searchFiles(query: string, opts?: { driveId?: string; top?: number }): Promise<DriveItem[]> {
    const base = this.driveRoot(opts?.driveId);
    const res = await this.client
      .api(`${base}/root/search(q='${query.replace(/'/g, "''")}')`)
      .top(Math.min(opts?.top || 25, 200))
      .get();
    return (res.value || []).map((i: any) => this.mapDriveItem(i, opts?.driveId));
  }

  async readFileContent(
    itemId: string,
    driveId?: string
  ): Promise<{ name: string; contentType: string; contentBytes: string }> {
    const base = this.driveRoot(driveId);
    const meta = await this.client.api(`${base}/items/${itemId}`).get();
    const response = await this.client
      .api(`${base}/items/${itemId}/content`)
      .responseType("arraybuffer" as any)
      .get();
    return {
      name: meta.name,
      contentType: meta.file?.mimeType || "application/octet-stream",
      contentBytes: Buffer.from(response as ArrayBuffer).toString("base64"),
    };
  }

  async uploadFile(opts: {
    name: string;
    contentBytes: string;
    parentPath?: string;
    parentItemId?: string;
    driveId?: string;
  }): Promise<DriveItem> {
    const base = this.driveRoot(opts.driveId);
    const buffer = Buffer.from(opts.contentBytes, "base64");
    let endpoint: string;
    if (opts.parentItemId) {
      endpoint = `${base}/items/${opts.parentItemId}:/${encodeURIComponent(opts.name)}:/content`;
    } else {
      const parent = opts.parentPath?.replace(/^\/+|\/+$/g, "");
      endpoint = parent
        ? `${base}/root:/${parent}/${encodeURIComponent(opts.name)}:/content`
        : `${base}/root:/${encodeURIComponent(opts.name)}:/content`;
    }
    // Simple upload (<4MB). Larger files would need an upload session.
    const result = await this.client.api(endpoint).put(buffer);
    return this.mapDriveItem(result, opts.driveId);
  }

  async createFolder(opts: {
    name: string;
    parentItemId?: string;
    driveId?: string;
  }): Promise<DriveItem> {
    const base = this.driveRoot(opts.driveId);
    const endpoint = opts.parentItemId
      ? `${base}/items/${opts.parentItemId}/children`
      : `${base}/root/children`;
    const result = await this.client.api(endpoint).post({
      name: opts.name,
      folder: {},
      "@microsoft.graph.conflictBehavior": "rename",
    });
    return this.mapDriveItem(result, opts.driveId);
  }

  async deleteFile(itemId: string, driveId?: string): Promise<void> {
    const base = this.driveRoot(driveId);
    await this.client.api(`${base}/items/${itemId}`).delete();
  }

  async createShareLink(opts: {
    itemId: string;
    driveId?: string;
    type?: "view" | "edit";
    scope?: "anonymous" | "organization";
  }): Promise<ShareLink> {
    const base = this.driveRoot(opts.driveId);
    const result = await this.client.api(`${base}/items/${opts.itemId}/createLink`).post({
      type: opts.type || "view",
      scope: opts.scope || "organization",
    });
    return {
      url: result.link?.webUrl,
      type: result.link?.type,
      scope: result.link?.scope,
      expiresOn: result.expirationDateTime,
    };
  }

  async listSharePointSites(query?: string): Promise<SharePointSite[]> {
    const endpoint = query
      ? `/sites?search=${encodeURIComponent(query)}`
      : `/sites?search=*`;
    const res = await this.client.api(endpoint).get();
    return (res.value || []).map((s: any) => ({
      id: s.id,
      displayName: s.displayName || s.name,
      webUrl: s.webUrl,
    }));
  }

  async listSiteDrives(siteId: string): Promise<Array<{ id: string; name: string }>> {
    const res = await this.client.api(`/sites/${siteId}/drives`).get();
    return (res.value || []).map((d: any) => ({ id: d.id, name: d.name }));
  }

  // ===========================================================================
  // Contacts
  // ===========================================================================

  private mapContact(c: any): ContactCard {
    return {
      id: c.id,
      displayName: c.displayName,
      givenName: c.givenName,
      surname: c.surname,
      emails: (c.emailAddresses || []).map((e: any) => e.address).filter(Boolean),
      phones: [...(c.businessPhones || []), ...(c.homePhones || []), c.mobilePhone].filter(Boolean),
      company: c.companyName,
      jobTitle: c.jobTitle,
    };
  }

  async listContacts(top: number = 50): Promise<ContactCard[]> {
    const res = await this.client
      .api(`${this.userPrefix}/contacts`)
      .top(Math.min(top, 200))
      .get();
    return (res.value || []).map((c: any) => this.mapContact(c));
  }

  async searchContacts(query: string, top: number = 25): Promise<ContactCard[]> {
    const clean = query.replace(/'/g, "''");
    const res = await this.client
      .api(`${this.userPrefix}/contacts`)
      .filter(
        `startswith(displayName,'${clean}') or startswith(givenName,'${clean}') or startswith(surname,'${clean}')`
      )
      .top(Math.min(top, 200))
      .get();
    return (res.value || []).map((c: any) => this.mapContact(c));
  }

  private contactPayload(c: any): any {
    const payload: any = {};
    if (c.givenName !== undefined) payload.givenName = c.givenName;
    if (c.surname !== undefined) payload.surname = c.surname;
    if (c.displayName !== undefined) payload.displayName = c.displayName;
    if (c.company !== undefined) payload.companyName = c.company;
    if (c.jobTitle !== undefined) payload.jobTitle = c.jobTitle;
    if (c.emails) {
      payload.emailAddresses = c.emails.map((address: string) => ({ address }));
    }
    if (c.phones) payload.businessPhones = c.phones;
    return payload;
  }

  async createContact(contact: {
    givenName?: string;
    surname?: string;
    displayName?: string;
    emails?: string[];
    phones?: string[];
    company?: string;
    jobTitle?: string;
  }): Promise<ContactCard> {
    const result = await this.client
      .api(`${this.userPrefix}/contacts`)
      .post(this.contactPayload(contact));
    return this.mapContact(result);
  }

  async updateContact(id: string, changes: any): Promise<ContactCard> {
    const result = await this.client
      .api(`${this.userPrefix}/contacts/${id}`)
      .patch(this.contactPayload(changes));
    return this.mapContact(result);
  }

  async deleteContact(id: string): Promise<void> {
    await this.client.api(`${this.userPrefix}/contacts/${id}`).delete();
  }

  // ===========================================================================
  // Tasks (Microsoft To Do)
  // ===========================================================================

  private mapTask(t: any, listId?: string): TaskItem {
    return {
      id: t.id,
      title: t.title,
      status: t.status,
      dueDateTime: t.dueDateTime?.dateTime,
      reminderDateTime: t.reminderDateTime?.dateTime,
      body: t.body?.content,
      importance: t.importance,
      listId,
    };
  }

  async listTaskLists(): Promise<TaskList[]> {
    const res = await this.client.api(`${this.userPrefix}/todo/lists`).get();
    return (res.value || []).map((l: any) => ({
      id: l.id,
      displayName: l.displayName,
      isDefault: l.wellknownListName === "defaultList",
    }));
  }

  async listTasks(
    listId: string,
    opts?: { includeCompleted?: boolean; top?: number }
  ): Promise<TaskItem[]> {
    let req = this.client
      .api(`${this.userPrefix}/todo/lists/${listId}/tasks`)
      .top(Math.min(opts?.top || 50, 200));
    if (!opts?.includeCompleted) {
      req = req.filter("status ne 'completed'");
    }
    const res = await req.get();
    return (res.value || []).map((t: any) => this.mapTask(t, listId));
  }

  private taskPayload(task: any): any {
    const payload: any = {};
    if (task.title !== undefined) payload.title = task.title;
    if (task.status !== undefined) payload.status = task.status;
    if (task.importance !== undefined) payload.importance = task.importance;
    if (task.body !== undefined) {
      payload.body = { content: task.body, contentType: "text" };
    }
    if (task.dueDateTime !== undefined) {
      payload.dueDateTime = { dateTime: task.dueDateTime, timeZone: "UTC" };
    }
    if (task.reminderDateTime !== undefined) {
      payload.reminderDateTime = { dateTime: task.reminderDateTime, timeZone: "UTC" };
    }
    return payload;
  }

  async createTask(
    listId: string,
    task: {
      title: string;
      body?: string;
      dueDateTime?: string;
      reminderDateTime?: string;
      importance?: "low" | "normal" | "high";
    }
  ): Promise<TaskItem> {
    const result = await this.client
      .api(`${this.userPrefix}/todo/lists/${listId}/tasks`)
      .post(this.taskPayload(task));
    return this.mapTask(result, listId);
  }

  async updateTask(listId: string, taskId: string, changes: any): Promise<TaskItem> {
    const result = await this.client
      .api(`${this.userPrefix}/todo/lists/${listId}/tasks/${taskId}`)
      .patch(this.taskPayload(changes));
    return this.mapTask(result, listId);
  }

  async completeTask(listId: string, taskId: string): Promise<TaskItem> {
    return this.updateTask(listId, taskId, { status: "completed" });
  }

  async deleteTask(listId: string, taskId: string): Promise<void> {
    await this.client
      .api(`${this.userPrefix}/todo/lists/${listId}/tasks/${taskId}`)
      .delete();
  }

  private mapEvent(e: any): CalendarEvent {
    return {
      id: e.id,
      subject: e.subject || "",
      start: e.start?.dateTime ? `${e.start.dateTime}${e.start.timeZone ? " " + e.start.timeZone : ""}` : "",
      end: e.end?.dateTime ? `${e.end.dateTime}${e.end.timeZone ? " " + e.end.timeZone : ""}` : "",
      organizer: e.organizer?.emailAddress?.address,
      attendees: (e.attendees || []).map((a: any) => ({
        email: a.emailAddress?.address,
        name: a.emailAddress?.name,
        response: a.status?.response,
      })),
      location: e.location?.displayName,
      body: e.bodyPreview,
      webLink: e.webLink,
      isOnlineMeeting: e.isOnlineMeeting,
      onlineMeetingUrl: e.onlineMeeting?.joinUrl,
    };
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
