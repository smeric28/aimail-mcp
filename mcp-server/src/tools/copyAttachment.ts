import { Client } from "@microsoft/microsoft-graph-client";

interface Tool {
    name: string;
    description: string;
    inputSchema: {
        type: "object";
        properties: {
            sourceMessageId: { type: "string"; description: string };
            attachmentName: { type: "string"; description: string };
            targetDraftId: { type: "string"; description: string };
        };
        required: string[];
    };
    handler: (
        client: Client,
        args: Record<string, unknown>
    ) => Promise<{
        content: Array<{ type: "text"; text: string }>;
        isError?: boolean;
    }>;
}

export const copyAttachmentTool: Tool = {
    name: "copy_attachment_to_draft",
    description: "Copy an attachment from an existing message/email to a draft.",
    inputSchema: {
        type: "object",
        properties: {
            sourceMessageId: {
                type: "string",
                description: "ID of the email containing the attachment",
            },
            attachmentName: {
                type: "string",
                description: "Name (or partial name) of the attachment to copy",
            },
            targetDraftId: {
                type: "string",
                description: "ID of the draft to add the attachment to",
            },
        },
        required: ["sourceMessageId", "attachmentName", "targetDraftId"],
    },
    async handler(client: Client, args: Record<string, unknown>) {
        const sourceId = args.sourceMessageId as string;
        const targetId = args.targetDraftId as string;
        const queryString = (args.attachmentName as string).toLowerCase();

        try {
            // 1. List attachments of source message to find the ID
            const attachmentsReq = await client.api(`/me/messages/${sourceId}/attachments`)
                .select("id,name,size,contentType")
                .get();

            const attachments = attachmentsReq.value;
            const match = attachments.find((a: any) => a.name.toLowerCase().includes(queryString));

            if (!match) {
                return {
                    content: [{
                        type: "text",
                        text: `Attachment matching '${queryString}' not found on message ${sourceId}. Available: ${attachments.map((a: any) => a.name).join(", ")}`
                    }],
                    isError: true
                };
            }

            // 2. Get the actual attachment content (bytes)
            // We need the raw Bytes.
            const attachmentContent = await client.api(`/me/messages/${sourceId}/attachments/${match.id}`)
                .get();

            // Note: Graph API returns fileAttachment resource which has 'contentBytes' (base64)
            if (!attachmentContent.contentBytes) {
                return {
                    content: [{
                        type: "text",
                        text: `Attachment '${match.name}' does not appear to be a file attachment (missing contentBytes). Type: ${attachmentContent['@odata.type']}`
                    }],
                    isError: true
                };
            }

            // 3. Add to target draft
            const newAttachment = {
                "@odata.type": "#microsoft.graph.fileAttachment",
                name: match.name, // Keep original name
                contentBytes: attachmentContent.contentBytes,
                contentType: attachmentContent.contentType
            };

            const result = await client.api(`/me/messages/${targetId}/attachments`)
                .post(newAttachment);

            return {
                content: [
                    {
                        type: "text",
                        text: `Successfully copied attachment '${match.name}' to draft ${targetId}. New Attachment ID: ${result.id}`,
                    },
                ],
            };
        } catch (error: any) {
            return {
                content: [
                    {
                        type: "text",
                        text: `Error copying attachment: ${error.message}`,
                    },
                ],
                isError: true,
            };
        }
    },
};
