import { MailProvider } from "../types/mail.js";

interface Tool {
    name: string;
    description: string;
    inputSchema: {
        type: "object";
        properties: {
            sourceMessageId: { type: "string"; description: string };
            attachmentName: { type: "string"; description: string };
            targetDraftId: { type: "string"; description: string };
            accountId: { type: "string"; description: "Optional account ID" };
            mailbox: { type: "string"; description: "Optional shared mailbox email" };
        };
        required: string[];
    };
    handler: (
        provider: MailProvider,
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
            accountId: {
                type: "string",
                description: "Optional account ID",
            },
            mailbox: {
                type: "string",
                description: "Optional shared mailbox email",
            },
        },
        required: ["sourceMessageId", "attachmentName", "targetDraftId"],
    },
    async handler(provider: MailProvider, args: Record<string, unknown>) {
        const sourceId = args.sourceMessageId as string;
        const targetId = args.targetDraftId as string;
        const queryString = (args.attachmentName as string).toLowerCase();

        try {
            // 1. List attachments of source message to find the match
            const attachments = await provider.getAttachments(sourceId);
            const match = attachments.find((a) => a.name.toLowerCase().includes(queryString));

            if (!match) {
                return {
                    content: [{
                        type: "text",
                        text: `Attachment matching '${queryString}' not found on message ${sourceId}. Available: ${attachments.map((a) => a.name).join(", ")}`
                    }],
                    isError: true
                };
            }

            // 2. Get the actual attachment content (base64)
            const contentBytes = await provider.getAttachmentContent(sourceId, match.id);

            if (!contentBytes) {
                return {
                    content: [{
                        type: "text",
                        text: `Attachment '${match.name}' has no downloadable content.`
                    }],
                    isError: true
                };
            }

            // 3. Add to target draft
            await provider.addAttachmentToDraft(targetId, {
                name: match.name,
                contentType: match.contentType,
                contentBytes: contentBytes,
            });

            return {
                content: [
                    {
                        type: "text",
                        text: `Successfully copied attachment '${match.name}' to draft ${targetId}.`,
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
