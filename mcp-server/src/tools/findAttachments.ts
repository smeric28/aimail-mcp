import { MailProvider } from "../types/mail.js";

interface Tool {
    name: string;
    description: string;
    inputSchema: {
        type: "object";
        properties: {
            query: { type: "string"; description: "Search term for attachment name or email content" };
            top: { type: "number"; description: "Max emails to search (default: 10)" };
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

export const findAttachmentsTool: Tool = {
    name: "find_attachments",
    description:
        "Search for attachments across emails. Returns a list of specific files found, " +
        "including their IDs and the parent email's ID. Use this to find a file to add to a draft.",
    inputSchema: {
        type: "object",
        properties: {
            query: {
                type: "string",
                description: "Search term for attachment name or email content",
            },
            top: {
                type: "number",
                description: "Max emails to search (default: 10)",
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
        required: ["query"],
    },
    async handler(provider: MailProvider, args: Record<string, unknown>) {
        const query = args.query as string;
        const top = Math.min(Number(args.top) || 10, 50);

        try {
            // Search for emails matching the query
            const emails = await provider.searchEmails(query, "inbox", top);

            const foundAttachments: any[] = [];

            for (const email of emails) {
                if (email.hasAttachments) {
                    try {
                        const attachments = await provider.getAttachments(email.id);
                        for (const att of attachments) {
                            foundAttachments.push({
                                attachmentId: att.id,
                                name: att.name,
                                size: att.size,
                                contentType: att.contentType,
                                parentEmailId: email.id,
                                parentEmailSubject: email.subject,
                                parentEmailDate: email.date,
                            });
                        }
                    } catch (err) {
                        // Skip emails where we can't retrieve attachments
                    }
                }
            }

            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify(foundAttachments, null, 2),
                    },
                ],
            };
        } catch (error: any) {
            return {
                content: [
                    {
                        type: "text",
                        text: `Error finding attachments: ${error.message}`,
                    },
                ],
                isError: true,
            };
        }
    },
};
