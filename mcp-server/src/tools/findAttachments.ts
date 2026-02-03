import { Client } from "@microsoft/microsoft-graph-client";
import { parseSearchQuery } from "./searchEmails.js";

interface Tool {
    name: string;
    description: string;
    inputSchema: {
        type: "object";
        properties: {
            query: { type: "string"; description: "Search term for attachment name or email content" };
            top: { type: "number"; description: "Max emails to search (default: 10)" };
        };
        required: string[];
    };
    handler: (
        client: Client,
        args: Record<string, unknown>
    ) => Promise<{
        content: Array<{ type: "text"; text: string }>;
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
        },
        required: ["query"],
    },
    async handler(client: Client, args: Record<string, unknown>) {
        const query = args.query as string;
        const top = Math.min(Number(args.top) || 10, 50);

        // Reuse existing search logic to find candidate emails
        // We force a search for items with attachments
        const queryStrategy = parseSearchQuery(query);

        // Build query - force querying for items with attachments
        let request = client.api("/me/messages")
            .select("id,subject,receivedDateTime,hasAttachments")
            .expand("attachments($select=id,name,size,contentType)")
            .top(top);

        // Apply filter or search
        if (queryStrategy.type === 'filter') {
            // Append attachment requirement if not present
            let filter = queryStrategy.value;
            if (!filter.includes("hasAttachments")) {
                filter = `(${filter}) and hasAttachments eq true`;
            }
            request = request.filter(filter).orderby("receivedDateTime desc");
        } else {
            // For KQL, we append 'hasAttachments:true' to the search query
            const searchVal = `${queryStrategy.value} hasAttachments:true`;
            request = request.search(`"${searchVal}"`);
        }

        try {
            const response = await request.get();

            const foundAttachments: any[] = [];

            if (response.value) {
                for (const msg of response.value) {
                    if (msg.attachments && msg.attachments.length > 0) {
                        for (const att of msg.attachments) {
                            // Optional: client-side filter if the query looks like a filename
                            // But for now return all attachments in the matching emails
                            foundAttachments.push({
                                attachmentId: att.id,
                                name: att.name,
                                size: att.size,
                                contentType: att.contentType,
                                parentEmailId: msg.id,
                                parentEmailSubject: msg.subject,
                                parentEmailDate: new Date(msg.receivedDateTime).toLocaleString()
                            });
                        }
                    }
                }
            }

            // Filter client-side if the query is specific to a filename to reduce noise
            // (This is a heuristic: if we searched for 'report' and got an email 'Weekly Report' 
            // with 5 images, we probably want to prioritize the pdf 'report.pdf' if it exists, 
            // but showing all is safer).

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
