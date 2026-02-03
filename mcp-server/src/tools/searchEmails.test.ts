import { parseSearchQuery } from "./searchEmails.js";

describe("parseSearchQuery", () => {
  it("should handle 'from' query", () => {
    const result = parseSearchQuery("from John Doe");
    expect(result).toBe("from/emailAddress/name contains 'john doe'");
  });

  it("should handle 'unread' and 'week' query", () => {
    const result = parseSearchQuery("unread emails from this week");
    // We can't easily check the exact date, so we'll check the structure
    expect(result).toMatch(/receivedDateTime gt .* and isRead eq false/);
  });

  it("should handle 'attachment' query", () => {
    const result = parseSearchQuery("emails with attachment");
    expect(result).toBe("hasAttachments eq true");
  });

  it("should default to subject and body search", () => {
    const result = parseSearchQuery("project update");
    expect(result).toBe("contains(subject,'project update') or contains(bodyPreview,'project update')");
  });
});
