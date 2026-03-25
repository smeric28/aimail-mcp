import { parseSearchQuery } from "./searchEmails.js";

describe("parseSearchQuery", () => {
  it("should handle 'from' query", () => {
    const result = parseSearchQuery("from John Doe");
    expect(result).toEqual({ type: "search", value: "from:john doe" });
  });

  it("should handle 'unread' and 'week' query", () => {
    const result = parseSearchQuery("unread emails from this week");
    expect(result.type).toBe("filter");
    expect(result.value).toMatch(/receivedDateTime gt .* and isRead eq false/);
  });

  it("should handle 'attachment' query", () => {
    const result = parseSearchQuery("emails with attachment");
    expect(result).toEqual({ type: "filter", value: "hasAttachments eq true" });
  });

  it("should handle 'is:unread' filter", () => {
    const result = parseSearchQuery("is:unread");
    expect(result).toEqual({ type: "filter", value: "isRead eq false" });
  });

  it("should handle 'from:' KQL syntax", () => {
    const result = parseSearchQuery("from:john@example.com");
    expect(result).toEqual({ type: "search", value: "from:john@example.com" });
  });

  it("should default to KQL search for general text", () => {
    const result = parseSearchQuery("project update");
    expect(result).toEqual({ type: "search", value: "project update" });
  });
});
