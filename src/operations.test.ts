import { afterEach, describe, expect, it, vi } from "vitest";
import { createOnlineMeeting, listChannels, listTeams, sendChannelMessage, sendChatMessage } from "./operations";
import { createOnlineMeetingInputSchema } from "./schemas";

const token = "test-token";
const json = (body: unknown, status = 200) => new Response(JSON.stringify(body), { status, headers: { "Content-Type": "application/json" } });
afterEach(() => vi.restoreAllMocks());

describe("schemas", () => {
  it("rejects inverted meeting times", () => {
    expect(() => createOnlineMeetingInputSchema.parse({ subject: "x", start: "2026-07-13T11:00:00Z", end: "2026-07-13T10:00:00Z" })).toThrow(/end must be after start/);
  });
});

describe("operations", () => {
  it("lists and formats joined teams", async () => {
    const fetchMock = vi.spyOn(globalThis, "fetch").mockResolvedValue(json({
      "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/joinedTeams?$skip=20",
      value: [{ id: "team-1", displayName: "Engineering", description: "Builds things" }]
    }));
    await expect(listTeams({ accessToken: token, maxResults: 20 })).resolves.toMatchObject({ returnedCount: 1, teams: [{ id: "team-1", displayName: "Engineering" }], nextLink: expect.stringContaining("skip") });
    expect(String(fetchMock.mock.calls[0]?.[0])).toContain("%24top=20");
  });

  it("lists channels for the requested team", async () => {
    const fetchMock = vi.spyOn(globalThis, "fetch").mockResolvedValue(json({ value: [{ id: "channel-1", displayName: "General", membershipType: "standard" }] }));
    await expect(listChannels({ accessToken: token, teamId: "team / 1", maxResults: 10 })).resolves.toMatchObject({ returnedCount: 1, channels: [{ id: "channel-1", displayName: "General" }] });
    expect(String(fetchMock.mock.calls[0]?.[0])).toContain("teams/team%20%2F%201/channels");
  });

  it("sends channel and chat messages with the correct bodies", async () => {
    const fetchMock = vi.spyOn(globalThis, "fetch")
      .mockResolvedValueOnce(json({ id: "message-1", createdDateTime: "2026-07-13T10:00:00Z" }, 201))
      .mockResolvedValueOnce(json({ id: "message-2", webUrl: "https://teams.microsoft.com/message/2" }, 201));
    await expect(sendChannelMessage({ accessToken: token, teamId: "team-1", channelId: "channel-1", content: "Done", contentType: "text" })).resolves.toMatchObject({ message: { id: "message-1" } });
    expect(JSON.parse(String(fetchMock.mock.calls[0]?.[1]?.body))).toEqual({ body: { content: "Done", contentType: "text" } });
    await expect(sendChatMessage({ accessToken: token, chatId: "chat-1", content: "<b>Alert</b>", contentType: "html" })).resolves.toMatchObject({ message: { id: "message-2" } });
    expect(String(fetchMock.mock.calls[1]?.[0])).toContain("/chats/chat-1/messages");
  });

  it("creates an online meeting and parses the join URL", async () => {
    const fetchMock = vi.spyOn(globalThis, "fetch").mockResolvedValue(json({ id: "meeting-1", subject: "Review", startDateTime: "2026-07-13T10:00:00Z", endDateTime: "2026-07-13T11:00:00Z", joinWebUrl: "https://teams.microsoft.com/l/meetup-join/test" }, 201));
    await expect(createOnlineMeeting({ accessToken: token, subject: "Review", start: "2026-07-13T10:00:00Z", end: "2026-07-13T11:00:00Z" })).resolves.toMatchObject({ meeting: { id: "meeting-1", joinUrl: expect.stringContaining("meetup-join") } });
    expect(JSON.parse(String(fetchMock.mock.calls[0]?.[1]?.body))).toMatchObject({ subject: "Review", startDateTime: "2026-07-13T10:00:00Z" });
  });

  it("reports Graph errors and rejects untrusted pagination URLs", async () => {
    vi.spyOn(globalThis, "fetch").mockResolvedValue(json({ error: { message: "permission denied" } }, 403));
    await expect(listTeams({ accessToken: token, maxResults: 20 })).rejects.toThrow(/permission denied/);
    await expect(listTeams({ accessToken: token, maxResults: 20, nextLink: "https://evil.example/v1.0/teams" })).rejects.toThrow(/Microsoft Graph v1.0/);
  });
});
