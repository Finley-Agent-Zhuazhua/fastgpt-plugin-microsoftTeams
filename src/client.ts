import type { Channel, CreateOnlineMeetingInput, ListChannelsInput, ListTeamsInput, OnlineMeeting, SendChannelMessageInput, SendChatMessageInput, SentMessage, Team } from "./schemas";

export const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

export class TeamsClient {
  constructor(private readonly token: string, private readonly fetchFn: typeof fetch = fetch) {
    if (!token.trim()) throw new Error("Microsoft Teams accessToken secret is required");
  }

  listTeams(input: ListTeamsInput) {
    const url = input.nextLink ? this.pageUrl(input.nextLink) : new URL(`${GRAPH_BASE}/me/joinedTeams`);
    if (!input.nextLink) {
      url.searchParams.set("$top", String(input.maxResults));
      url.searchParams.set("$select", "id,displayName,description,webUrl");
    }
    return this.request("GET", url);
  }

  listChannels(input: ListChannelsInput) {
    const url = input.nextLink ? this.pageUrl(input.nextLink) : new URL(`${GRAPH_BASE}/teams/${encodeURIComponent(input.teamId)}/channels`);
    if (!input.nextLink) {
      url.searchParams.set("$top", String(input.maxResults));
      url.searchParams.set("$select", "id,displayName,description,webUrl,membershipType");
    }
    return this.request("GET", url);
  }

  sendChannelMessage(input: SendChannelMessageInput) {
    const url = new URL(`${GRAPH_BASE}/teams/${encodeURIComponent(input.teamId)}/channels/${encodeURIComponent(input.channelId)}/messages`);
    return this.request("POST", url, { body: { content: input.content, contentType: input.contentType } });
  }

  sendChatMessage(input: SendChatMessageInput) {
    const url = new URL(`${GRAPH_BASE}/chats/${encodeURIComponent(input.chatId)}/messages`);
    return this.request("POST", url, { body: { content: input.content, contentType: input.contentType } });
  }

  createOnlineMeeting(input: CreateOnlineMeetingInput) {
    return this.request("POST", new URL(`${GRAPH_BASE}/me/onlineMeetings`), {
      subject: input.subject,
      startDateTime: input.start,
      endDateTime: input.end
    });
  }

  private pageUrl(nextLink: string) {
    const url = new URL(nextLink);
    if (url.origin !== "https://graph.microsoft.com" || !url.pathname.startsWith("/v1.0/")) {
      throw new Error("nextLink must be a Microsoft Graph v1.0 URL");
    }
    return url;
  }

  private async request(method: string, url: URL, body?: Record<string, unknown>) {
    const response = await this.fetchFn(url, {
      method,
      headers: {
        Authorization: `Bearer ${this.token}`,
        Accept: "application/json",
        ...(body ? { "Content-Type": "application/json" } : {})
      },
      ...(body ? { body: JSON.stringify(body) } : {})
    });
    const text = await response.text();
    let payload: unknown = {};
    if (text) {
      try {
        payload = JSON.parse(text);
      } catch {
        throw new Error(`Invalid JSON from Microsoft Graph ${url.pathname}`);
      }
    }
    if (!response.ok) throw new Error(`Microsoft Graph ${method} ${url.pathname} failed: ${errorMessage(payload, response.statusText)}`);
    return payload && typeof payload === "object" && !Array.isArray(payload) ? payload as Record<string, unknown> : {};
  }
}

export function formatTeam(value: unknown): Team {
  const item = object(value, "team");
  const result: Team = { id: requiredId(item, "team") };
  if (typeof item.displayName === "string") result.displayName = item.displayName;
  if (typeof item.description === "string") result.description = item.description;
  if (typeof item.webUrl === "string") result.webUrl = item.webUrl;
  return result;
}
export function formatChannel(value: unknown): Channel {
  const item = object(value, "channel");
  const result: Channel = { id: requiredId(item, "channel") };
  if (typeof item.displayName === "string") result.displayName = item.displayName;
  if (typeof item.description === "string") result.description = item.description;
  if (typeof item.webUrl === "string") result.webUrl = item.webUrl;
  if (typeof item.membershipType === "string") result.membershipType = item.membershipType;
  return result;
}
export function formatSentMessage(value: unknown): SentMessage {
  const item = object(value, "message");
  const result: SentMessage = { id: requiredId(item, "message") };
  if (typeof item.webUrl === "string") result.webUrl = item.webUrl;
  if (typeof item.createdDateTime === "string") result.createdDateTime = item.createdDateTime;
  return result;
}
export function formatMeeting(value: unknown): OnlineMeeting {
  const item = object(value, "online meeting");
  const result: OnlineMeeting = { id: requiredId(item, "online meeting") };
  if (typeof item.subject === "string") result.subject = item.subject;
  if (typeof item.joinWebUrl === "string") result.joinUrl = item.joinWebUrl;
  if (typeof item.startDateTime === "string") result.start = item.startDateTime;
  if (typeof item.endDateTime === "string") result.end = item.endDateTime;
  return result;
}

function object(value: unknown, kind: string) {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new Error(`Microsoft Graph ${kind} must be an object`);
  return value as Record<string, unknown>;
}
function requiredId(value: Record<string, unknown>, kind: string) {
  if (typeof value.id !== "string" || !value.id) throw new Error(`Microsoft Graph ${kind} missing id`);
  return value.id;
}
function errorMessage(value: unknown, fallback: string) {
  if (value && typeof value === "object") {
    const error = (value as Record<string, unknown>).error;
    if (error && typeof error === "object" && typeof (error as Record<string, unknown>).message === "string") return (error as Record<string, unknown>).message as string;
  }
  return fallback;
}
