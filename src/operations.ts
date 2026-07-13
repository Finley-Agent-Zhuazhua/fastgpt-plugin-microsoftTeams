import { formatChannel, formatMeeting, formatSentMessage, formatTeam, TeamsClient } from "./client";
import type { CreateOnlineMeetingInput, ListChannelsInput, ListTeamsInput, Secrets, SendChannelMessageInput, SendChatMessageInput } from "./schemas";

export async function listTeams(input: ListTeamsInput & Secrets) {
  const data = await new TeamsClient(input.accessToken).listTeams(input);
  const teams = Array.isArray(data.value) ? data.value.map(formatTeam) : [];
  return { success: true as const, teams, returnedCount: teams.length, ...(typeof data["@odata.nextLink"] === "string" ? { nextLink: data["@odata.nextLink"] } : {}) };
}
export async function listChannels(input: ListChannelsInput & Secrets) {
  const data = await new TeamsClient(input.accessToken).listChannels(input);
  const channels = Array.isArray(data.value) ? data.value.map(formatChannel) : [];
  return { success: true as const, channels, returnedCount: channels.length, ...(typeof data["@odata.nextLink"] === "string" ? { nextLink: data["@odata.nextLink"] } : {}) };
}
export async function sendChannelMessage(input: SendChannelMessageInput & Secrets) {
  return { success: true as const, message: formatSentMessage(await new TeamsClient(input.accessToken).sendChannelMessage(input)) };
}
export async function sendChatMessage(input: SendChatMessageInput & Secrets) {
  return { success: true as const, message: formatSentMessage(await new TeamsClient(input.accessToken).sendChatMessage(input)) };
}
export async function createOnlineMeeting(input: CreateOnlineMeetingInput & Secrets) {
  return { success: true as const, meeting: formatMeeting(await new TeamsClient(input.accessToken).createOnlineMeeting(input)) };
}
