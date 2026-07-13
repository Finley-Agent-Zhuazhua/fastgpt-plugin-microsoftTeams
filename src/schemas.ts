import type { InputSchemaMetaType, OutputSchemaMetaType, SecretSchemaMetaType } from "@fastgpt-plugin/sdk-factory";
import z from "zod";

const input = (title: string, description: string, max = 2048) =>
  z.string().min(1).max(max).meta({ title, description, toolDescription: description } satisfies InputSchemaMetaType);
const limit = z.number().int().min(1).max(100).default(20);
const nextLink = z.string().url().max(4096).optional().nullable();

export const secretSchema = z.object({
  accessToken: z.string().min(1).meta({
    title: "Microsoft OAuth access token",
    description: "OAuth bearer token with the minimum Microsoft Graph Teams and online meeting permissions.",
    isSecret: true
  } satisfies SecretSchemaMetaType)
});

export const listTeamsInputSchema = z.object({ maxResults: limit, nextLink });
export const listChannelsInputSchema = z.object({ teamId: input("Team ID", "Microsoft Teams team ID."), maxResults: limit, nextLink });
export const sendChannelMessageInputSchema = z.object({
  teamId: input("Team ID", "Microsoft Teams team ID."),
  channelId: input("Channel ID", "Microsoft Teams channel ID."),
  content: input("Message", "Message body.", 28000),
  contentType: z.enum(["text", "html"]).default("text")
});
export const sendChatMessageInputSchema = z.object({
  chatId: input("Chat ID", "Microsoft Teams chat ID."),
  content: input("Message", "Message body.", 28000),
  contentType: z.enum(["text", "html"]).default("text")
});
const dateTime = (title: string) => input(title, "ISO 8601 date-time with time-zone offset.", 64);
export const createOnlineMeetingInputSchema = z.object({
  subject: input("Subject", "Online meeting subject."),
  start: dateTime("Start"),
  end: dateTime("End")
}).refine((value) => Date.parse(value.end) > Date.parse(value.start), { message: "end must be after start" });

const teamSchema = z.object({ id: z.string(), displayName: z.string().optional(), description: z.string().optional(), webUrl: z.string().optional() });
const channelSchema = z.object({ id: z.string(), displayName: z.string().optional(), description: z.string().optional(), webUrl: z.string().optional(), membershipType: z.string().optional() });
const messageSchema = z.object({ id: z.string(), webUrl: z.string().optional(), createdDateTime: z.string().optional() });
const meetingSchema = z.object({ id: z.string(), subject: z.string().optional(), start: z.string().optional(), end: z.string().optional(), joinUrl: z.string().optional() });
const success = z.literal(true).meta({ title: "Success" } satisfies OutputSchemaMetaType);
export const listTeamsOutputSchema = z.object({ success, teams: z.array(teamSchema), returnedCount: z.number().int().nonnegative(), nextLink: z.string().optional() });
export const listChannelsOutputSchema = z.object({ success, channels: z.array(channelSchema), returnedCount: z.number().int().nonnegative(), nextLink: z.string().optional() });
export const sendMessageOutputSchema = z.object({ success, message: messageSchema });
export const createOnlineMeetingOutputSchema = z.object({ success, meeting: meetingSchema });

export type Secrets = z.output<typeof secretSchema>;
export type ListTeamsInput = z.output<typeof listTeamsInputSchema>;
export type ListChannelsInput = z.output<typeof listChannelsInputSchema>;
export type SendChannelMessageInput = z.output<typeof sendChannelMessageInputSchema>;
export type SendChatMessageInput = z.output<typeof sendChatMessageInputSchema>;
export type CreateOnlineMeetingInput = z.output<typeof createOnlineMeetingInputSchema>;
export type Team = z.output<typeof teamSchema>;
export type Channel = z.output<typeof channelSchema>;
export type SentMessage = z.output<typeof messageSchema>;
export type OnlineMeeting = z.output<typeof meetingSchema>;
