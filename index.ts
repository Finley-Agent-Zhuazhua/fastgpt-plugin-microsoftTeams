import { createToolHandler, defineToolSet } from "@fastgpt-plugin/sdk-factory";
import { createOnlineMeeting, listChannels, listTeams, sendChannelMessage, sendChatMessage } from "./src/operations";
import { createOnlineMeetingInputSchema, createOnlineMeetingOutputSchema, listChannelsInputSchema, listChannelsOutputSchema, listTeamsInputSchema, listTeamsOutputSchema, secretSchema, sendChannelMessageInputSchema, sendChatMessageInputSchema, sendMessageOutputSchema, type Secrets } from "./src/schemas";

const secrets = (value: Secrets | undefined) => {
  if (!value?.accessToken?.trim()) throw new Error("Microsoft Teams accessToken secret is required");
  return value;
};
const handlers = {
  listTeams: createToolHandler({ inputSchema: listTeamsInputSchema, outputSchema: listTeamsOutputSchema, secretSchema, handler: (input, ctx) => listTeams({ ...input, ...secrets(ctx.secrets) }) }),
  listChannels: createToolHandler({ inputSchema: listChannelsInputSchema, outputSchema: listChannelsOutputSchema, secretSchema, handler: (input, ctx) => listChannels({ ...input, ...secrets(ctx.secrets) }) }),
  sendChannelMessage: createToolHandler({ inputSchema: sendChannelMessageInputSchema, outputSchema: sendMessageOutputSchema, secretSchema, handler: (input, ctx) => sendChannelMessage({ ...input, ...secrets(ctx.secrets) }) }),
  sendChatMessage: createToolHandler({ inputSchema: sendChatMessageInputSchema, outputSchema: sendMessageOutputSchema, secretSchema, handler: (input, ctx) => sendChatMessage({ ...input, ...secrets(ctx.secrets) }) }),
  createOnlineMeeting: createToolHandler({ inputSchema: createOnlineMeetingInputSchema, outputSchema: createOnlineMeetingOutputSchema, secretSchema, handler: (input, ctx) => createOnlineMeeting({ ...input, ...secrets(ctx.secrets) }) })
};

export default defineToolSet({
  manifest: {
    pluginId: "microsoftTeams",
    name: { en: "Microsoft Teams", "zh-CN": "Microsoft Teams" },
    description: { en: "Read Teams structure, send messages, and create online meetings.", "zh-CN": "读取 Teams 团队与频道、发送消息并创建在线会议。" },
    version: "0.1.0",
    versionDescription: { en: "Initial Microsoft Teams collaboration toolset.", "zh-CN": "初始 Microsoft Teams 协作工具集。" },
    toolDescription: "Use Microsoft Graph to work with the authenticated user's Microsoft Teams resources.",
    tutorialUrl: "https://learn.microsoft.com/graph/teams-concept-overview",
    tags: ["tools", "productivity"],
    permission: []
  },
  secretSchema,
  children: [
    { id: "listTeams", name: { en: "List Joined Teams", "zh-CN": "查询已加入团队" }, description: { en: "List teams joined by the authenticated user.", "zh-CN": "查询当前用户已加入的团队。" }, toolDescription: "List joined Microsoft Teams teams.", handler: handlers.listTeams },
    { id: "listChannels", name: { en: "List Channels", "zh-CN": "查询频道" }, description: { en: "List channels in a team.", "zh-CN": "查询团队中的频道。" }, toolDescription: "List channels for a Microsoft Teams team.", handler: handlers.listChannels },
    { id: "sendChannelMessage", name: { en: "Send Channel Message", "zh-CN": "发送频道消息" }, description: { en: "Post a message to a Teams channel.", "zh-CN": "向 Teams 频道发送消息。" }, toolDescription: "Send text or HTML to a team channel.", handler: handlers.sendChannelMessage },
    { id: "sendChatMessage", name: { en: "Send Chat Message", "zh-CN": "发送聊天消息" }, description: { en: "Post a message to an existing Teams chat.", "zh-CN": "向已有 Teams 聊天发送消息。" }, toolDescription: "Send text or HTML to an existing Teams chat.", handler: handlers.sendChatMessage },
    { id: "createOnlineMeeting", name: { en: "Create Online Meeting", "zh-CN": "创建在线会议" }, description: { en: "Create a Teams online meeting.", "zh-CN": "创建 Teams 在线会议。" }, toolDescription: "Create a standalone online meeting and return its join URL.", handler: handlers.createOnlineMeeting }
  ]
});
