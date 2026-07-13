# Microsoft Teams for FastGPT

A FastGPT tool suite for Microsoft Teams through the fixed Microsoft Graph v1.0 endpoint.

## Tools

- **List Joined Teams**: list teams joined by the authenticated user, with pagination.
- **List Channels**: list channels in a team, with pagination.
- **Send Channel Message**: post text or HTML to a team channel.
- **Send Chat Message**: post text or HTML to an existing Teams chat.
- **Create Online Meeting**: create a standalone Teams online meeting and return its join URL.

## Secret and permissions

`accessToken` is a Microsoft OAuth 2.0 bearer token. Grant only the delegated Microsoft Graph permissions needed by enabled tools, typically `Team.ReadBasic.All`, `Channel.ReadBasic.All`, `ChannelMessage.Send`, `ChatMessage.Send`, and `OnlineMeetings.ReadWrite`. Tenant policies may restrict online meeting creation. Never commit tokens. Tokens are sent only to `https://graph.microsoft.com/v1.0`.

## Inputs and outputs

List tools return normalized arrays and an optional Graph `nextLink`; pass that link back unchanged to retrieve the next page. Message tools require existing team/channel or chat IDs. Meeting times are ISO 8601 date-times with offsets, and the end must be after the start.

## Example

Send a channel notification with `teamId`, `channelId`, `content: "Deployment completed"`, and `contentType: "text"`.

## Local verification

```bash
corepack pnpm install
corepack pnpm test
corepack pnpm type-check
corepack pnpm build
corepack pnpm check
corepack pnpm pack
```

Tests mock Microsoft Graph and cover request construction, response parsing, pagination URL validation, schema validation, and API error handling. No live Microsoft credential integration test is included.
