# @copilotkit/bot-teams

The **Microsoft Teams platform adapter** for [`@copilotkit/bot`](../bot) — a
concrete `PlatformAdapter` that plugs Teams into the platform-agnostic bot
engine, exactly like [`@copilotkit/bot-slack`](../bot-slack) does for Slack. You
write your bot once with `createBot` (handlers, JSX, tools, context) and run it
on Teams by adding this adapter.

It is built on the **Microsoft 365 Agents SDK** (`@microsoft/agents-hosting`),
the successor to the Bot Framework SDK.

## Install

```sh
pnpm add @copilotkit/bot @copilotkit/bot-ui @copilotkit/bot-teams
```

## Quickstart

```ts
import { createBot } from "@copilotkit/bot";
import { teams } from "@copilotkit/bot-teams";

const bot = createBot({
  adapters: [teams({ port: 3978 })],
});

bot.onMessage(({ thread, message }) => thread.post(`Echo: ${message.text}`));

await bot.start(); // POST /api/messages now listening on :3978
```

Then point the **Microsoft 365 Agents Playground** at it — no Microsoft
credentials required for local development:

```sh
npx @microsoft/m365agentsplayground   # opens http://localhost:56150
```

The Playground connects to `http://127.0.0.1:3978/api/messages` and gives you a
Teams-like chat UI to test against. See [`examples/teams`](../../examples/teams)
for a complete, runnable echo bot, and the
[Microsoft Teams guide](../../showcase/shell-docs/src/content/docs/microsoft-teams.mdx)
for sideloading into real Teams via Azure Bot Service.

## How it maps onto the `PlatformAdapter` contract

- **Ingress** — a `CloudAdapter` receives Teams activities at
  `POST /api/messages` (stood up by an Express server). Each `message` activity
  is normalized into `sink.onTurn(...)`.
- **Egress** — replies are rendered to Markdown and sent on the live
  `TurnContext` _within the originating turn_. The bot engine awaits the whole
  turn handler, so a reply — or a full `runAgent()` tool loop — completes before
  the HTTP response closes. (Out-of-turn / proactive sends fall back to
  `CloudAdapter.continueConversation` via the captured conversation reference.)
- **Agent runs** — `createRunRenderer` bridges AG-UI events to Teams: it
  accumulates each text message and posts it with `sendActivity`, and captures
  tool calls and interrupts for the run loop.
- **History** — Teams does not hand the bot a queryable transcript, so an
  in-memory `TeamsConversationStore` keeps one per conversation and seeds each
  agent run with it. Swap in a durable `ConversationStore` for production.

## Options

```ts
teams({
  port: 3978, // POST /api/messages port (Playground default)
  clientId, // Microsoft app id — omit for anonymous local dev
  clientSecret, // omit for anonymous local dev
  tenantId, // omit for multi-tenant / anonymous
  interruptEventNames, // custom-event names treated as agent interrupts
});
```

Credentials also resolve from the `clientId` / `clientSecret` / `tenantId`
environment variables (the names the M365 Agents SDK reads).

## Status & roadmap

This is the first vertical slice — enough to run a real `createBot` bot on Teams
and verify it in the M365 Agents Playground. Implemented: message ingress,
Markdown egress, `runAgent` streaming (posted as a single message), tool-call /
interrupt capture, conversation history, `update`/`delete`.

Planned follow-ups (the architecture leaves room for each):

- **Adaptive Cards** — render the interactive bot-ui vocabulary (`<Button>`,
  `<Select>`, `<Input>`, tables) to Adaptive Cards, and decode card actions in
  `decodeInteraction` (currently non-interactive Markdown only).
- **Live streaming** — token-by-token replies via the SDK's `StreamingResponse`
  (`queueInformativeUpdate` / `queueTextChunk` / `endStream`).
- **User lookup** — directory resolution via Microsoft Graph.
- **File upload** — attachments via the Teams/Graph file APIs.

## Exports

`teams`, `TeamsAdapter`, `TeamsAdapterOptions`, `TeamsReplyTarget`,
`ConversationKey`; `TeamsConversationStore`; `createRunRenderer`;
`renderTeamsMarkdown`; `createTeamsServer` / `TeamsServer` / `TeamsServerConfig`.
