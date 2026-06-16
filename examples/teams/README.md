# Teams example — echo bot

The thinnest runnable demo of [`@copilotkit/bot-teams`](../../packages/bot-teams):
a Microsoft Teams bot that echoes your messages, testable locally in the
**Microsoft 365 Agents Playground** with **no Microsoft credentials**.

## Run it

From this directory (after `pnpm install` at the repo root):

```sh
pnpm start        # starts the bot on http://localhost:3978/api/messages
```

In a second terminal:

```sh
pnpm playground   # opens the M365 Agents Playground at http://localhost:56150
```

Type `Hello from Teams` in the Playground. You should see `Echo: Hello from
Teams` come back. That confirms the CopilotKit bot engine and the Teams adapter
are working end-to-end.

## What's in here

- `app/index.ts` — the entire bot: `createBot({ adapters: [teams()] })` plus an
  `onMessage` handler that echoes. This is user-land code, not SDK code.

## Drive a real agent instead of echoing

Pass an `agent` to `createBot` and call `thread.runAgent()`:

```ts
import { createBot } from "@copilotkit/bot";
import { teams } from "@copilotkit/bot-teams";
import { HttpAgent } from "@ag-ui/client";

const bot = createBot({
  adapters: [teams({ port: 3978 })],
  agent: (threadId) => {
    const a = new HttpAgent({ url: process.env.AGENT_URL! });
    a.threadId = threadId;
    return a;
  },
});

bot.onMessage(({ thread }) => thread.runAgent());

await bot.start();
```

## Real Teams

The Playground needs no credentials. To sideload into actual Microsoft Teams
(tunnel + Entra app registration + Azure Bot resource + manifest), follow the
[Microsoft Teams guide](../../showcase/shell-docs/src/content/docs/microsoft-teams.mdx).
