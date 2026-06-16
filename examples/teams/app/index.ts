/**
 * Microsoft Teams echo bot — the thinnest `@copilotkit/bot-teams` example.
 *
 * It stands up the bot's `POST /api/messages` endpoint and echoes every
 * message back. No Microsoft credentials are needed: the Microsoft 365 Agents
 * Playground talks to the endpoint anonymously.
 *
 *   pnpm start        # starts the bot on http://localhost:3978/api/messages
 *   pnpm playground   # opens the Playground UI (http://localhost:56150)
 *
 * To drive a real CopilotKit agent instead of echoing, pass an `agent` to
 * `createBot` and call `thread.runAgent()` in the handler — see the README.
 */
import "dotenv/config";
import { createBot } from "@copilotkit/bot";
import { teams } from "@copilotkit/bot-teams";

const port = Number(process.env.PORT ?? 3978);

const bot = createBot({
  adapters: [teams({ port })],
});

bot.onMessage(async ({ thread, message }) => {
  await thread.post(`Echo: ${message.text}`);
});

await bot.start();

console.log(`Teams bot listening at http://localhost:${port}/api/messages`);
console.log(
  "Now run `pnpm playground` (in another terminal) and send a message.",
);
