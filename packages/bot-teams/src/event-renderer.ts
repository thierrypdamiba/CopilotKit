import type { AgentSubscriber } from "@ag-ui/client";
import type {
  RunRenderer,
  CapturedToolCall,
  CapturedInterrupt,
} from "@copilotkit/bot";
import type { TeamsReplyTarget } from "./types.js";

const INTERRUPTED_SUFFIX = "\n\n_(interrupted)_";

/**
 * Build a {@link RunRenderer} for a single agent run in Teams.
 *
 * The `subscriber` is passed to `runAgent`. We accumulate each AG-UI text
 * message locally and post it in one `sendActivity` when the message ends —
 * Teams renders the final Markdown reply in the originating turn. (Live
 * token-by-token streaming via the SDK's `StreamingResponse` is a planned
 * enhancement; see the package README.) Tool calls and interrupts are captured
 * for the run-loop to read after `runAgent` resolves, exactly as the Slack
 * adapter does.
 */
export function createRunRenderer(args: {
  target: TeamsReplyTarget;
  interruptEventNames?: ReadonlySet<string>;
  /** Persist the agent's reply text to the conversation transcript. */
  recordAssistant?: (text: string) => void;
}): RunRenderer {
  const { target } = args;
  const interruptEventNames =
    args.interruptEventNames ?? new Set<string>(["on_interrupt"]);

  /** Per-AG-UI-message accumulated text. */
  const buffers = new Map<string, string>();
  const capturedToolCalls: CapturedToolCall[] = [];
  let pendingInterrupt: CapturedInterrupt | undefined;
  let aborted = false;

  const send = async (text: string): Promise<void> => {
    const trimmed = text.trim();
    if (!trimmed || !target.context) return;
    await target.context.sendActivity(trimmed);
    args.recordAssistant?.(trimmed);
  };

  const captureToolCall = (
    toolCallId: string,
    toolCallName: string,
    toolCallArgs: Record<string, unknown>,
  ): void => {
    const existing = capturedToolCalls.find((c) => c.toolCallId === toolCallId);
    if (existing) {
      existing.toolCallName = toolCallName;
      existing.toolCallArgs = toolCallArgs;
    } else {
      capturedToolCalls.push({ toolCallId, toolCallName, toolCallArgs });
    }
  };

  const subscriber: AgentSubscriber = {
    onTextMessageStartEvent({ event }) {
      if (aborted) return;
      buffers.set(event.messageId, "");
    },
    onTextMessageContentEvent({ event }) {
      if (aborted) return;
      const next = (buffers.get(event.messageId) ?? "") + (event.delta ?? "");
      buffers.set(event.messageId, next);
    },
    async onTextMessageEndEvent({ event }) {
      if (aborted) return;
      const text = buffers.get(event.messageId) ?? "";
      buffers.delete(event.messageId);
      await send(text);
    },

    onToolCallArgsEvent({ event, toolCallName, partialToolCallArgs }) {
      if (aborted) return;
      captureToolCall(
        event.toolCallId,
        toolCallName,
        (partialToolCallArgs ?? {}) as Record<string, unknown>,
      );
    },
    onToolCallEndEvent({ event, toolCallName, toolCallArgs }) {
      if (aborted) return;
      captureToolCall(
        event.toolCallId,
        toolCallName,
        (toolCallArgs ?? {}) as Record<string, unknown>,
      );
    },

    onCustomEvent({ event }) {
      if (aborted) return;
      const e = event as { name?: string; value?: unknown };
      if (!e.name || !interruptEventNames.has(e.name)) return;
      let value = e.value;
      if (typeof value === "string") {
        try {
          value = JSON.parse(value);
        } catch {
          // Leave as a string — the handler's schema rejects it explicitly.
        }
      }
      pendingInterrupt = { eventName: e.name, value };
    },

    async onRunErrorEvent({ event }) {
      if (aborted) return;
      if (!target.context) return;
      await target.context.sendActivity(
        `⚠️ Agent error: ${event.message ?? "unknown error"}`,
      );
    },
  };

  return {
    subscriber,
    getCapturedToolCalls: () => capturedToolCalls,
    getPendingInterrupt: () => pendingInterrupt,
    clearPendingInterrupt: () => {
      pendingInterrupt = undefined;
    },
    async markInterrupted() {
      if (aborted) return;
      aborted = true;
      // Flush any partial reply with a marker so the user sees the run stopped.
      const tasks: Promise<void>[] = [];
      for (const [id, buf] of Array.from(buffers.entries())) {
        if (buf.length > 0) tasks.push(send(buf + INTERRUPTED_SUFFIX));
        buffers.delete(id);
      }
      await Promise.all(tasks);
    },
  };
}
