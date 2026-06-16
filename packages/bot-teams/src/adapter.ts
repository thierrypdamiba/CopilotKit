import { CloudAdapter } from "@microsoft/agents-hosting";
import type { AuthConfiguration, TurnContext } from "@microsoft/agents-hosting";
import { ActivityTypes } from "@microsoft/agents-activity";
import type {
  PlatformAdapter,
  SurfaceCapabilities,
  IngressSink,
  InteractionEvent,
  RunRenderer,
  ReplyTarget,
  ConversationStore,
  MessageRef,
  PlatformUser,
  UserQuery,
} from "@copilotkit/bot";
import type { BotNode, ThreadMessage } from "@copilotkit/bot-ui";
import { TeamsConversationStore } from "./conversation-store.js";
import { createTeamsServer } from "./listener.js";
import type { TeamsServer } from "./listener.js";
import { createRunRenderer } from "./event-renderer.js";
import { renderTeamsMarkdown } from "./render/markdown.js";
import type { TeamsAdapterOptions, TeamsReplyTarget } from "./types.js";

/** A Teams `MessageRef` carries the live turn context so `update`/`delete` can act on it. */
interface TeamsMessageRef extends MessageRef {
  conversationKey: string;
  context?: TurnContext;
}

/**
 * Microsoft Teams `PlatformAdapter`.
 *
 * Ingress: a `CloudAdapter` receives Teams activities at `POST /api/messages`
 * (M365 Agents SDK). Egress: replies are rendered to Markdown and sent on the
 * live `TurnContext` within the originating turn — the engine awaits the whole
 * turn handler, so a reply (or a full `runAgent` loop) completes before the
 * HTTP response closes. This is the path the M365 Agents Playground exercises,
 * and it needs no Microsoft credentials locally.
 */
export class TeamsAdapter implements PlatformAdapter {
  readonly platform = "teams";
  readonly capabilities: SurfaceCapabilities;
  // Teams keeps the inbound HTTP turn open while the bot works; ~15s is the
  // practical channel window. Declarative today (the engine doesn't enforce it).
  readonly ackDeadlineMs = 15000;

  private readonly store = new TeamsConversationStore();
  private cloud: CloudAdapter | undefined;
  private server: TeamsServer | undefined;
  private sink: IngressSink | undefined;

  constructor(private readonly opts: TeamsAdapterOptions = {}) {
    this.capabilities = {
      supportsModals: false,
      supportsTyping: true,
      supportsReactions: false,
      // Replies are posted as a single message (no token-by-token streaming yet).
      supportsStreaming: false,
    };
  }

  async start(sink: IngressSink): Promise<void> {
    this.sink = sink;

    const authConfig: AuthConfiguration = {
      clientId: this.opts.clientId ?? process.env.clientId,
      clientSecret: this.opts.clientSecret ?? process.env.clientSecret,
      tenantId: this.opts.tenantId ?? process.env.tenantId,
    };
    this.cloud = new CloudAdapter(authConfig);

    this.server = createTeamsServer({
      adapter: this.cloud,
      port: this.opts.port ?? 3978,
      onTurnContext: (context) => this.handleActivity(context, sink),
    });
    await this.server.start();
  }

  async stop(): Promise<void> {
    await this.server?.stop();
  }

  /** Normalize an inbound activity and drive it into the engine via the sink. */
  private async handleActivity(
    context: TurnContext,
    sink: IngressSink,
  ): Promise<void> {
    const activity = context.activity;
    if (activity.type !== ActivityTypes.Message) return;

    // Strip any `<at>bot</at>` mention (channel/group scope) — best-effort.
    let text = "";
    try {
      text = (activity.removeRecipientMention() ?? activity.text ?? "").trim();
    } catch {
      text = (activity.text ?? "").trim();
    }

    const conversationKey = activity.conversation?.id ?? "";
    const target: TeamsReplyTarget = {
      conversationKey,
      reference: activity.getConversationReference(),
      context,
    };

    // Record the incoming message so the conversation transcript (and thus the
    // agent's history on `runAgent`) includes it.
    this.store.recordUser(conversationKey, text);

    const from = activity.from;
    const user: PlatformUser | undefined = from?.id
      ? { id: from.id, name: from.name }
      : undefined;

    await sink.onTurn({
      conversationKey,
      replyTarget: target,
      userText: text,
      user,
      platform: this.platform,
    });
  }

  render(ir: BotNode[]): string {
    return renderTeamsMarkdown(ir);
  }

  async post(target: ReplyTarget, ir: BotNode[]): Promise<MessageRef> {
    const t = target as TeamsReplyTarget;
    const text = renderTeamsMarkdown(ir);
    const id = await this.sendText(t, text);
    return { id, conversationKey: t.conversationKey, context: t.context };
  }

  async update(ref: MessageRef, ir: BotNode[]): Promise<void> {
    const r = ref as TeamsMessageRef;
    if (!r.context || !r.id) return;
    const text = renderTeamsMarkdown(ir);
    const activity = r.context.activity.clone();
    activity.id = r.id;
    activity.type = ActivityTypes.Message;
    activity.text = text;
    await r.context.updateActivity(activity);
  }

  async stream(
    target: ReplyTarget,
    chunks: AsyncIterable<string>,
  ): Promise<MessageRef> {
    const t = target as TeamsReplyTarget;
    let acc = "";
    for await (const chunk of chunks) acc += chunk;
    const id = await this.sendText(t, acc);
    return { id, conversationKey: t.conversationKey, context: t.context };
  }

  async delete(ref: MessageRef): Promise<void> {
    const r = ref as TeamsMessageRef;
    if (!r.context || !r.id) return;
    await r.context.deleteActivity(r.id);
  }

  createRunRenderer(target: ReplyTarget): RunRenderer {
    const t = target as TeamsReplyTarget;
    return createRunRenderer({
      target: t,
      interruptEventNames: this.opts.interruptEventNames,
      recordAssistant: (text) =>
        this.store.recordAssistant(t.conversationKey, text),
    });
  }

  decodeInteraction(_raw: unknown): InteractionEvent | undefined {
    // Adaptive Card actions (button/select/input) are a planned follow-up;
    // milestone-1 replies are non-interactive Markdown. See the package README.
    return undefined;
  }

  async lookupUser(_q: UserQuery): Promise<PlatformUser | undefined> {
    // Directory lookups require Microsoft Graph; not wired in milestone-1.
    return undefined;
  }

  get conversationStore(): ConversationStore {
    return this.store;
  }

  /** Return the conversation transcript the adapter has accumulated. */
  async getMessages(target: ReplyTarget): Promise<ThreadMessage[]> {
    const t = target as TeamsReplyTarget;
    return this.store.getTranscript(t.conversationKey);
  }

  /** Send plain Markdown text, preferring the live turn context. */
  private async sendText(t: TeamsReplyTarget, text: string): Promise<string> {
    const trimmed = text.trim();
    if (!trimmed) return "";
    if (t.context) {
      const res = await t.context.sendActivity(trimmed);
      return res?.id ?? "";
    }
    // Out-of-turn (proactive) send: re-enter the conversation by reference.
    if (this.cloud && t.reference) {
      let id = "";
      const appId = this.opts.clientId ?? process.env.clientId ?? "";
      await this.cloud.continueConversation(
        appId,
        t.reference as Parameters<CloudAdapter["continueConversation"]>[1],
        async (context) => {
          const res = await context.sendActivity(trimmed);
          id = res?.id ?? "";
        },
      );
      return id;
    }
    return "";
  }
}

/** Construct a Microsoft Teams `PlatformAdapter`. */
export function teams(opts: TeamsAdapterOptions = {}): TeamsAdapter {
  return new TeamsAdapter(opts);
}
