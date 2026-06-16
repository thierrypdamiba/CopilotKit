// Public API for @copilotkit/bot-teams.

export { teams, TeamsAdapter } from "./adapter.js";
export type {
  TeamsAdapterOptions,
  TeamsReplyTarget,
  ConversationKey,
} from "./types.js";

export { TeamsConversationStore } from "./conversation-store.js";

export { createRunRenderer } from "./event-renderer.js";

export { renderTeamsMarkdown } from "./render/markdown.js";

export { createTeamsServer } from "./listener.js";
export type { TeamsServer, TeamsServerConfig } from "./listener.js";
