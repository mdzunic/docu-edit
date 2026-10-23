// The presenter session protocol (SPEC-100) — the NORMATIVE owner/follower
// contract between the presenter window (/slides/:id/presenter/:n) and the
// audience window it opens (/slides/:id/present/:n?follow=<session>). One
// owner at a time: state flows owner → follower, intents flow follower →
// owner. The owner posts `state` on every change and every
// PRESENTER_HEARTBEAT_MS; a follower writes its step machine only from an
// accepted `state` (usePlayback.setStep) and posts an `intent` for every local
// input. `hello` from either side makes the owner post one immediate `state`;
// `bye`, or PRESENTER_LOST_MS of silence, promotes a follower to owner.
//
// Everything here is plain data and pure functions — no Vue, no DOM — and the
// transport is core/windowChannel.ts. A payload from another window is
// untrusted: `parsePresenterMessage` drops (never throws on) a wrong shape, a
// foreign session or a non-finite number. The deck never travels: both windows
// read it from the library, so only positions cross the channel.

/** What the audience shows instead of the slide while blanked (B/W). */
export type PresenterBlank = "none" | "black" | "white";

export type PresenterWindowRole = "presenter" | "audience";

export interface PresenterState {
  /** Monotonic per owner; a lower or equal seq is stale. */
  seq: number;
  slideIndex: number;
  buildIndex: number;
  atEnd: boolean;
  blank: PresenterBlank;
}

export type PresenterIntentAction =
  | "advance"
  | "back"
  | "first"
  | "last"
  | "goTo"
  | "blackout"
  | "whiteout"
  | "exit";

export type PresenterMessage =
  | { kind: "state"; session: string; state: PresenterState }
  | { kind: "hello"; session: string; from: PresenterWindowRole }
  | { kind: "intent"; session: string; action: PresenterIntentAction; slideIndex?: number }
  | { kind: "bye"; session: string; from: PresenterWindowRole };

/** How often the owner re-posts its state while nothing changes. */
export const PRESENTER_HEARTBEAT_MS = 1000;

/** Silence after which a follower promotes itself to owner. */
export const PRESENTER_LOST_MS = 3000;

const BLANKS: readonly PresenterBlank[] = ["none", "black", "white"];
const ROLES: readonly PresenterWindowRole[] = ["presenter", "audience"];
const ACTIONS: readonly PresenterIntentAction[] = [
  "advance",
  "back",
  "first",
  "last",
  "goTo",
  "blackout",
  "whiteout",
  "exit",
];

/** The channel both windows of one deck's session open. */
export function presenterChannelName(deckId: string): string {
  return `docu-edit:present:${deckId}`;
}

/** The audience window's name: `window.open` with the same name re-uses the
 * window already open for this deck rather than stacking a second one. */
export function audienceWindowName(deckId: string): string {
  return `docu-edit-audience-${deckId}`;
}

/** The audience route for a session, 1-based `n`. `placed` marks a window the
 * presenter already put on another screen (it then skips the drag hint). */
export function audienceUrl(deckId: string, n: number, session: string, placed = false): string {
  const query = `follow=${encodeURIComponent(session)}${placed ? "&placed=1" : ""}`;
  return `/slides/${encodeURIComponent(deckId)}/present/${n}?${query}`;
}

/** A fresh session id: `crypto.randomUUID()` where available, else 128 random
 * bits from `getRandomValues`, else Math.random (a session id only has to be
 * unguessable enough not to collide between two decks' sessions). */
export function mintPresenterSession(): string {
  const cryptoApi = globalThis.crypto;
  if (cryptoApi && typeof cryptoApi.randomUUID === "function") {
    return cryptoApi.randomUUID();
  }
  const bytes = new Uint8Array(16);
  if (cryptoApi && typeof cryptoApi.getRandomValues === "function") {
    cryptoApi.getRandomValues(bytes);
  } else {
    for (let i = 0; i < bytes.length; i += 1) bytes[i] = Math.floor(Math.random() * 256);
  }
  return Array.from(bytes, (b) => b.toString(16).padStart(2, "0")).join("");
}

/** Whether `next` must be ignored because the follower already holds a state
 * at least as new (`prev` null = nothing received yet, so nothing is stale). */
export function isStaleState(prev: PresenterState | null, next: PresenterState): boolean {
  return prev !== null && next.seq <= prev.seq;
}

/** Whether two states show the same thing (seq aside). */
export function sameStep(a: PresenterState, b: PresenterState): boolean {
  return a.slideIndex === b.slideIndex && a.buildIndex === b.buildIndex && a.atEnd === b.atEnd && a.blank === b.blank;
}

/** The blank a B/W toggle lands on: the same key again restores the slide. */
export function toggledBlank(current: PresenterBlank, action: "blackout" | "whiteout"): PresenterBlank {
  const target: PresenterBlank = action === "blackout" ? "black" : "white";
  return current === target ? "none" : target;
}

function isRecord(raw: unknown): raw is Record<string, unknown> {
  return raw !== null && typeof raw === "object" && !Array.isArray(raw);
}

/** A finite, non-negative integer (an index or a seq). */
function isCount(raw: unknown): raw is number {
  return typeof raw === "number" && Number.isFinite(raw) && Number.isInteger(raw) && raw >= 0;
}

function parseState(raw: unknown): PresenterState | null {
  if (!isRecord(raw)) return null;
  const { seq, slideIndex, buildIndex, atEnd, blank } = raw;
  if (!isCount(seq) || !isCount(slideIndex) || !isCount(buildIndex)) return null;
  if (typeof atEnd !== "boolean" || !BLANKS.includes(blank as PresenterBlank)) return null;
  return { seq, slideIndex, buildIndex, atEnd, blank: blank as PresenterBlank };
}

/** Validates a payload received from another window. Returns null — never
 * throws — for anything that is not exactly a PresenterMessage: an unknown
 * kind, a missing or mistyped field, a non-finite or negative number, or (when
 * `session` is given) a message from any other session. The result is a fresh
 * object holding only the declared fields. */
export function parsePresenterMessage(raw: unknown, session?: string): PresenterMessage | null {
  if (!isRecord(raw)) return null;
  const sessionId = raw.session;
  if (typeof sessionId !== "string" || sessionId.length === 0) return null;
  if (session !== undefined && sessionId !== session) return null;
  switch (raw.kind) {
    case "state": {
      const state = parseState(raw.state);
      return state ? { kind: "state", session: sessionId, state } : null;
    }
    case "hello":
    case "bye": {
      const from = raw.from as PresenterWindowRole;
      if (!ROLES.includes(from)) return null;
      return raw.kind === "hello" ? { kind: "hello", session: sessionId, from } : { kind: "bye", session: sessionId, from };
    }
    case "intent": {
      const action = raw.action as PresenterIntentAction;
      if (!ACTIONS.includes(action)) return null;
      if (raw.slideIndex === undefined) {
        return action === "goTo" ? null : { kind: "intent", session: sessionId, action };
      }
      return isCount(raw.slideIndex) ? { kind: "intent", session: sessionId, action, slideIndex: raw.slideIndex } : null;
    }
    default:
      return null;
  }
}
