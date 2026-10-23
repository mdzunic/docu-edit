// The presenter session (SPEC-100) — one composable for both windows of a
// presenter-view pair, on SPEC-100's normative protocol
// (lib/presentation/presenterSession.ts) over core/windowChannel.ts.
//
//   · The presenter window (/slides/:id/presenter/:n) is the OWNER of the
//     session it minted: its single usePlayback machine is the show. It posts
//     `state` on every change and every PRESENTER_HEARTBEAT_MS, answers an
//     audience `hello` with one immediate snapshot, and applies the audience's
//     `intent`s through its own step functions.
//   · The audience window (/slides/:id/present/:n?follow=<session>) is a
//     FOLLOWER: every local input becomes an `intent` (`perform` posts it
//     instead of running it) and its machine is written only by an accepted
//     `state`, through usePlayback.setStep — so the view's own
//     watch(playback.current) still drives builds, and `onRemoteStep` lets it
//     stage the same transition an advance would have played.
//   · `bye` from the presenter, or PRESENTER_LOST_MS without a word from it,
//     promotes the follower: a role flip on the same machine, never a second
//     one. It says `bye` itself (so a presenter that is merely slow shows its
//     audience as lost and offers to reopen it), closes the channel and carries
//     on as plain SPEC-084 playback.
//   · Without ?follow= the play view is plain playback: an owner with no
//     channel, so `perform` just runs the input locally.
//
// Blanking (B/W) is session state, not navigation: `perform` toggles it, any
// navigation clears it — an advance or back that lands while blanked only
// clears it, so the room sees the slide it was looking at — and it travels in
// PresenterState.blank. Never persisted. Each window reads its own
// prefers-reduced-motion; nothing about motion crosses the channel.
import { computed, onBeforeUnmount, onMounted, ref, watch, type Ref } from "vue";
import {
  openWindowChannel,
  windowChannelTransport,
  type WindowChannel,
  type WindowChannelTransport,
} from "../../lib/core/windowChannel";
import {
  isStaleState,
  parsePresenterMessage,
  PRESENTER_HEARTBEAT_MS,
  PRESENTER_LOST_MS,
  presenterChannelName,
  toggledBlank,
  type PresenterBlank,
  type PresenterIntentAction,
  type PresenterMessage,
  type PresenterState,
  type PresenterWindowRole,
} from "../../lib/presentation/presenterSession";
import type { UsePlayback } from "./usePlayback";

/** The owning window's own step functions — the view's, so a remote input
 * completes builds and stages transitions exactly as a local one does. */
export interface PresenterNavigation {
  advance: () => void;
  back: () => void;
  goTo: (slideIndex: number) => void;
  exit: () => void;
}

export interface PresenterSessionDeps {
  /** Which window of the pair this is. */
  window: PresenterWindowRole;
  deckId: Ref<string>;
  /** The session: minted by the presenter; the audience's ?follow= value
   * ("" = not following — plain playback, no channel). */
  session: string;
  playback: UsePlayback;
  slideCount: Ref<number>;
  navigation: PresenterNavigation;
  /** The machine's own slide facts, so a follower can tell a one-step
   * advance (which plays a transition) from a jump (which does not). */
  buildCountAt?: (slideIndex: number) => number;
  isHidden?: (slideIndex: number) => boolean;
  /** Follower: an accepted state moved the show; called after setStep with
   * the slide it left and whether the move is exactly what advance() does. */
  onRemoteStep?: (fromIndex: number, forward: boolean) => void;
  /** Follower: the owner re-posted an unchanged state — it consumed an input
   * to complete an in-flight build, so this window completes its own. */
  onRemoteAck?: () => void;
  /** Follower: the first state arrived (the view re-reads the store once). */
  onFirstState?: () => void;
  /** Follower: promoted to owner. */
  onPromote?: () => void;
  /** Owner: an audience window said hello (joined) or bye (left). */
  onAudience?: (event: "hello" | "bye") => void;
  /** Test seams (default: core/windowChannel over the real browser). */
  openChannel?: (name: string, onMessage: (payload: unknown) => void) => WindowChannel;
  transport?: () => WindowChannelTransport;
}

export type PresenterSessionRole = "owner" | "follower";

export function usePresenterSession(deps: PresenterSessionDeps) {
  const following = deps.window === "audience" && deps.session.length > 0;
  const role = ref<PresenterSessionRole>(following ? "follower" : "owner");
  /** The blank screen (B/W): view state only, never persisted. */
  const blank = ref<PresenterBlank>("none");
  const transport = deps.window === "presenter" || following ? (deps.transport ?? windowChannelTransport)() : "none";
  let channel: WindowChannel | null = null;
  let left = false;

  // --- owner: state out --------------------------------------------------------
  let seq = 0;
  let posted: PresenterState | null = null;
  let heartbeat = 0;

  function snapshot(): PresenterState {
    const { slideIndex, buildIndex } = deps.playback.current.value;
    return { seq, slideIndex, buildIndex, atEnd: deps.playback.atEnd.value, blank: blank.value };
  }
  function differsFromPosted(): boolean {
    const now = snapshot();
    return (
      !posted ||
      now.slideIndex !== posted.slideIndex ||
      now.buildIndex !== posted.buildIndex ||
      now.atEnd !== posted.atEnd ||
      now.blank !== posted.blank
    );
  }
  function send(message: PresenterMessage): void {
    channel?.post(message);
  }
  /** Posts the current state — under a new seq (`fresh`), or the last one
   * again (a heartbeat, or the snapshot a hello asks for). */
  function postState(fresh: boolean): void {
    if (!channel || role.value !== "owner") return;
    if (fresh || !posted) {
      seq += 1;
      posted = snapshot();
    }
    send({ kind: "state", session: deps.session, state: posted });
  }
  // Changes that did not come through `perform` (an auto-advance, the load
  // landing) post here; `perform` posts its own one fresh state per input.
  watch([() => deps.playback.current.value, () => deps.playback.atEnd.value, () => blank.value], () => {
    if (differsFromPosted()) postState(true);
  });

  // --- input ------------------------------------------------------------------
  function run(action: PresenterIntentAction, slideIndex?: number): void {
    if (action === "blackout" || action === "whiteout") {
      blank.value = toggledBlank(blank.value, action);
      return;
    }
    if (action === "exit") {
      deps.navigation.exit();
      return;
    }
    if (blank.value !== "none") {
      blank.value = "none";
      if (action === "advance" || action === "back") return;
    }
    if (action === "advance") deps.navigation.advance();
    else if (action === "back") deps.navigation.back();
    else if (action === "first") deps.navigation.goTo(0);
    else if (action === "last") deps.navigation.goTo(deps.slideCount.value - 1);
    else deps.navigation.goTo(slideIndex ?? deps.playback.current.value.slideIndex);
  }

  /** Every local input enters here. A follower forwards it as an intent and
   * waits for the state that comes back; an owner runs it and posts exactly
   * one fresh state — the change, or (when the input was consumed completing
   * a build) an unchanged one the follower reads as "complete yours too". */
  function perform(action: PresenterIntentAction, slideIndex?: number): void {
    if (role.value === "follower") {
      const target = slideIndex !== undefined && Number.isFinite(slideIndex) ? Math.max(0, Math.trunc(slideIndex)) : undefined;
      send({ kind: "intent", session: deps.session, action, ...(target === undefined ? {} : { slideIndex: target }) });
      return;
    }
    run(action, slideIndex);
    postState(true);
  }

  // --- follower: state in -----------------------------------------------------
  let received: PresenterState | null = null;
  let lostTimer = 0;

  function armLost(): void {
    if (lostTimer) clearTimeout(lostTimer);
    lostTimer = window.setTimeout(promote, PRESENTER_LOST_MS);
  }

  function isForward(from: { slideIndex: number; buildIndex: number }, fromEnd: boolean, next: PresenterState): boolean {
    if (fromEnd || next.atEnd || next.buildIndex !== 0 || next.slideIndex <= from.slideIndex) return false;
    if (from.buildIndex < (deps.buildCountAt ? deps.buildCountAt(from.slideIndex) : 0)) return false;
    for (let index = from.slideIndex + 1; index < next.slideIndex; index += 1) {
      if (!deps.isHidden?.(index)) return false;
    }
    return true;
  }

  function applyState(next: PresenterState): void {
    if (isStaleState(received, next)) return;
    const first = received === null;
    received = next;
    const from = { ...deps.playback.current.value };
    const fromEnd = deps.playback.atEnd.value;
    const blankChanged = blank.value !== next.blank;
    blank.value = next.blank;
    if (from.slideIndex === next.slideIndex && from.buildIndex === next.buildIndex && fromEnd === next.atEnd) {
      if (!blankChanged && !first) deps.onRemoteAck?.();
    } else {
      deps.playback.setStep({ slideIndex: next.slideIndex, buildIndex: next.buildIndex }, next.atEnd);
      deps.onRemoteStep?.(from.slideIndex, isForward(from, fromEnd, next));
    }
    if (first) deps.onFirstState?.();
  }

  /** Re-applies the last accepted state (after the view reloaded the deck and
   * re-landed on its route's slide). */
  function reapply(): void {
    if (role.value !== "follower" || !received) return;
    const last = received;
    deps.playback.setStep({ slideIndex: last.slideIndex, buildIndex: last.buildIndex }, last.atEnd);
    blank.value = last.blank;
  }

  function promote(): void {
    if (role.value === "owner") return;
    if (lostTimer) clearTimeout(lostTimer);
    lostTimer = 0;
    role.value = "owner";
    send({ kind: "bye", session: deps.session, from: "audience" });
    channel?.close();
    channel = null;
    deps.onPromote?.();
  }

  // --- the channel ------------------------------------------------------------
  function onMessage(raw: unknown): void {
    const message = parsePresenterMessage(raw, deps.session);
    if (!message) return;
    if (role.value === "follower") {
      if (message.kind === "state") {
        armLost();
        applyState(message.state);
      } else if (message.kind === "bye" && message.from === "presenter") {
        promote();
      } else if (message.kind === "hello" && message.from === "presenter") {
        armLost();
      }
      return;
    }
    if (message.kind === "intent") {
      perform(message.action, message.slideIndex);
    } else if (message.kind === "hello" && message.from === "audience") {
      postState(false);
      deps.onAudience?.("hello");
    } else if (message.kind === "bye" && message.from === "audience") {
      deps.onAudience?.("bye");
    }
  }

  function start(): void {
    if (left || channel || transport === "none" || (deps.window === "audience" && !following)) {
      if (following) armLost(); // no channel can ever answer — promote on schedule
      return;
    }
    const open = deps.openChannel ?? ((name, handler) => openWindowChannel(name, handler));
    channel = open(presenterChannelName(deps.deckId.value), onMessage);
    if (role.value === "follower") {
      send({ kind: "hello", session: deps.session, from: "audience" });
      armLost();
    } else {
      postState(true);
      heartbeat = window.setInterval(() => postState(false), PRESENTER_HEARTBEAT_MS);
    }
  }

  /** Says bye and closes the channel (unmount, page hide). Idempotent. */
  function leave(): void {
    if (left) return;
    left = true;
    if (heartbeat) clearInterval(heartbeat);
    if (lostTimer) clearTimeout(lostTimer);
    heartbeat = 0;
    lostTimer = 0;
    if (channel) {
      send({ kind: "bye", session: deps.session, from: deps.window });
      channel.close();
      channel = null;
    }
  }

  onMounted(() => {
    start();
    window.addEventListener("pagehide", leave);
  });
  onBeforeUnmount(() => {
    window.removeEventListener("pagehide", leave);
    leave();
  });

  return {
    role,
    blank,
    /** Whether this window currently follows another (false once promoted). */
    following: computed(() => role.value === "follower"),
    /** Whether this window was opened as an audience window at all. */
    isAudience: following,
    transport,
    session: deps.session,
    perform,
    reapply,
    leave,
  };
}

export type PresenterSession = ReturnType<typeof usePresenterSession>;
