// SPEC-100 — the presenter session's owner/follower roles over an in-memory
// channel hub: the owner's state on every change plus a heartbeat, a hello
// answered with one snapshot, follower input forwarded as intents (its machine
// written only by setStep), stale-seq rejection, promotion on bye and on
// PRESENTER_LOST_MS of silence, blanking, the "complete yours too" re-ack,
// and plain (unfollowing) playback running locally with no channel.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { defineComponent, nextTick, ref } from "vue";
import { mount, type VueWrapper } from "@vue/test-utils";
import type { WindowChannel, WindowChannelTransport } from "../../lib/core/windowChannel";
import {
  PRESENTER_HEARTBEAT_MS,
  PRESENTER_LOST_MS,
  presenterChannelName,
  type PresenterMessage,
} from "../../lib/presentation/presenterSession";
import { usePlayback } from "./usePlayback";
import { usePresenterSession, type PresenterSession, type PresenterSessionDeps } from "./usePresenterSession";

/** Connects channel ends by name; a post reaches every OTHER open end, as a
 * JSON copy (what crosses a window boundary is never the sender's object). */
function channelHub() {
  const ends: Array<{ name: string; onMessage: (payload: unknown) => void; open: boolean }> = [];
  const log: Array<{ name: string; payload: PresenterMessage }> = [];
  function open(name: string, onMessage: (payload: unknown) => void): WindowChannel {
    const end = { name, onMessage, open: true };
    ends.push(end);
    return {
      post(payload) {
        if (!end.open) return;
        log.push({ name, payload: payload as PresenterMessage });
        for (const other of ends) {
          if (other !== end && other.open && other.name === name) other.onMessage(JSON.parse(JSON.stringify(payload)));
        }
      },
      close() {
        end.open = false;
      },
    };
  }
  /** Delivers a raw payload to every open end of `name`, as if another window posted it. */
  function inject(name: string, payload: unknown): void {
    for (const end of ends) if (end.open && end.name === name) end.onMessage(payload);
  }
  return { open, inject, log, ends };
}

const DECK = "deck-1";
const SESSION = "s-1";
const CHANNEL = presenterChannelName(DECK);
const wrappers: VueWrapper[] = [];

interface Harness {
  session: PresenterSession;
  playback: ReturnType<typeof usePlayback>;
  navigation: { advance: ReturnType<typeof vi.fn>; back: ReturnType<typeof vi.fn>; goTo: ReturnType<typeof vi.fn>; exit: ReturnType<typeof vi.fn> };
  hooks: { onRemoteStep: ReturnType<typeof vi.fn>; onRemoteAck: ReturnType<typeof vi.fn>; onFirstState: ReturnType<typeof vi.fn>; onPromote: ReturnType<typeof vi.fn>; onAudience: ReturnType<typeof vi.fn> };
  wrapper: VueWrapper;
}

function start(
  hub: ReturnType<typeof channelHub>,
  opts: { window: "presenter" | "audience"; session?: string; count?: number; builds?: number[]; transport?: WindowChannelTransport; consume?: boolean },
): Harness {
  const builds = opts.builds ?? [];
  const playback = usePlayback({ slideCount: ref(opts.count ?? 4), buildCountAt: (i) => builds[i] ?? 0 });
  const navigation = {
    advance: vi.fn(() => {
      if (!opts.consume) playback.advance();
    }),
    back: vi.fn(() => playback.back()),
    goTo: vi.fn((i: number) => playback.goTo(i)),
    exit: vi.fn(),
  };
  const hooks = { onRemoteStep: vi.fn(), onRemoteAck: vi.fn(), onFirstState: vi.fn(), onPromote: vi.fn(), onAudience: vi.fn() };
  let session!: PresenterSession;
  const deps: PresenterSessionDeps = {
    window: opts.window,
    deckId: ref(DECK),
    session: opts.session ?? SESSION,
    playback,
    slideCount: ref(opts.count ?? 4),
    buildCountAt: (i) => builds[i] ?? 0,
    navigation,
    ...hooks,
    openChannel: hub.open,
    transport: () => opts.transport ?? "broadcast",
  };
  const wrapper = mount(
    defineComponent({
      setup() {
        session = usePresenterSession(deps);
        return () => null;
      },
    }),
  );
  wrappers.push(wrapper);
  return { session, playback, navigation, hooks, wrapper };
}

function statesIn(hub: ReturnType<typeof channelHub>) {
  return hub.log.filter((entry) => entry.payload.kind === "state").map((entry) => entry.payload);
}

beforeEach(() => {
  vi.useFakeTimers();
});

afterEach(() => {
  for (const wrapper of wrappers.splice(0)) wrapper.unmount();
  vi.useRealTimers();
});

describe("usePresenterSession — the owner (presenter window)", () => {
  it("posts its state on mount and re-posts it every PRESENTER_HEARTBEAT_MS under the same seq", async () => {
    const hub = channelHub();
    start(hub, { window: "presenter" });
    expect(statesIn(hub)).toHaveLength(1);
    vi.advanceTimersByTime(PRESENTER_HEARTBEAT_MS * 3);
    const states = statesIn(hub);
    expect(states).toHaveLength(4);
    expect(new Set(states.map((s) => (s.kind === "state" ? s.state.seq : -1))).size).toBe(1);
    expect(hub.log.every((entry) => entry.name === CHANNEL && entry.payload.session === SESSION)).toBe(true);
  });

  it("posts a fresh state (seq + 1) on every change, including ones made outside perform", async () => {
    const hub = channelHub();
    const owner = start(hub, { window: "presenter" });
    owner.playback.goTo(2); // e.g. an auto-advance or the load landing
    await nextTick();
    const last = statesIn(hub).at(-1);
    expect(last).toMatchObject({ kind: "state", state: { seq: 2, slideIndex: 2, buildIndex: 0, atEnd: false, blank: "none" } });
  });

  it("answers an audience hello with one immediate snapshot and reports the audience", () => {
    const hub = channelHub();
    const owner = start(hub, { window: "presenter" });
    const before = statesIn(hub).length;
    hub.inject(CHANNEL, { kind: "hello", session: SESSION, from: "audience" });
    expect(statesIn(hub)).toHaveLength(before + 1);
    expect(owner.hooks.onAudience).toHaveBeenCalledWith("hello");
    hub.inject(CHANNEL, { kind: "bye", session: SESSION, from: "audience" });
    expect(owner.hooks.onAudience).toHaveBeenCalledWith("bye");
  });

  it("runs an audience intent through its own navigation and broadcasts the result", () => {
    const hub = channelHub();
    const owner = start(hub, { window: "presenter" });
    hub.inject(CHANNEL, { kind: "intent", session: SESSION, action: "advance" });
    expect(owner.navigation.advance).toHaveBeenCalledTimes(1);
    expect(owner.playback.current.value.slideIndex).toBe(1);
    expect(statesIn(hub).at(-1)).toMatchObject({ state: { slideIndex: 1 } });
    hub.inject(CHANNEL, { kind: "intent", session: SESSION, action: "goTo", slideIndex: 3 });
    expect(owner.navigation.goTo).toHaveBeenCalledWith(3);
    hub.inject(CHANNEL, { kind: "intent", session: SESSION, action: "last" });
    expect(owner.navigation.goTo).toHaveBeenLastCalledWith(3);
    hub.inject(CHANNEL, { kind: "intent", session: SESSION, action: "first" });
    expect(owner.navigation.goTo).toHaveBeenLastCalledWith(0);
    hub.inject(CHANNEL, { kind: "intent", session: SESSION, action: "exit" });
    expect(owner.navigation.exit).toHaveBeenCalled();
  });

  it("drops a foreign session's or malformed message", () => {
    const hub = channelHub();
    const owner = start(hub, { window: "presenter" });
    hub.inject(CHANNEL, { kind: "intent", session: "someone-else", action: "advance" });
    hub.inject(CHANNEL, { kind: "intent", session: SESSION, action: "goTo", slideIndex: Number.NaN });
    hub.inject(CHANNEL, "garbage");
    expect(owner.navigation.advance).not.toHaveBeenCalled();
    expect(owner.navigation.goTo).not.toHaveBeenCalled();
  });

  it("toggles the blank with B/W, and a navigation while blanked only clears it", () => {
    const hub = channelHub();
    const owner = start(hub, { window: "presenter" });
    owner.session.perform("blackout");
    expect(owner.session.blank.value).toBe("black");
    expect(statesIn(hub).at(-1)).toMatchObject({ state: { blank: "black", slideIndex: 0 } });
    owner.session.perform("whiteout");
    expect(owner.session.blank.value).toBe("white");
    owner.session.perform("advance"); // clears, does not move
    expect(owner.session.blank.value).toBe("none");
    expect(owner.navigation.advance).not.toHaveBeenCalled();
    expect(owner.playback.current.value.slideIndex).toBe(0);
    owner.session.perform("blackout");
    owner.session.perform("blackout"); // the same key again restores
    expect(owner.session.blank.value).toBe("none");
    owner.session.perform("whiteout");
    owner.session.perform("goTo", 2); // a jump clears AND jumps
    expect(owner.session.blank.value).toBe("none");
    expect(owner.playback.current.value.slideIndex).toBe(2);
  });

  it("posts one fresh state per input even when the input changed nothing (the follower's re-ack)", () => {
    const hub = channelHub();
    const owner = start(hub, { window: "presenter", consume: true });
    const before = statesIn(hub);
    owner.session.perform("advance"); // consumed (e.g. completing a build)
    const after = statesIn(hub);
    expect(after).toHaveLength(before.length + 1);
    const prev = before.at(-1);
    const next = after.at(-1);
    expect(next?.kind === "state" && prev?.kind === "state" && next.state.seq > prev.state.seq).toBe(true);
  });

  it("says bye on unmount and on page hide, once", () => {
    const hub = channelHub();
    const owner = start(hub, { window: "presenter" });
    window.dispatchEvent(new Event("pagehide"));
    owner.wrapper.unmount();
    expect(hub.log.filter((entry) => entry.payload.kind === "bye")).toEqual([
      { name: CHANNEL, payload: { kind: "bye", session: SESSION, from: "presenter" } },
    ]);
  });
});

describe("usePresenterSession — the follower (audience window)", () => {
  it("says hello, then renders the owner's snapshot through setStep", async () => {
    const hub = channelHub();
    const owner = start(hub, { window: "presenter" });
    owner.playback.goTo(2);
    await nextTick();
    const audience = start(hub, { window: "audience" });
    expect(hub.log.some((entry) => entry.payload.kind === "hello" && entry.payload.from === "audience")).toBe(true);
    expect(audience.session.following.value).toBe(true);
    expect(audience.playback.current.value.slideIndex).toBe(2);
    expect(audience.hooks.onFirstState).toHaveBeenCalledTimes(1);
  });

  it("forwards every local input as an intent and never advances its own machine", async () => {
    const hub = channelHub();
    const owner = start(hub, { window: "presenter" });
    const audience = start(hub, { window: "audience" });
    const localAdvance = vi.spyOn(audience.playback, "advance");
    audience.session.perform("advance");
    expect(audience.navigation.advance).not.toHaveBeenCalled();
    expect(localAdvance).not.toHaveBeenCalled();
    expect(owner.navigation.advance).toHaveBeenCalledTimes(1);
    expect(audience.playback.current.value.slideIndex).toBe(1); // what came back
    audience.session.perform("goTo", 3.6);
    expect(owner.navigation.goTo).toHaveBeenCalledWith(3);
    expect(audience.playback.current.value.slideIndex).toBe(3);
    audience.session.perform("blackout");
    expect(owner.session.blank.value).toBe("black");
    expect(audience.session.blank.value).toBe("black");
  });

  it("rejects a stale state (lower or equal seq)", () => {
    const hub = channelHub();
    const audience = start(hub, { window: "audience" });
    const send = (seq: number, slideIndex: number) =>
      hub.inject(CHANNEL, { kind: "state", session: SESSION, state: { seq, slideIndex, buildIndex: 0, atEnd: false, blank: "none" } });
    send(5, 2);
    expect(audience.playback.current.value.slideIndex).toBe(2);
    send(4, 0);
    send(5, 3);
    expect(audience.playback.current.value.slideIndex).toBe(2);
    send(6, 3);
    expect(audience.playback.current.value.slideIndex).toBe(3);
  });

  it("tells a one-step advance (forward → transition) from a jump or a back", () => {
    const hub = channelHub();
    const owner = start(hub, { window: "presenter", builds: [0, 1, 0, 0] });
    const audience = start(hub, { window: "audience", builds: [0, 1, 0, 0] });
    owner.session.perform("advance"); // 0 → 1
    expect(audience.hooks.onRemoteStep).toHaveBeenLastCalledWith(0, true);
    owner.session.perform("advance"); // build on slide 1: not a slide step
    expect(audience.hooks.onRemoteStep).toHaveBeenLastCalledWith(1, false);
    owner.session.perform("advance"); // 1 (fully built) → 2
    expect(audience.hooks.onRemoteStep).toHaveBeenLastCalledWith(1, true);
    owner.session.perform("back");
    expect(audience.hooks.onRemoteStep).toHaveBeenLastCalledWith(2, false);
    owner.session.perform("goTo", 3);
    expect(audience.hooks.onRemoteStep).toHaveBeenLastCalledWith(1, false);
  });

  it("completes its own in-flight build when the owner re-posts an unchanged state", () => {
    const hub = channelHub();
    const owner = start(hub, { window: "presenter", consume: true });
    const audience = start(hub, { window: "audience" });
    owner.session.perform("advance"); // consumed by the owner's build completion
    expect(audience.hooks.onRemoteAck).toHaveBeenCalledTimes(1);
    vi.advanceTimersByTime(PRESENTER_HEARTBEAT_MS); // a heartbeat is NOT an ack
    expect(audience.hooks.onRemoteAck).toHaveBeenCalledTimes(1);
  });

  it("promotes itself on the presenter's bye — a role flip on the same machine", () => {
    const hub = channelHub();
    const owner = start(hub, { window: "presenter" });
    const audience = start(hub, { window: "audience" });
    owner.session.perform("goTo", 2);
    owner.wrapper.unmount(); // closed or navigated away: bye
    expect(audience.session.following.value).toBe(false);
    expect(audience.session.role.value).toBe("owner");
    expect(audience.hooks.onPromote).toHaveBeenCalledTimes(1);
    expect(audience.playback.current.value.slideIndex).toBe(2);
    audience.session.perform("advance"); // now plain playback: runs locally
    expect(audience.navigation.advance).toHaveBeenCalledTimes(1);
    expect(audience.playback.current.value.slideIndex).toBe(3);
    // It said bye itself, so a presenter that is merely slow sees it go.
    expect(hub.log.at(-1)?.payload).toEqual({ kind: "bye", session: SESSION, from: "audience" });
  });

  it("promotes itself after PRESENTER_LOST_MS of silence — so a stale ?follow= URL works alone", () => {
    const hub = channelHub();
    const audience = start(hub, { window: "audience" });
    vi.advanceTimersByTime(PRESENTER_LOST_MS - 1);
    expect(audience.session.following.value).toBe(true);
    vi.advanceTimersByTime(1);
    expect(audience.session.following.value).toBe(false);
    expect(audience.hooks.onPromote).toHaveBeenCalledTimes(1);
    audience.session.perform("advance");
    expect(audience.navigation.advance).toHaveBeenCalled();
  });

  it("stays a follower while heartbeats keep arriving", () => {
    const hub = channelHub();
    start(hub, { window: "presenter" });
    const audience = start(hub, { window: "audience" });
    vi.advanceTimersByTime(PRESENTER_LOST_MS * 4);
    expect(audience.session.following.value).toBe(true);
  });

  it("re-syncs a reloaded audience from its hello with one full snapshot", async () => {
    const hub = channelHub();
    const owner = start(hub, { window: "presenter" });
    const first = start(hub, { window: "audience" });
    owner.session.perform("goTo", 3);
    owner.session.perform("blackout");
    first.wrapper.unmount(); // reload: bye…
    expect(owner.hooks.onAudience).toHaveBeenLastCalledWith("bye");
    const back = start(hub, { window: "audience" }); // …and hello
    expect(owner.hooks.onAudience).toHaveBeenLastCalledWith("hello");
    expect(back.playback.current.value.slideIndex).toBe(3);
    expect(back.session.blank.value).toBe("black");
    expect(owner.playback.current.value.slideIndex).toBe(3); // the presenter kept its state
  });

  it("re-applies its last state after the view reloads the deck", () => {
    const hub = channelHub();
    const owner = start(hub, { window: "presenter" });
    const audience = start(hub, { window: "audience" });
    owner.session.perform("goTo", 2);
    audience.playback.setStep({ slideIndex: 0, buildIndex: 0 }); // the view's load re-lands on its route
    audience.session.reapply();
    expect(audience.playback.current.value.slideIndex).toBe(2);
  });

  it("with no transport, promotes on schedule rather than waiting forever", () => {
    const hub = channelHub();
    const audience = start(hub, { window: "audience", transport: "none" });
    expect(hub.ends).toHaveLength(0);
    vi.advanceTimersByTime(PRESENTER_LOST_MS);
    expect(audience.session.following.value).toBe(false);
  });
});

describe("usePresenterSession — plain playback (no ?follow=)", () => {
  it("opens no channel and runs every input locally", () => {
    const hub = channelHub();
    const plain = start(hub, { window: "audience", session: "" });
    expect(hub.ends).toHaveLength(0);
    expect(plain.session.following.value).toBe(false);
    expect(plain.session.isAudience).toBe(false);
    plain.session.perform("advance");
    expect(plain.navigation.advance).toHaveBeenCalledTimes(1);
    plain.session.perform("whiteout");
    expect(plain.session.blank.value).toBe("white");
    vi.advanceTimersByTime(PRESENTER_LOST_MS * 2);
    expect(plain.hooks.onPromote).not.toHaveBeenCalled();
  });
});
