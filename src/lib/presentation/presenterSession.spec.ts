// SPEC-100 — the presenter session protocol: parsePresenterMessage treats
// every payload from another window as untrusted (malformed, foreign-session
// and non-finite payloads are dropped, never thrown), isStaleState, channel
// and window naming, the audience URL, and the heartbeat/lost constants.
import { describe, expect, it } from "vitest";
import {
  audienceUrl,
  audienceWindowName,
  isStaleState,
  mintPresenterSession,
  parsePresenterMessage,
  PRESENTER_HEARTBEAT_MS,
  PRESENTER_LOST_MS,
  presenterChannelName,
  sameStep,
  toggledBlank,
  type PresenterState,
} from "./presenterSession";

const S = "session-1";
const state = (over: Partial<PresenterState> = {}): PresenterState => ({
  seq: 1,
  slideIndex: 2,
  buildIndex: 0,
  atEnd: false,
  blank: "none",
  ...over,
});

describe("presenter session constants and names", () => {
  it("heartbeats every second and calls the owner lost after three", () => {
    expect(PRESENTER_HEARTBEAT_MS).toBe(1000);
    expect(PRESENTER_LOST_MS).toBe(3000);
    expect(PRESENTER_LOST_MS).toBeGreaterThan(PRESENTER_HEARTBEAT_MS);
  });

  it("names one channel per deck and one audience window per deck", () => {
    expect(presenterChannelName("deck-9")).toBe("docu-edit:present:deck-9");
    expect(audienceWindowName("deck-9")).toBe("docu-edit-audience-deck-9");
  });

  it("builds the audience route with ?follow= (and placed=1 only when placed)", () => {
    expect(audienceUrl("d1", 3, "abc")).toBe("/slides/d1/present/3?follow=abc");
    expect(audienceUrl("d1", 3, "abc", true)).toBe("/slides/d1/present/3?follow=abc&placed=1");
  });

  it("mints distinct non-empty session ids", () => {
    const a = mintPresenterSession();
    const b = mintPresenterSession();
    expect(a.length).toBeGreaterThan(8);
    expect(a).not.toBe(b);
  });
});

describe("isStaleState", () => {
  it("nothing is stale before the first state", () => {
    expect(isStaleState(null, state({ seq: 1 }))).toBe(false);
  });

  it("a lower or equal seq is stale; a higher one is not", () => {
    expect(isStaleState(state({ seq: 5 }), state({ seq: 4 }))).toBe(true);
    expect(isStaleState(state({ seq: 5 }), state({ seq: 5 }))).toBe(true);
    expect(isStaleState(state({ seq: 5 }), state({ seq: 6 }))).toBe(false);
  });
});

describe("sameStep / toggledBlank", () => {
  it("compares positions and blank, ignoring seq", () => {
    expect(sameStep(state({ seq: 1 }), state({ seq: 9 }))).toBe(true);
    expect(sameStep(state(), state({ blank: "black" }))).toBe(false);
    expect(sameStep(state(), state({ buildIndex: 1 }))).toBe(false);
  });

  it("B/W toggle their own blank and switch between the two", () => {
    expect(toggledBlank("none", "blackout")).toBe("black");
    expect(toggledBlank("black", "blackout")).toBe("none");
    expect(toggledBlank("black", "whiteout")).toBe("white");
    expect(toggledBlank("white", "whiteout")).toBe("none");
  });
});

describe("parsePresenterMessage", () => {
  it("accepts every well-formed kind and returns only the declared fields", () => {
    expect(parsePresenterMessage({ kind: "state", session: S, state: { ...state(), extra: 1 }, junk: true }, S)).toEqual({
      kind: "state",
      session: S,
      state: state(),
    });
    expect(parsePresenterMessage({ kind: "hello", session: S, from: "audience" }, S)).toEqual({
      kind: "hello",
      session: S,
      from: "audience",
    });
    expect(parsePresenterMessage({ kind: "bye", session: S, from: "presenter" }, S)).toEqual({
      kind: "bye",
      session: S,
      from: "presenter",
    });
    expect(parsePresenterMessage({ kind: "intent", session: S, action: "advance" }, S)).toEqual({
      kind: "intent",
      session: S,
      action: "advance",
    });
    expect(parsePresenterMessage({ kind: "intent", session: S, action: "goTo", slideIndex: 3 }, S)).toEqual({
      kind: "intent",
      session: S,
      action: "goTo",
      slideIndex: 3,
    });
  });

  it("accepts every intent action of the vocabulary", () => {
    for (const action of ["advance", "back", "first", "last", "blackout", "whiteout", "exit"]) {
      expect(parsePresenterMessage({ kind: "intent", session: S, action }, S)).not.toBeNull();
    }
  });

  it("drops a message from another session (or any session when unscoped is not asked)", () => {
    expect(parsePresenterMessage({ kind: "hello", session: "other", from: "audience" }, S)).toBeNull();
    // Without an expected session, the shape alone is checked.
    expect(parsePresenterMessage({ kind: "hello", session: "other", from: "audience" })).not.toBeNull();
    expect(parsePresenterMessage({ kind: "hello", session: "", from: "audience" })).toBeNull();
  });

  it("drops malformed payloads without throwing", () => {
    const bad: unknown[] = [
      null,
      undefined,
      42,
      "state",
      [],
      { kind: "state", session: S },
      { kind: "state", session: S, state: null },
      { kind: "state", session: S, state: { ...state(), atEnd: "no" } },
      { kind: "state", session: S, state: { ...state(), blank: "grey" } },
      { kind: "state", session: S, state: { ...state(), slideIndex: "2" } },
      { kind: "hello", session: S, from: "projector" },
      { kind: "bye", session: S },
      { kind: "intent", session: S, action: "jump" },
      { kind: "intent", session: S, action: "goTo" },
      { kind: "intent", session: S, action: "goTo", slideIndex: "3" },
      { kind: "nope", session: S },
      { session: S },
      { kind: "hello", from: "audience" },
      { kind: "hello", session: 7, from: "audience" },
    ];
    for (const raw of bad) {
      expect(() => parsePresenterMessage(raw, S)).not.toThrow();
      expect(parsePresenterMessage(raw, S), JSON.stringify(raw)).toBeNull();
    }
  });

  it("drops non-finite, negative and fractional numbers", () => {
    for (const value of [Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY, -1, 1.5]) {
      expect(parsePresenterMessage({ kind: "state", session: S, state: state({ seq: value }) }, S)).toBeNull();
      expect(parsePresenterMessage({ kind: "state", session: S, state: state({ slideIndex: value }) }, S)).toBeNull();
      expect(parsePresenterMessage({ kind: "state", session: S, state: state({ buildIndex: value }) }, S)).toBeNull();
      expect(parsePresenterMessage({ kind: "intent", session: S, action: "goTo", slideIndex: value }, S)).toBeNull();
    }
  });

  it("never throws on a hostile object", () => {
    const hostile = new Proxy(
      {},
      {
        get() {
          return undefined;
        },
      },
    );
    expect(parsePresenterMessage(hostile, S)).toBeNull();
  });
});
