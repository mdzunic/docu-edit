// SPEC-100 — the presenter timer: pure and clock-injected. Start, pause,
// resume and reset accumulate against the `nowMs` each call is given (no test
// touches the real clock), and formatElapsed switches from M:SS to H:MM:SS at
// one hour.
import { describe, expect, it } from "vitest";
import {
  COUNTDOWN_MINUTES_DEFAULT,
  COUNTDOWN_MINUTES_MAX,
  COUNTDOWN_MINUTES_MIN,
  elapsedMs,
  formatElapsed,
  IDLE_TIMER,
  isCountdownMinutes,
  isTimerRunning,
  parseCountdownMinutes,
  parseTimerState,
  pauseTimer,
  resetTimer,
  showTimerRecordKey,
  startTimer,
  stopTimer,
  TIMER_MODES,
  timerPhase,
  timerReadout,
  toggleTimer,
  type TimerMode,
} from "./presenterTimer";

describe("presenter timer state", () => {
  it("is idle at zero until started", () => {
    expect(isTimerRunning(IDLE_TIMER)).toBe(false);
    expect(elapsedMs(IDLE_TIMER, 99_999)).toBe(0);
  });

  it("runs from the start instant", () => {
    const running = startTimer(IDLE_TIMER, 1_000);
    expect(isTimerRunning(running)).toBe(true);
    expect(elapsedMs(running, 1_000)).toBe(0);
    expect(elapsedMs(running, 6_500)).toBe(5_500);
  });

  it("pausing banks the stretch and freezes the readout", () => {
    const paused = pauseTimer(startTimer(IDLE_TIMER, 0), 4_000);
    expect(isTimerRunning(paused)).toBe(false);
    expect(paused.accumulatedMs).toBe(4_000);
    expect(elapsedMs(paused, 60_000)).toBe(4_000);
  });

  it("resuming continues from the banked time", () => {
    let t = startTimer(IDLE_TIMER, 0);
    t = pauseTimer(t, 10_000); // 10s
    t = startTimer(t, 50_000); // paused for 40s
    expect(elapsedMs(t, 53_000)).toBe(13_000);
    t = pauseTimer(t, 55_000);
    expect(elapsedMs(t, 99_000)).toBe(15_000);
  });

  it("start on a running timer and pause on a paused one change nothing", () => {
    const running = startTimer(IDLE_TIMER, 100);
    expect(startTimer(running, 5_000)).toBe(running);
    const paused = pauseTimer(running, 200);
    expect(pauseTimer(paused, 9_000)).toBe(paused);
  });

  it("reset zeroes it, keeping whether it runs", () => {
    const running = startTimer(IDLE_TIMER, 0);
    const restarted = resetTimer(running, 30_000);
    expect(isTimerRunning(restarted)).toBe(true);
    expect(elapsedMs(restarted, 30_000)).toBe(0);
    expect(elapsedMs(restarted, 31_000)).toBe(1_000);
    const paused = resetTimer(pauseTimer(running, 7_000), 8_000);
    expect(isTimerRunning(paused)).toBe(false);
    expect(elapsedMs(paused, 20_000)).toBe(0);
  });

  it("a clock stepping backwards reads as no time passing", () => {
    expect(elapsedMs(startTimer(IDLE_TIMER, 10_000), 9_000)).toBe(0);
  });
});

describe("formatElapsed", () => {
  it("reads M:SS below an hour", () => {
    expect(formatElapsed(0)).toBe("0:00");
    expect(formatElapsed(999)).toBe("0:00");
    expect(formatElapsed(1_000)).toBe("0:01");
    expect(formatElapsed(65_000)).toBe("1:05");
    expect(formatElapsed(59 * 60_000 + 59_999)).toBe("59:59");
  });

  it("reads H:MM:SS at and above an hour", () => {
    expect(formatElapsed(3_600_000)).toBe("1:00:00");
    expect(formatElapsed(3_600_000 + 5 * 60_000 + 7_000)).toBe("1:05:07");
    expect(formatElapsed(12 * 3_600_000)).toBe("12:00:00");
  });

  it("reads 0:00 for negative or non-finite input", () => {
    expect(formatElapsed(-5_000)).toBe("0:00");
    expect(formatElapsed(Number.NaN)).toBe("0:00");
    expect(formatElapsed(Number.POSITIVE_INFINITY)).toBe("0:00");
  });
});

// SPEC-146 — the show timer's phases, modes, talk length and readout, and the
// validator a TimerState from another window or from sessionStorage passes.
describe("timerPhase", () => {
  it("reads stopped, running and paused from the two numbers", () => {
    expect(timerPhase(IDLE_TIMER)).toBe("stopped");
    const running = startTimer(IDLE_TIMER, 1_000);
    expect(timerPhase(running)).toBe("running");
    expect(timerPhase(startTimer(IDLE_TIMER, 0))).toBe("running"); // a stretch begun at 0 still runs
    expect(timerPhase(pauseTimer(running, 4_000))).toBe("paused");
  });

  it("a 0 ms pause banks nothing and reads as stopped (both show the start)", () => {
    expect(timerPhase(pauseTimer(startTimer(IDLE_TIMER, 5_000), 5_000))).toBe("stopped");
  });
});

describe("toggleTimer / stopTimer", () => {
  it("toggles stopped → running → paused → running, the paused stretch not counted", () => {
    let t = toggleTimer(IDLE_TIMER, 1_000);
    expect(timerPhase(t)).toBe("running");
    expect(elapsedMs(t, 1_000)).toBe(0);
    t = toggleTimer(t, 11_000);
    expect(timerPhase(t)).toBe("paused");
    expect(elapsedMs(t, 99_000)).toBe(10_000);
    t = toggleTimer(t, 50_000);
    expect(timerPhase(t)).toBe("running");
    expect(elapsedMs(t, 53_000)).toBe(13_000);
  });

  it("starts a stopped timer from zero", () => {
    expect(toggleTimer(IDLE_TIMER, 7_000)).toEqual({ startedAtMs: 7_000, accumulatedMs: 0 });
  });

  it("stops from every phase to a fresh IDLE_TIMER", () => {
    const running = startTimer(IDLE_TIMER, 0);
    const paused = pauseTimer(running, 3_000);
    for (const state of [IDLE_TIMER, running, paused]) {
      const stopped = stopTimer();
      expect(stopped).toEqual(IDLE_TIMER);
      expect(stopped).not.toBe(IDLE_TIMER);
      expect(timerPhase(stopped)).toBe("stopped");
      expect(elapsedMs(stopped, 60_000), JSON.stringify(state)).toBe(0);
    }
  });
});

/** The readout with `elapsed` banked (a 2-minute countdown unless told otherwise). */
function at(elapsed: number, mode: TimerMode = "countdown", lengthMs = 120_000) {
  return timerReadout({ startedAtMs: null, accumulatedMs: elapsed }, 0, mode, lengthMs);
}

describe("timerReadout", () => {
  it("matches every row of the countdown boundary table (2 minutes)", () => {
    const rows: Array<[number, string, boolean]> = [
      [0, "2:00", false],
      [1, "2:00", false],
      [1_000, "1:59", false],
      [75_000, "0:45", false],
      [119_999, "0:01", false],
      [120_000, "0:00", true],
      [120_999, "0:00", true],
      [121_000, "−0:01", true],
      [162_000, "−0:42", true],
      [120_000 + 3_661_000, "−1:01:01", true],
    ];
    for (const [elapsed, text, overtime] of rows) {
      expect(at(elapsed), String(elapsed)).toEqual({ text, overtime });
    }
  });

  it("uses U+2212, never a hyphen, in overtime", () => {
    expect(at(162_000).text.startsWith("−")).toBe(true);
    expect(at(162_000).text).not.toContain("-");
  });

  it("reads a 240-minute countdown with nothing elapsed as 4:00:00", () => {
    expect(at(0, "countdown", COUNTDOWN_MINUTES_MAX * 60_000)).toEqual({ text: "4:00:00", overtime: false });
  });

  it("stopwatch reads formatElapsed and is never overtime, whatever the length", () => {
    expect(at(0, "stopwatch")).toEqual({ text: "0:00", overtime: false });
    expect(at(75_000, "stopwatch")).toEqual({ text: "1:15", overtime: false });
    expect(at(162_000, "stopwatch")).toEqual({ text: "2:42", overtime: false });
    expect(at(3_661_000, "stopwatch", 60_000)).toEqual({ text: "1:01:01", overtime: false });
  });

  it("reads a running timer at nowMs", () => {
    const running = startTimer(IDLE_TIMER, 10_000);
    expect(timerReadout(running, 85_000, "countdown", 120_000)).toEqual({ text: "0:45", overtime: false });
    expect(timerReadout(running, 85_000, "stopwatch", 120_000)).toEqual({ text: "1:15", overtime: false });
  });
});

describe("talk length", () => {
  it("defaults to 20 minutes within 1–240, over the two modes", () => {
    expect(COUNTDOWN_MINUTES_DEFAULT).toBe(20);
    expect(COUNTDOWN_MINUTES_MIN).toBe(1);
    expect(COUNTDOWN_MINUTES_MAX).toBe(240);
    expect(TIMER_MODES).toEqual(["stopwatch", "countdown"]);
  });

  it("isCountdownMinutes accepts whole minutes in range only", () => {
    expect(isCountdownMinutes(1)).toBe(true);
    expect(isCountdownMinutes(240)).toBe(true);
    for (const bad of [0, 241, 2.5, Number.NaN, "20", null, undefined]) {
      expect(isCountdownMinutes(bad), String(bad)).toBe(false);
    }
  });

  it("parseCountdownMinutes rounds and clamps what is typed; non-numbers are null", () => {
    expect(parseCountdownMinutes("25")).toBe(25);
    expect(parseCountdownMinutes(" 7 ")).toBe(7);
    expect(parseCountdownMinutes("2.6")).toBe(3);
    expect(parseCountdownMinutes("0")).toBe(1);
    expect(parseCountdownMinutes("999")).toBe(240);
    expect(parseCountdownMinutes("")).toBeNull();
    expect(parseCountdownMinutes("   ")).toBeNull();
    expect(parseCountdownMinutes("abc")).toBeNull();
  });
});

describe("parseTimerState", () => {
  it("accepts a running and a paused (null start) state", () => {
    expect(parseTimerState({ startedAtMs: 1_000, accumulatedMs: 0 })).toEqual({ startedAtMs: 1_000, accumulatedMs: 0 });
    expect(parseTimerState({ startedAtMs: null, accumulatedMs: 75_000 })).toEqual({ startedAtMs: null, accumulatedMs: 75_000 });
  });

  it("drops extra fields", () => {
    const parsed = parseTimerState({ startedAtMs: null, accumulatedMs: 5, mode: "countdown", seq: 3 });
    expect(parsed).toEqual({ startedAtMs: null, accumulatedMs: 5 });
    expect(Object.keys(parsed ?? {})).toEqual(["startedAtMs", "accumulatedMs"]);
  });

  it("rejects NaN, Infinity, negatives, strings, missing fields and non-objects, never throwing", () => {
    const bad: unknown[] = [
      { startedAtMs: Number.NaN, accumulatedMs: 0 },
      { startedAtMs: null, accumulatedMs: Number.NaN },
      { startedAtMs: Number.POSITIVE_INFINITY, accumulatedMs: 0 },
      { startedAtMs: null, accumulatedMs: Number.POSITIVE_INFINITY },
      { startedAtMs: -1, accumulatedMs: 0 },
      { startedAtMs: null, accumulatedMs: -5 },
      { startedAtMs: "1000", accumulatedMs: 0 },
      { startedAtMs: null, accumulatedMs: "0" },
      { accumulatedMs: 0 },
      { startedAtMs: null },
      {},
      null,
      undefined,
      42,
      "state",
      [],
      [null, 0],
    ];
    for (const raw of bad) {
      expect(() => parseTimerState(raw)).not.toThrow();
      expect(parseTimerState(raw), JSON.stringify(raw)).toBeNull();
    }
    const hostile = new Proxy({}, { get: () => { throw new Error("nope"); } });
    expect(parseTimerState(hostile)).toBeNull();
  });
});

describe("showTimerRecordKey", () => {
  it("names one sessionStorage record per deck", () => {
    expect(showTimerRecordKey("deck-9")).toBe("docu-edit:show-timer:deck-9");
  });
});
