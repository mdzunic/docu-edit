// SPEC-100 — the presenter timer: pure and clock-injected. Start, pause,
// resume and reset accumulate against the `nowMs` each call is given (no test
// touches the real clock), and formatElapsed switches from M:SS to H:MM:SS at
// one hour.
import { describe, expect, it } from "vitest";
import {
  elapsedMs,
  formatElapsed,
  IDLE_TIMER,
  isTimerRunning,
  pauseTimer,
  resetTimer,
  startTimer,
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
