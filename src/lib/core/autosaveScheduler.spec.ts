import { afterEach, beforeEach, describe, it, expect, vi } from "vitest";
import {
  AUTOSAVE_DEBOUNCE_MS,
  AUTOSAVE_MAX_INTERVAL_MS,
  AUTOSAVE_RETRY_MS,
  createAutosaveScheduler,
} from "./autosaveScheduler";

// SPEC-010 AC: at most 2 s after typing stops, at least every 30 s during
// continuous typing. Driven entirely by fake timers.
describe("autosaveScheduler", () => {
  beforeEach(() => {
    vi.useFakeTimers();
  });

  afterEach(() => {
    vi.useRealTimers();
  });

  it("flushes 2 s after the last change (debounce)", () => {
    const flush = vi.fn();
    const scheduler = createAutosaveScheduler({ flush });
    scheduler.noteChange();
    vi.advanceTimersByTime(AUTOSAVE_DEBOUNCE_MS - 1);
    expect(flush).not.toHaveBeenCalled();
    vi.advanceTimersByTime(1);
    expect(flush).toHaveBeenCalledTimes(1);
    scheduler.dispose();
  });

  it("resets the debounce on every change", () => {
    const flush = vi.fn();
    const scheduler = createAutosaveScheduler({ flush });
    scheduler.noteChange();
    vi.advanceTimersByTime(1500);
    scheduler.noteChange();
    vi.advanceTimersByTime(1500);
    expect(flush).not.toHaveBeenCalled();
    vi.advanceTimersByTime(500);
    expect(flush).toHaveBeenCalledTimes(1);
    scheduler.dispose();
  });

  it("flushes at the 30 s deadline during continuous typing", () => {
    const flush = vi.fn();
    const scheduler = createAutosaveScheduler({ flush });
    // A change every second keeps the debounce from ever firing.
    for (let elapsed = 0; elapsed < AUTOSAVE_MAX_INTERVAL_MS; elapsed += 1000) {
      scheduler.noteChange();
      vi.advanceTimersByTime(1000);
    }
    expect(flush).toHaveBeenCalledTimes(1);
    scheduler.dispose();
  });

  it("starts a fresh 30 s deadline after each flush", () => {
    const flush = vi.fn();
    const scheduler = createAutosaveScheduler({ flush });
    for (let elapsed = 0; elapsed < 2 * AUTOSAVE_MAX_INTERVAL_MS; elapsed += 1000) {
      scheduler.noteChange();
      vi.advanceTimersByTime(1000);
    }
    expect(flush).toHaveBeenCalledTimes(2);
    scheduler.dispose();
  });

  it("retries after a failed save when no typing re-arms the debounce", () => {
    const flush = vi.fn();
    const scheduler = createAutosaveScheduler({ flush });
    scheduler.noteChange();
    vi.advanceTimersByTime(AUTOSAVE_DEBOUNCE_MS);
    expect(flush).toHaveBeenCalledTimes(1);
    scheduler.notifySaveResult(false);
    vi.advanceTimersByTime(AUTOSAVE_RETRY_MS);
    expect(flush).toHaveBeenCalledTimes(2);
    scheduler.dispose();
  });

  it("does not retry after a successful save", () => {
    const flush = vi.fn();
    const scheduler = createAutosaveScheduler({ flush });
    scheduler.noteChange();
    vi.advanceTimersByTime(AUTOSAVE_DEBOUNCE_MS);
    scheduler.notifySaveResult(true);
    vi.advanceTimersByTime(AUTOSAVE_MAX_INTERVAL_MS);
    expect(flush).toHaveBeenCalledTimes(1);
    scheduler.dispose();
  });

  it("skips the retry when typing has already re-armed the debounce", () => {
    const flush = vi.fn();
    const scheduler = createAutosaveScheduler({ flush });
    scheduler.noteChange();
    vi.advanceTimersByTime(AUTOSAVE_DEBOUNCE_MS);
    scheduler.noteChange();
    scheduler.notifySaveResult(false);
    vi.advanceTimersByTime(AUTOSAVE_DEBOUNCE_MS);
    // Only the debounce flush — no additional retry flush afterwards.
    expect(flush).toHaveBeenCalledTimes(2);
    vi.advanceTimersByTime(AUTOSAVE_RETRY_MS);
    expect(flush).toHaveBeenCalledTimes(2);
    scheduler.dispose();
  });

  it("flushNow fires immediately and cancels pending timers", () => {
    const flush = vi.fn();
    const scheduler = createAutosaveScheduler({ flush });
    scheduler.noteChange();
    scheduler.flushNow();
    expect(flush).toHaveBeenCalledTimes(1);
    vi.advanceTimersByTime(AUTOSAVE_MAX_INTERVAL_MS);
    expect(flush).toHaveBeenCalledTimes(1);
    scheduler.dispose();
  });

  it("dispose cancels everything and ignores later calls", () => {
    const flush = vi.fn();
    const scheduler = createAutosaveScheduler({ flush });
    scheduler.noteChange();
    scheduler.dispose();
    scheduler.noteChange();
    scheduler.flushNow();
    vi.advanceTimersByTime(AUTOSAVE_MAX_INTERVAL_MS);
    expect(flush).not.toHaveBeenCalled();
  });
});
