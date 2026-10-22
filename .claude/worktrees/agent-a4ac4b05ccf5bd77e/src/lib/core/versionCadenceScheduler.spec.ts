import { afterEach, beforeEach, describe, it, expect, vi } from "vitest";
import { createVersionCadenceTimer } from "./versionCadenceScheduler";

// SPEC-040 AC-1: a deadline armed once on the first edit after idle, firing
// exactly `intervalMs` later regardless of continued activity — no debounce
// component (see the module's header comment for why).
describe("versionCadenceScheduler", () => {
  beforeEach(() => {
    vi.useFakeTimers();
  });

  afterEach(() => {
    vi.useRealTimers();
  });

  it("fires once, intervalMs after the first change", () => {
    const onDue = vi.fn();
    const timer = createVersionCadenceTimer({ onDue, intervalMs: 1000 });
    timer.noteChange();
    vi.advanceTimersByTime(999);
    expect(onDue).not.toHaveBeenCalled();
    vi.advanceTimersByTime(1);
    expect(onDue).toHaveBeenCalledTimes(1);
    timer.dispose();
  });

  it("does not reset the deadline on further activity — unlike a debounce", () => {
    const onDue = vi.fn();
    const timer = createVersionCadenceTimer({ onDue, intervalMs: 1000 });
    timer.noteChange();
    vi.advanceTimersByTime(600);
    timer.noteChange(); // continued typing must not push the deadline out
    vi.advanceTimersByTime(400);
    expect(onDue).toHaveBeenCalledTimes(1);
    timer.dispose();
  });

  it("re-arms after firing, for the next burst of activity", () => {
    const onDue = vi.fn();
    const timer = createVersionCadenceTimer({ onDue, intervalMs: 1000 });
    timer.noteChange();
    vi.advanceTimersByTime(1000);
    expect(onDue).toHaveBeenCalledTimes(1);
    timer.noteChange();
    vi.advanceTimersByTime(1000);
    expect(onDue).toHaveBeenCalledTimes(2);
    timer.dispose();
  });

  it("reset clears a pending deadline without disabling future ones", () => {
    const onDue = vi.fn();
    const timer = createVersionCadenceTimer({ onDue, intervalMs: 1000 });
    timer.noteChange();
    vi.advanceTimersByTime(500);
    timer.reset();
    vi.advanceTimersByTime(1000);
    expect(onDue).not.toHaveBeenCalled();
    timer.noteChange();
    vi.advanceTimersByTime(1000);
    expect(onDue).toHaveBeenCalledTimes(1);
    timer.dispose();
  });

  it("dispose cancels everything and ignores later calls", () => {
    const onDue = vi.fn();
    const timer = createVersionCadenceTimer({ onDue, intervalMs: 1000 });
    timer.noteChange();
    timer.dispose();
    timer.noteChange();
    vi.advanceTimersByTime(1000);
    expect(onDue).not.toHaveBeenCalled();
  });
});
