// SPEC-028 — the eager macrotask scheduler the pagination measure loop
// chains on: slots run asynchronously in FIFO order with the caller's
// budget, and a cancelled slot never fires.
import { describe, expect, it, vi } from "vitest";
import { scheduleEagerWork } from "./idleScheduler";

// These used to await a single `setTimeout(0)` and assert the slot had run by
// then. That races the scheduler's transport rather than testing it: delivery
// goes through a MessageChannel, and jsdom does not guarantee a port message
// beats a zero-delay timer. On an idle machine the port wins every time; in the
// factory's container, where the suite runs ~7x slower, the timer wins often
// enough to redden `main` (see SPEC-028's scheduler). Waiting for the
// scheduler's own effect keeps every assertion below identical while removing
// the dependency on which of the two queues drains first.
const untilSettled = { timeout: 5000, interval: 10 } as const;

describe("scheduleEagerWork", () => {
  it("runs the callback asynchronously with the given budget", async () => {
    const calls: number[] = [];
    scheduleEagerWork((budget) => calls.push(budget), 24);
    expect(calls).toEqual([]);
    await vi.waitFor(() => expect(calls).toEqual([24]), untilSettled);
  });

  it("runs multiple slots in scheduling order", async () => {
    const calls: string[] = [];
    scheduleEagerWork(() => calls.push("a"), 1);
    scheduleEagerWork(() => calls.push("b"), 1);
    await vi.waitFor(() => expect(calls).toEqual(["a", "b"]), untilSettled);
  });

  it("never fires a cancelled slot", async () => {
    const calls: string[] = [];
    const handle = scheduleEagerWork(() => calls.push("cancelled"), 1);
    scheduleEagerWork(() => calls.push("kept"), 1);
    handle.cancel();
    // Slots run FIFO and the cancelled one was queued first, so observing
    // "kept" is proof the cancelled slot was skipped, not merely still pending.
    await vi.waitFor(() => expect(calls).toEqual(["kept"]), untilSettled);
  });

  it("supports rescheduling from within a slot", async () => {
    const calls: number[] = [];
    scheduleEagerWork(() => {
      calls.push(1);
      scheduleEagerWork(() => calls.push(2), 1);
    }, 1);
    await vi.waitFor(() => expect(calls).toEqual([1, 2]), untilSettled);
  });
});
