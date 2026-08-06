// SPEC-028 — the eager macrotask scheduler the pagination measure loop
// chains on: slots run asynchronously in FIFO order with the caller's
// budget, and a cancelled slot never fires.
import { describe, expect, it } from "vitest";
import { scheduleEagerWork } from "./idleScheduler";

function nextMacrotask(): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

describe("scheduleEagerWork", () => {
  it("runs the callback asynchronously with the given budget", async () => {
    const calls: number[] = [];
    scheduleEagerWork((budget) => calls.push(budget), 24);
    expect(calls).toEqual([]);
    await nextMacrotask();
    expect(calls).toEqual([24]);
  });

  it("runs multiple slots in scheduling order", async () => {
    const calls: string[] = [];
    scheduleEagerWork(() => calls.push("a"), 1);
    scheduleEagerWork(() => calls.push("b"), 1);
    await nextMacrotask();
    expect(calls).toEqual(["a", "b"]);
  });

  it("never fires a cancelled slot", async () => {
    const calls: string[] = [];
    const handle = scheduleEagerWork(() => calls.push("cancelled"), 1);
    scheduleEagerWork(() => calls.push("kept"), 1);
    handle.cancel();
    await nextMacrotask();
    expect(calls).toEqual(["kept"]);
  });

  it("supports rescheduling from within a slot", async () => {
    const calls: number[] = [];
    scheduleEagerWork(() => {
      calls.push(1);
      scheduleEagerWork(() => calls.push(2), 1);
    }, 1);
    await nextMacrotask();
    await nextMacrotask();
    expect(calls).toEqual([1, 2]);
  });
});
