import { beforeEach, afterEach, describe, it, expect, vi } from "vitest";

// Fresh module per test: the toast list is module-level state on purpose
// (it must survive route changes), so tests isolate it via resetModules.
async function loadToasts() {
  const module = await import("./useToasts");
  return { ...module.useToasts(), TOAST_DURATION_MS: module.TOAST_DURATION_MS };
}

beforeEach(() => {
  vi.resetModules();
  vi.useFakeTimers();
});

afterEach(() => {
  vi.useRealTimers();
});

describe("useToasts", () => {
  it("adds toasts with unique ids", async () => {
    const { toasts, addToast } = await loadToasts();
    addToast("first");
    addToast("second");
    expect(toasts.value.map((toast) => toast.message)).toEqual([
      "first",
      "second",
    ]);
    expect(toasts.value[0]!.id).not.toBe(toasts.value[1]!.id);
  });

  it("auto-dismisses a toast after its duration", async () => {
    const { toasts, addToast, TOAST_DURATION_MS } = await loadToasts();
    addToast("temporary");
    expect(toasts.value).toHaveLength(1);
    vi.advanceTimersByTime(TOAST_DURATION_MS);
    expect(toasts.value).toHaveLength(0);
  });

  it("dismisses a specific toast on demand", async () => {
    const { toasts, addToast, dismissToast } = await loadToasts();
    addToast("keep");
    addToast("drop");
    const dropId = toasts.value[1]!.id;
    dismissToast(dropId);
    expect(toasts.value.map((toast) => toast.message)).toEqual(["keep"]);
  });

  it("shares one toast list across composable instances", async () => {
    const first = await loadToasts();
    const module = await import("./useToasts");
    const second = module.useToasts();
    first.addToast("shared");
    expect(second.toasts.value.map((toast) => toast.message)).toEqual([
      "shared",
    ]);
  });
});
