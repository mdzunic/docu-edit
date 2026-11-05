// SPEC-146 — the show timer composable, mounted on a bare component with an
// injected clock and an in-memory storage: begin() (auto-start on and off,
// once per mount, a reload's record restored exactly), the record written on
// every change and removed on unmount, blocked or absent storage, adopt, the
// live-region announcements ("Time is up" exactly once per crossing, never
// on begin, adopt or a restore), mode and length as per-device preferences
// that leave the count alone, the once-a-second tick, and the timer keys.
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { defineComponent, ref, watch } from "vue";
import { mount, type VueWrapper } from "@vue/test-utils";
import { loadPreferences, resetPreferences } from "../../lib/preferences";
import { showTimerRecordKey, type TimerState } from "../../lib/presentation/presenterTimer";
import { showTimerKeys, TIMER_TICK_MS, timerKeyAction, useShowTimer, type ShowTimer, type ShowTimerDeps } from "./useShowTimer";

const DECK = "deck-7";
const KEY = showTimerRecordKey(DECK);
const T0 = 1_000_000;
const wrappers: VueWrapper[] = [];

/** A Storage stand-in that records every call. */
function memoryStorage(seed: Record<string, string> = {}) {
  const data = new Map(Object.entries(seed));
  return {
    data,
    getItem: vi.fn((key: string) => data.get(key) ?? null),
    setItem: vi.fn((key: string, value: string) => void data.set(key, value)),
    removeItem: vi.fn((key: string) => void data.delete(key)),
  };
}

/** What a browser with storage blocked does on every call. */
function boom(): never {
  throw new Error("SecurityError");
}

function record(state: TimerState): Record<string, string> {
  return { [KEY]: JSON.stringify(state) };
}

function harness(deps: Partial<ShowTimerDeps> = {}) {
  let clock = T0;
  const storage = "storage" in deps ? deps.storage : memoryStorage();
  const onChange = vi.fn();
  let timer!: ShowTimer;
  const wrapper = mount(
    defineComponent({
      setup() {
        timer = useShowTimer({ deckId: ref(DECK), nowMs: () => clock, onChange, ...deps, storage });
        return () => null;
      },
    }),
  );
  wrappers.push(wrapper);
  return {
    timer,
    wrapper,
    onChange,
    storage: storage as ReturnType<typeof memoryStorage>,
    /** Moves the injected clock to T0 + ms. */
    at(ms: number) {
      clock = T0 + ms;
    },
  };
}

beforeEach(() => {
  localStorage.clear();
  resetPreferences();
});

afterEach(() => {
  for (const wrapper of wrappers.splice(0)) wrapper.unmount();
  vi.useRealTimers();
});

describe("useShowTimer — begin()", () => {
  it("with auto-start on, starts running from 0:00 at nowMs", () => {
    const { timer, storage, onChange } = harness();
    expect(timer.phase.value).toBe("stopped");
    timer.begin();
    expect(timer.phase.value).toBe("running");
    expect(timer.timer.value).toEqual({ startedAtMs: T0, accumulatedMs: 0 });
    expect(timer.text.value).toBe("0:00");
    expect(JSON.parse(storage.data.get(KEY) ?? "null")).toEqual({ startedAtMs: T0, accumulatedMs: 0 });
    expect(onChange).toHaveBeenCalledWith({ startedAtMs: T0, accumulatedMs: 0 });
  });

  it("with auto-start off, stays at the stopped phase", () => {
    const { timer } = harness({ autoStart: () => false });
    timer.begin();
    expect(timer.phase.value).toBe("stopped");
    expect(timer.text.value).toBe("0:00");
  });

  it("reads presentationTimerAutoStart by default", async () => {
    const { savePreference } = await import("../../lib/preferences");
    savePreference("presentationTimerAutoStart", false);
    const { timer } = harness();
    timer.begin();
    expect(timer.phase.value).toBe("stopped");
  });

  it("runs once per mount: a second call (the view's second load pass) changes nothing", () => {
    const { timer, at, onChange } = harness();
    timer.begin();
    at(30_000);
    timer.toggle(); // paused at 0:30
    timer.begin();
    expect(timer.phase.value).toBe("paused");
    expect(timer.text.value).toBe("0:30");
    expect(onChange).toHaveBeenCalledTimes(2);
  });

  it("restores a running record exactly, counting the time the reload took", () => {
    const storage = memoryStorage(record({ startedAtMs: T0 - 30_000, accumulatedMs: 15_000 }));
    const { timer } = harness({ storage, autoStart: () => true });
    timer.begin();
    expect(timer.phase.value).toBe("running");
    expect(timer.text.value).toBe("0:45");
  });

  it("restores a paused or a stopped record as it was; auto-start never overrides it", () => {
    const paused = harness({ storage: memoryStorage(record({ startedAtMs: null, accumulatedMs: 75_000 })), autoStart: () => true });
    paused.timer.begin();
    expect(paused.timer.phase.value).toBe("paused");
    expect(paused.timer.text.value).toBe("1:15");
    const stopped = harness({ storage: memoryStorage(record({ startedAtMs: null, accumulatedMs: 0 })), autoStart: () => true });
    stopped.timer.begin();
    expect(stopped.timer.phase.value).toBe("stopped");
  });

  it("falls back to a fresh start for a corrupt record, or another deck's", () => {
    for (const raw of ["{not json", JSON.stringify({ startedAtMs: -5, accumulatedMs: 0 }), JSON.stringify("running"), "null"]) {
      const { timer } = harness({ storage: memoryStorage({ [KEY]: raw }), autoStart: () => false });
      timer.begin();
      expect(timer.phase.value, raw).toBe("stopped");
    }
    const other = harness({ storage: memoryStorage({ [showTimerRecordKey("other")]: JSON.stringify({ startedAtMs: null, accumulatedMs: 9_000 }) }) });
    other.timer.begin();
    expect(other.timer.timer.value).toEqual({ startedAtMs: T0, accumulatedMs: 0 });
  });
});

describe("useShowTimer — toggle, stop and the record", () => {
  it("toggles stopped → running → paused → running and stops from either, the paused stretch not counted", () => {
    const { timer, at } = harness({ autoStart: () => false });
    timer.begin();
    timer.toggle();
    expect(timer.phase.value).toBe("running");
    at(75_000);
    timer.toggle();
    expect(timer.phase.value).toBe("paused");
    expect(timer.text.value).toBe("1:15");
    at(600_000);
    timer.toggle();
    expect(timer.text.value).toBe("1:15");
    at(605_000);
    timer.stop();
    expect(timer.phase.value).toBe("stopped");
    expect(timer.text.value).toBe("0:00");
    timer.toggle();
    at(610_000);
    timer.toggle();
    timer.stop(); // from paused
    expect(timer.phase.value).toBe("stopped");
  });

  it("stop does nothing when already stopped", () => {
    const { timer, onChange, storage } = harness({ autoStart: () => false });
    timer.begin();
    const writes = storage.setItem.mock.calls.length;
    timer.stop();
    expect(onChange).toHaveBeenCalledTimes(1); // begin only
    expect(storage.setItem.mock.calls.length).toBe(writes);
    expect(timer.status.value).toBe("");
  });

  it("writes the record on every change and removes it on unmount", () => {
    const { timer, at, storage, wrapper, onChange } = harness();
    timer.begin();
    at(10_000);
    timer.toggle();
    expect(JSON.parse(storage.data.get(KEY) ?? "null")).toEqual({ startedAtMs: null, accumulatedMs: 10_000 });
    timer.toggle();
    timer.stop();
    expect(storage.setItem).toHaveBeenCalledTimes(4);
    expect(onChange).toHaveBeenCalledTimes(4);
    expect(JSON.parse(storage.data.get(KEY) ?? "null")).toEqual({ startedAtMs: null, accumulatedMs: 0 });
    wrapper.unmount();
    expect(storage.removeItem).toHaveBeenCalledWith(KEY);
    expect(storage.data.has(KEY)).toBe(false);
  });

  it("a window that never owned the show (a follower) never reads, writes or removes the record", () => {
    const storage = memoryStorage(record({ startedAtMs: null, accumulatedMs: 5_000 }));
    const { wrapper } = harness({ storage });
    wrapper.unmount();
    expect(storage.getItem).not.toHaveBeenCalled();
    expect(storage.setItem).not.toHaveBeenCalled();
    expect(storage.removeItem).not.toHaveBeenCalled();
    expect(storage.data.has(KEY)).toBe(true);
  });

  it("tolerates a storage that throws on every call, running without persistence", () => {
    const storage = { getItem: vi.fn(boom), setItem: vi.fn(boom), removeItem: vi.fn(boom) };
    const { timer, at, wrapper } = harness({ storage });
    expect(() => timer.begin()).not.toThrow();
    expect(timer.phase.value).toBe("running");
    at(5_000);
    expect(() => timer.toggle()).not.toThrow();
    expect(timer.text.value).toBe("0:05");
    expect(() => wrapper.unmount()).not.toThrow();
    expect(storage.setItem).toHaveBeenCalled();
  });

  it("with null storage persists nothing and still runs", () => {
    const { timer } = harness({ storage: null });
    timer.begin();
    timer.toggle();
    expect(timer.phase.value).toBe("stopped"); // a 0 ms pause banks nothing
    expect(sessionStorage.getItem(KEY)).toBeNull();
  });

  it("defaults to the tab's sessionStorage", () => {
    sessionStorage.clear();
    let timer!: ShowTimer;
    const wrapper = mount(
      defineComponent({
        setup() {
          timer = useShowTimer({ deckId: ref(DECK), nowMs: () => T0 });
          return () => null;
        },
      }),
    );
    timer.begin();
    expect(JSON.parse(sessionStorage.getItem(KEY) ?? "null")).toEqual({ startedAtMs: T0, accumulatedMs: 0 });
    wrapper.unmount();
    expect(sessionStorage.getItem(KEY)).toBeNull();
  });
});

describe("useShowTimer — adopt", () => {
  it("takes over another window's timer as it was, running or paused", () => {
    const running = harness({ autoStart: () => false });
    running.timer.adopt({ startedAtMs: T0 - 75_000, accumulatedMs: 0 });
    expect(running.timer.phase.value).toBe("running");
    expect(running.timer.text.value).toBe("1:15");
    const paused = harness();
    paused.timer.adopt({ startedAtMs: null, accumulatedMs: 75_000 });
    expect(paused.timer.phase.value).toBe("paused");
    paused.at(600_000);
    paused.timer.begin(); // a promoted window's later load pass never restarts it
    expect(paused.timer.text.value).toBe("1:15");
    expect(JSON.parse(paused.storage.data.get(KEY) ?? "null")).toEqual({ startedAtMs: null, accumulatedMs: 75_000 });
    expect(paused.onChange).toHaveBeenCalledTimes(1);
    expect(paused.timer.status.value).toBe("");
  });

  it("shows the adopted timer in the preferences' mode and length", async () => {
    const { savePreference } = await import("../../lib/preferences");
    savePreference("presentationTimerMode", "countdown");
    savePreference("presentationTimerCountdownMinutes", 2);
    const { timer } = harness();
    timer.adopt({ startedAtMs: null, accumulatedMs: 75_000 });
    expect(timer.text.value).toBe("0:45");
    expect(timer.mode.value).toBe("countdown");
  });
});

describe("useShowTimer — the live region", () => {
  it("announces each transition, and nothing on begin", () => {
    const { timer, at } = harness({ autoStart: () => false });
    timer.begin();
    expect(timer.status.value).toBe("");
    timer.toggle();
    expect(timer.status.value).toBe("Timer started");
    at(1_000);
    timer.toggle();
    expect(timer.status.value).toBe("Timer paused");
    timer.toggle();
    expect(timer.status.value).toBe("Timer resumed");
    timer.stop();
    expect(timer.status.value).toBe("Timer stopped");
  });

  it("announces nothing on a restore or an ordinary tick", () => {
    vi.useFakeTimers({ toFake: ["setInterval", "clearInterval"] });
    const { timer, at } = harness({ storage: memoryStorage(record({ startedAtMs: T0, accumulatedMs: 0 })) });
    timer.begin();
    at(5_000);
    vi.advanceTimersByTime(TIMER_TICK_MS * 3);
    expect(timer.text.value).toBe("0:05");
    expect(timer.status.value).toBe("");
  });

  it("says 'Time is up' exactly once when a running countdown crosses zero on a tick", () => {
    vi.useFakeTimers({ toFake: ["setInterval", "clearInterval"] });
    const { timer, at } = harness({ mode: () => "countdown", lengthMinutes: () => 2 });
    timer.begin();
    expect(timer.text.value).toBe("2:00");
    const said: string[] = [];
    const unwatch = watch(timer.status, (value) => said.push(value), { flush: "sync" });
    for (let second = 100; second <= 200; second += 1) {
      at(second * 1000);
      vi.advanceTimersByTime(TIMER_TICK_MS);
      if (second === 119) expect(timer.overtime.value).toBe(false);
      if (second === 120) expect(timer.text.value).toBe("0:00");
    }
    unwatch();
    expect(said).toEqual(["Time is up"]);
    expect(timer.text.value).toBe("\u22121:20");
    expect(timer.overtime.value).toBe(true);
  });

  it("says 'Time is up' once when setLengthMinutes drops the length below the elapsed time", () => {
    const { timer, at } = harness();
    timer.begin();
    timer.setMode("countdown");
    at(75_000);
    timer.setLengthMinutes(1);
    expect(timer.text.value).toBe("−0:15");
    expect(timer.status.value).toBe("Time is up");
    timer.setLengthMinutes(1); // no second crossing
    expect(timer.status.value).toBe("Time is up");
  });

  it("does not announce a restore that lands in overtime, nor a crossing while paused", () => {
    const restored = harness({
      storage: memoryStorage(record({ startedAtMs: T0 - 162_000, accumulatedMs: 0 })),
      mode: () => "countdown",
      lengthMinutes: () => 2,
    });
    restored.timer.begin();
    expect(restored.timer.overtime.value).toBe(true);
    expect(restored.timer.status.value).toBe("");

    const adopted = harness({ mode: () => "countdown", lengthMinutes: () => 2 });
    adopted.timer.adopt({ startedAtMs: T0 - 162_000, accumulatedMs: 0 });
    expect(adopted.timer.overtime.value).toBe(true);
    expect(adopted.timer.status.value).toBe("");

    const paused = harness();
    paused.timer.begin();
    paused.timer.setMode("countdown");
    paused.at(75_000);
    paused.timer.toggle(); // paused at 1:15
    paused.timer.setLengthMinutes(1);
    expect(paused.timer.overtime.value).toBe(true);
    expect(paused.timer.status.value).toBe("Timer paused");
  });
});

describe("useShowTimer — mode and talk length", () => {
  it("setMode and setLengthMinutes write the per-device preferences and leave the timer untouched", () => {
    const { timer, at } = harness();
    timer.begin();
    at(75_000);
    const before = timer.timer.value;
    timer.setMode("countdown");
    expect(loadPreferences().presentationTimerMode).toBe("countdown");
    expect(timer.text.value).toBe("18:45"); // re-stamped at once
    timer.setLengthMinutes(30.4);
    expect(loadPreferences().presentationTimerCountdownMinutes).toBe(30);
    expect(timer.text.value).toBe("28:45");
    timer.setLengthMinutes(999);
    expect(loadPreferences().presentationTimerCountdownMinutes).toBe(240);
    timer.setLengthMinutes(Number.NaN); // ignored
    expect(loadPreferences().presentationTimerCountdownMinutes).toBe(240);
    timer.setMode("stopwatch");
    expect(timer.text.value).toBe("1:15");
    expect(timer.timer.value).toBe(before);
    expect(timer.phase.value).toBe("running");
  });

  it("a stopped timer reads the start: 0:00, or the full talk length", () => {
    const { timer } = harness({ autoStart: () => false });
    timer.begin();
    expect(timer.text.value).toBe("0:00");
    timer.setMode("countdown");
    expect(timer.text.value).toBe("20:00");
    expect(timer.lengthMinutes.value).toBe(20);
  });
});

describe("useShowTimer — the tick", () => {
  // Moved here from usePresenterView.spec.ts (SPEC-100) with its three
  // assertions intact.
  it("ticks the timer display once a second from the injected clock", () => {
    vi.useFakeTimers({ toFake: ["setInterval", "clearInterval"] });
    let clock = 1_000_000;
    let timer!: ShowTimer;
    wrappers.push(
      mount(
        defineComponent({
          setup() {
            timer = useShowTimer({ deckId: ref(DECK), nowMs: () => clock, storage: null, autoStart: () => true });
            return () => null;
          },
        }),
      ),
    );
    timer.begin();
    clock += 61_000;
    expect(timer.text.value).toBe("0:00"); // not until the tick
    vi.advanceTimersByTime(TIMER_TICK_MS);
    expect(TIMER_TICK_MS).toBe(1000);
    expect(timer.text.value).toBe("1:01");
    clock += 3_600_000;
    vi.advanceTimersByTime(TIMER_TICK_MS);
    expect(timer.text.value).toBe("1:01:01");
  });

  it("stops ticking on unmount", () => {
    vi.useFakeTimers({ toFake: ["setInterval", "clearInterval"] });
    const { wrapper } = harness();
    expect(vi.getTimerCount()).toBe(1);
    wrapper.unmount();
    expect(vi.getTimerCount()).toBe(0);
  });
});

describe("the timer keys", () => {
  it("t toggles; s and r stop; nothing else is a timer key", () => {
    expect(timerKeyAction("t")).toBe("timerToggle");
    expect(timerKeyAction("T")).toBe("timerToggle");
    expect(timerKeyAction("s")).toBe("timerStop");
    expect(timerKeyAction("r")).toBe("timerStop");
    for (const key of ["b", "w", "n", "p", "ArrowRight", " ", "Enter", "f", "1"]) expect(timerKeyAction(key)).toBeNull();
  });

  it("showTimerKeys performs and reveals, falls through on the end screen, and leaves Enter/Space to a focused timer button", () => {
    const atEnd = ref(false);
    const perform = vi.fn();
    const onKey = vi.fn();
    const extraKey = showTimerKeys({ atEnd, perform, onKey });
    const press = (key: string, target?: EventTarget) => {
      const event = new KeyboardEvent("keydown", { key, cancelable: true });
      if (target) Object.defineProperty(event, "target", { value: target });
      return { consumed: extraKey(event), prevented: event.defaultPrevented };
    };
    expect(press("t")).toEqual({ consumed: true, prevented: true });
    expect(press("s")).toEqual({ consumed: true, prevented: true });
    expect(perform.mock.calls).toEqual([["timerToggle"], ["timerStop"]]);
    expect(onKey).toHaveBeenCalledTimes(2);
    expect(press("ArrowRight")).toEqual({ consumed: false, prevented: false });
    const button = document.createElement("button");
    button.setAttribute("data-show-timer-control", "");
    expect(press("Enter", button)).toEqual({ consumed: true, prevented: false }); // the button's own activation
    expect(press(" ", button)).toEqual({ consumed: true, prevented: false });
    expect(press("Enter", document.createElement("button"))).toEqual({ consumed: false, prevented: false });
    atEnd.value = true; // the end screen: everything falls through to its any-key exit
    expect(press("t")).toEqual({ consumed: false, prevented: false });
    expect(press("Enter", button)).toEqual({ consumed: false, prevented: false });
    expect(perform).toHaveBeenCalledTimes(2);
  });
});

describe("the show timer is inert (SPEC-146)", () => {
  it("no timer module dirties, edits, locks or writes the deck", () => {
    const here = dirname(fileURLToPath(import.meta.url));
    const files = ["useShowTimer.ts", "useAutoAdvance.ts", "../../components/presentation/ShowTimer.vue", "../../lib/presentation/presenterTimer.ts"];
    for (const file of files) {
      const source = readFileSync(join(here, file), "utf8").replace(/\/\/.*$/gm, "").replace(/\/\*[\s\S]*?\*\//g, "");
      for (const forbidden of ["acquireEditLock", "markSlideEdited", "notesDirty", "structureDirty", "dirty.value", "storePresentation", "writePptx", "writeOdp", "localStorage"]) {
        expect(source.includes(forbidden), `${file} must not use ${forbidden}`).toBe(false);
      }
    }
  });
});
