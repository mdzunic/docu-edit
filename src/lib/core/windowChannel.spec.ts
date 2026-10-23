// SPEC-100 — the product-neutral window channel: transport selection, the
// localStorage `storage`-event fallback (posting and receiving through an
// injected storage), payload passthrough, and close() idempotence. A fake
// "browser" connects two channel ends the way the real globals would.
import { describe, expect, it, vi } from "vitest";
import {
  channelStorageKey,
  openWindowChannel,
  windowChannelTransport,
  type BroadcastPort,
  type ChannelEnvironment,
  type ChannelStorage,
} from "./windowChannel";

/** A Map-backed Storage whose writes notify every OTHER window's listeners,
 * only when the value changes — the `storage` event's real semantics. */
function sharedStorage() {
  const data = new Map<string, string>();
  const windows: Array<Set<(key: string | null, value: string | null) => void>> = [];
  function windowEnv(): ChannelEnvironment {
    const listeners = new Set<(key: string | null, value: string | null) => void>();
    windows.push(listeners);
    const storage: ChannelStorage = {
      getItem: (key) => data.get(key) ?? null,
      setItem: (key, value) => {
        const old = data.get(key) ?? null;
        data.set(key, value);
        if (old === value) return;
        for (const other of windows) if (other !== listeners) for (const listener of other) listener(key, value);
      },
      removeItem: (key) => {
        const had = data.has(key);
        data.delete(key);
        if (!had) return;
        for (const other of windows) if (other !== listeners) for (const listener of other) listener(key, null);
      },
    };
    return {
      storage,
      now: () => 1234,
      onStorage: (listener) => {
        listeners.add(listener);
        return () => listeners.delete(listener);
      },
    };
  }
  return { data, windowEnv };
}

/** A fake BroadcastChannel hub: a post reaches every other port of the name. */
function broadcastHub() {
  const ports: Array<{ name: string; port: BroadcastPort; listeners: Set<(event: MessageEvent) => void> }> = [];
  const open = (name: string): BroadcastPort => {
    const listeners = new Set<(event: MessageEvent) => void>();
    const port: BroadcastPort = {
      addEventListener: (_type, listener) => listeners.add(listener),
      removeEventListener: (_type, listener) => listeners.delete(listener),
      postMessage: (data) => {
        for (const other of ports) {
          if (other.port !== port && other.name === name) {
            for (const listener of other.listeners) listener({ data } as MessageEvent);
          }
        }
      },
      close: vi.fn(),
    };
    ports.push({ name, port, listeners });
    return port;
  };
  return { open, ports };
}

describe("windowChannelTransport", () => {
  it("prefers BroadcastChannel where it exists", () => {
    const { open } = broadcastHub();
    expect(windowChannelTransport({ broadcast: open, ...sharedStorage().windowEnv() })).toBe("broadcast");
  });

  it("falls back to storage events when only storage is usable", () => {
    expect(windowChannelTransport(sharedStorage().windowEnv())).toBe("storage");
  });

  it("is none when storage throws (blocked) and there is no BroadcastChannel", () => {
    const env = sharedStorage().windowEnv();
    env.storage = {
      getItem: () => null,
      setItem: () => {
        throw new Error("SecurityError");
      },
      removeItem: () => {},
    };
    expect(windowChannelTransport(env)).toBe("none");
    expect(windowChannelTransport({})).toBe("none");
  });

  it("reads the real browser by default (jsdom has localStorage)", () => {
    expect(["broadcast", "storage"]).toContain(windowChannelTransport());
  });
});

describe("openWindowChannel over BroadcastChannel", () => {
  it("delivers each payload verbatim to the other window, never to itself", () => {
    const { open } = broadcastHub();
    const a = vi.fn();
    const b = vi.fn();
    const left = openWindowChannel("deck", a, { broadcast: open });
    openWindowChannel("deck", b, { broadcast: open });
    const payload = { kind: "hello", nested: [1, { two: 2 }] };
    left.post(payload);
    expect(b).toHaveBeenCalledWith(payload);
    expect(a).not.toHaveBeenCalled();
  });

  it("close() stops delivery, closes the port once, and later posts are no-ops", () => {
    const { open, ports } = broadcastHub();
    const b = vi.fn();
    const left = openWindowChannel("deck", vi.fn(), { broadcast: open });
    const right = openWindowChannel("deck", b, { broadcast: open });
    right.close();
    right.close();
    expect(ports[1].port.close).toHaveBeenCalledTimes(1);
    left.post("after close");
    expect(b).not.toHaveBeenCalled();
    expect(() => right.post("x")).not.toThrow();
  });
});

describe("openWindowChannel over the storage fallback", () => {
  it("writes { seq, at, payload } under docu-edit:channel:<name> and the other window receives the payload", () => {
    const shared = sharedStorage();
    const received = vi.fn();
    const presenter = openWindowChannel("docu-edit:present:d1", vi.fn(), shared.windowEnv());
    openWindowChannel("docu-edit:present:d1", received, shared.windowEnv());
    presenter.post({ kind: "state", n: 1 });
    const stored = JSON.parse(shared.data.get(channelStorageKey("docu-edit:present:d1")) ?? "{}");
    expect(stored).toEqual({ seq: 1, at: 1234, payload: { kind: "state", n: 1 } });
    expect(received).toHaveBeenCalledWith({ kind: "state", n: 1 });
  });

  it("an identical payload posted twice still fires twice (seq always moves on)", () => {
    const shared = sharedStorage();
    const received = vi.fn();
    const sender = openWindowChannel("c", vi.fn(), shared.windowEnv());
    openWindowChannel("c", received, shared.windowEnv());
    sender.post("same");
    sender.post("same");
    expect(received).toHaveBeenCalledTimes(2);
  });

  it("seq continues past what another window stored, so the value always differs", () => {
    const shared = sharedStorage();
    const received = vi.fn();
    const a = openWindowChannel("c", vi.fn(), shared.windowEnv());
    const b = openWindowChannel("c", vi.fn(), shared.windowEnv());
    openWindowChannel("c", received, shared.windowEnv());
    a.post("one");
    a.post("two");
    b.post("three");
    expect(JSON.parse(shared.data.get(channelStorageKey("c")) ?? "{}").seq).toBe(3);
    expect(received.mock.calls.map((call) => call[0])).toEqual(["one", "two", "three"]);
  });

  it("ignores other keys, removals and unparseable values", () => {
    const shared = sharedStorage();
    const received = vi.fn();
    openWindowChannel("c", received, shared.windowEnv());
    const other = shared.windowEnv();
    other.storage?.setItem("docu-edit:preferences", "{}");
    other.storage?.setItem(channelStorageKey("c"), "not json");
    other.storage?.setItem(channelStorageKey("c"), JSON.stringify({ seq: 9 }));
    other.storage?.removeItem(channelStorageKey("c"));
    expect(received).not.toHaveBeenCalled();
  });

  it("close() unsubscribes and clears its key; it is idempotent and later posts write nothing", () => {
    const shared = sharedStorage();
    const received = vi.fn();
    const sender = openWindowChannel("c", vi.fn(), shared.windowEnv());
    const receiver = openWindowChannel("c", received, shared.windowEnv());
    sender.post("before");
    receiver.close();
    receiver.close();
    sender.post("after");
    expect(received.mock.calls.map((call) => call[0])).toEqual(["before"]);
    sender.close();
    expect(shared.data.has(channelStorageKey("c"))).toBe(false);
    sender.post("closed");
    expect(shared.data.has(channelStorageKey("c"))).toBe(false);
  });

  it("a storage that fills up mid-session loses the message, never throws", () => {
    const env = sharedStorage().windowEnv();
    const channel = openWindowChannel("c", vi.fn(), env);
    env.storage!.setItem = () => {
      throw new Error("QuotaExceededError");
    };
    expect(() => channel.post("x")).not.toThrow();
  });
});

describe("openWindowChannel with no transport", () => {
  it("returns an inert channel", () => {
    const channel = openWindowChannel("c", vi.fn(), {});
    expect(() => {
      channel.post({});
      channel.close();
      channel.close();
    }).not.toThrow();
  });
});
