// A same-origin message channel between two windows of this app (SPEC-100) —
// transport only, payload-agnostic. SPEC-100's presenter session is its first
// caller, but nothing here knows about slides: a channel carries whatever plain
// JSON-shaped payload the caller posts, and every receiver must validate what
// arrives, because it came from another window.
//
// Three transports, picked once per channel:
//   · "broadcast" — `BroadcastChannel(name)`, where the browser has it;
//   · "storage"   — the `localStorage` key `docu-edit:channel:<name>`, written
//     as `{ seq, at, payload }`. The `storage` event fires only in OTHER
//     same-origin windows, which is exactly the delivery a channel needs, and
//     only when the stored value changes — so `seq` is always one past the
//     value already stored and the event always fires;
//   · "none"      — neither is usable (no BroadcastChannel and storage blocked):
//     `post` is a no-op, and callers fall back to a single window.
//
// This is not core/editLock.ts: that Web Lock provides exclusion, not
// messages. Framework-free; the environment is injectable so the storage
// fallback is unit-testable without a second window.

export type WindowChannelTransport = "broadcast" | "storage" | "none";

export interface WindowChannel {
  post(payload: unknown): void;
  close(): void;
}

/** The slice of `BroadcastChannel` a channel uses. */
export interface BroadcastPort {
  addEventListener(type: "message", listener: (event: MessageEvent) => void): void;
  removeEventListener(type: "message", listener: (event: MessageEvent) => void): void;
  postMessage(data: unknown): void;
  close(): void;
}

/** The slice of `Storage` the fallback uses. */
export interface ChannelStorage {
  getItem(key: string): string | null;
  setItem(key: string, value: string): void;
  removeItem(key: string): void;
}

/** What a channel needs from the browser. Every field is optional so a test
 * can describe exactly the browser it wants; `defaultChannelEnvironment()`
 * reads the real globals. */
export interface ChannelEnvironment {
  /** Opens a broadcast port, or absent when BroadcastChannel is not there. */
  broadcast?: (name: string) => BroadcastPort;
  storage?: ChannelStorage | null;
  /** Subscribes to other windows' storage writes; returns the unsubscribe. */
  onStorage?: (listener: (key: string | null, value: string | null) => void) => () => void;
  /** Milliseconds since the epoch, stamped as `at` on a storage write. */
  now?: () => number;
}

/** The storage key the fallback writes for channel `name`. */
export function channelStorageKey(name: string): string {
  return `docu-edit:channel:${name}`;
}

const PROBE_KEY = "docu-edit:channel:probe";

/** The real browser globals, guarded: reading `localStorage` can itself throw
 * when storage is blocked. */
export function defaultChannelEnvironment(): ChannelEnvironment {
  const env: ChannelEnvironment = { now: () => Date.now() };
  const Broadcast = typeof globalThis.BroadcastChannel === "function" ? globalThis.BroadcastChannel : null;
  if (Broadcast) {
    env.broadcast = (name) => new Broadcast(name);
  }
  try {
    env.storage = typeof globalThis.localStorage === "undefined" ? null : globalThis.localStorage;
  } catch {
    env.storage = null;
  }
  if (typeof globalThis.addEventListener === "function") {
    env.onStorage = (listener) => {
      const handler = (event: Event): void => {
        const storageEvent = event as StorageEvent;
        listener(storageEvent.key, storageEvent.newValue);
      };
      globalThis.addEventListener("storage", handler);
      return () => globalThis.removeEventListener("storage", handler);
    };
  }
  return env;
}

/** Whether the storage fallback can actually write (a blocked or full
 * storage throws on `setItem`). */
function storageWritable(storage: ChannelStorage | null | undefined): storage is ChannelStorage {
  if (!storage) return false;
  try {
    storage.setItem(PROBE_KEY, "1");
    storage.removeItem(PROBE_KEY);
    return true;
  } catch {
    return false;
  }
}

/** Which transport a channel opened in `env` would use. */
export function windowChannelTransport(env: ChannelEnvironment = defaultChannelEnvironment()): WindowChannelTransport {
  if (env.broadcast) return "broadcast";
  if (env.onStorage && storageWritable(env.storage)) return "storage";
  return "none";
}

/** The `seq` of the record currently stored under `key`, or 0. */
function storedSeq(storage: ChannelStorage, key: string): number {
  try {
    const record: unknown = JSON.parse(storage.getItem(key) ?? "null");
    const seq = record !== null && typeof record === "object" ? (record as { seq?: unknown }).seq : undefined;
    return typeof seq === "number" && Number.isFinite(seq) ? seq : 0;
  } catch {
    return 0;
  }
}

/** Opens channel `name`: `onMessage` receives every payload another window
 * posts on the same name (never this channel's own posts). `post` and `close`
 * never throw, and both are no-ops once closed. */
export function openWindowChannel(
  name: string,
  onMessage: (payload: unknown) => void,
  env: ChannelEnvironment = defaultChannelEnvironment(),
): WindowChannel {
  const transport = windowChannelTransport(env);
  let closed = false;

  if (transport === "broadcast" && env.broadcast) {
    const port = env.broadcast(name);
    const listener = (event: MessageEvent): void => {
      if (!closed) onMessage(event.data);
    };
    port.addEventListener("message", listener);
    return {
      post(payload) {
        if (closed) return;
        try {
          port.postMessage(payload);
        } catch {
          // An uncloneable payload is the caller's bug; the channel stays up.
        }
      },
      close() {
        if (closed) return;
        closed = true;
        port.removeEventListener("message", listener);
        port.close();
      },
    };
  }

  if (transport === "storage" && env.storage && env.onStorage) {
    const storage = env.storage;
    const key = channelStorageKey(name);
    let seq = 0;
    const unsubscribe = env.onStorage((changedKey, value) => {
      if (closed || changedKey !== key || value === null) return;
      let record: unknown;
      try {
        record = JSON.parse(value);
      } catch {
        return;
      }
      if (record !== null && typeof record === "object" && "payload" in record) {
        onMessage((record as { payload: unknown }).payload);
      }
    });
    return {
      post(payload) {
        if (closed) return;
        seq = Math.max(seq, storedSeq(storage, key)) + 1;
        try {
          storage.setItem(key, JSON.stringify({ seq, at: env.now ? env.now() : 0, payload }));
        } catch {
          // Storage filled up mid-session: the message is lost, never thrown.
        }
      },
      close() {
        if (closed) return;
        closed = true;
        unsubscribe();
        try {
          storage.removeItem(key);
        } catch {
          // ignore
        }
      },
    };
  }

  return {
    post() {},
    close() {
      closed = true;
    },
  };
}
