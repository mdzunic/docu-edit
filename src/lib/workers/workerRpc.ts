// Generic request/response RPC over a dedicated Worker (SPEC-047): every
// off-main-thread client (word count, spell check, export) shares this
// instead of hand-rolling postMessage bookkeeping. jsdom (Vitest) has no
// `Worker` global, so callers feature-detect with isWorkerSupported() and
// fall back to calling their pure function directly — the same
// typeof-guard idiom src/lib/idleScheduler.ts already uses for
// requestIdleCallback.

export function isWorkerSupported(): boolean {
  return typeof Worker === "function";
}

interface PendingCall<TResponse> {
  resolve: (value: TResponse) => void;
  reject: (reason: unknown) => void;
}

export interface WorkerRpcClient<TRequest, TResponse> {
  /** Sends `request` (tagged with a fresh correlation id) to the worker and
   * resolves with its reply. The worker is created lazily, on the first
   * call, not at client-construction time — opening a document never pays
   * for a worker its session turns out not to need. */
  call(request: TRequest, transfer?: Transferable[]): Promise<TResponse>;
  /** Tears down the worker (and rejects anything still in flight) — call on
   * view unmount so a closed document's worker doesn't linger. */
  terminate(): void;
}

/** `factory` constructs the worker; every request/response is a plain
 * object carrying a numeric `id` the worker echoes back unchanged so
 * out-of-order replies still resolve the right caller. */
export function createWorkerRpcClient<TRequest extends object, TResponse>(
  factory: () => Worker,
): WorkerRpcClient<TRequest, TResponse> {
  let worker: Worker | null = null;
  let nextId = 1;
  const pending = new Map<number, PendingCall<TResponse>>();

  function ensureWorker(): Worker {
    if (worker) {
      return worker;
    }
    const instance = factory();
    instance.addEventListener("message", (event: MessageEvent) => {
      const data = event.data as { id: number } & Record<string, unknown>;
      const entry = pending.get(data.id);
      if (!entry) {
        return;
      }
      pending.delete(data.id);
      const { id: _id, ...rest } = data;
      entry.resolve(rest as unknown as TResponse);
    });
    instance.addEventListener("error", (event: ErrorEvent) => {
      for (const entry of pending.values()) {
        entry.reject(event.error ?? new Error(event.message));
      }
      pending.clear();
    });
    worker = instance;
    return instance;
  }

  return {
    call(request, transfer = []) {
      const id = nextId;
      nextId += 1;
      return new Promise<TResponse>((resolve, reject) => {
        pending.set(id, { resolve, reject });
        try {
          ensureWorker().postMessage({ id, ...request }, transfer);
        } catch (error) {
          pending.delete(id);
          reject(error);
        }
      });
    },
    terminate() {
      worker?.terminate();
      worker = null;
      pending.clear();
    },
  };
}
