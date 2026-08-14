import { describe, it, expect } from "vitest";
import { createWorkerRpcClient, isWorkerSupported } from "./workerRpc";

type Listener = (event: { data?: unknown; message?: string; error?: unknown }) => void;

class FakeWorker {
  listeners = new Map<string, Listener[]>();
  posted: unknown[] = [];
  terminated = false;

  addEventListener(type: string, callback: Listener): void {
    const list = this.listeners.get(type) ?? [];
    list.push(callback);
    this.listeners.set(type, list);
  }

  postMessage(message: unknown): void {
    this.posted.push(message);
  }

  terminate(): void {
    this.terminated = true;
  }

  emit(type: string, event: { data?: unknown; message?: string; error?: unknown }): void {
    for (const callback of this.listeners.get(type) ?? []) {
      callback(event);
    }
  }
}

describe("isWorkerSupported", () => {
  it("is false under jsdom, which has no global Worker constructor", () => {
    expect(isWorkerSupported()).toBe(false);
  });
});

describe("createWorkerRpcClient", () => {
  it("tags a call with a fresh id and resolves it from the matching reply", async () => {
    const fake = new FakeWorker();
    const client = createWorkerRpcClient<{ n: number }, { doubled: number }>(
      () => fake as unknown as Worker,
    );
    const promise = client.call({ n: 21 });
    expect(fake.posted).toEqual([{ id: 1, n: 21 }]);
    fake.emit("message", { data: { id: 1, doubled: 42 } });
    await expect(promise).resolves.toEqual({ doubled: 42 });
  });

  it("resolves out-of-order replies to the caller whose id matches, not arrival order", async () => {
    const fake = new FakeWorker();
    const client = createWorkerRpcClient<{ n: number }, { doubled: number }>(
      () => fake as unknown as Worker,
    );
    const first = client.call({ n: 1 });
    const second = client.call({ n: 2 });
    fake.emit("message", { data: { id: 2, doubled: 4 } });
    fake.emit("message", { data: { id: 1, doubled: 2 } });
    await expect(second).resolves.toEqual({ doubled: 4 });
    await expect(first).resolves.toEqual({ doubled: 2 });
  });

  it("rejects every pending call when the worker errors", async () => {
    const fake = new FakeWorker();
    const client = createWorkerRpcClient<{ n: number }, unknown>(() => fake as unknown as Worker);
    const a = client.call({ n: 1 });
    const b = client.call({ n: 2 });
    fake.emit("error", { message: "boom" });
    await expect(a).rejects.toThrow("boom");
    await expect(b).rejects.toThrow("boom");
  });

  it("creates the worker lazily, only on the first call", () => {
    let created = 0;
    const client = createWorkerRpcClient<{ n: number }, unknown>(() => {
      created += 1;
      return new FakeWorker() as unknown as Worker;
    });
    expect(created).toBe(0);
    void client.call({ n: 1 });
    expect(created).toBe(1);
    void client.call({ n: 2 });
    expect(created).toBe(1);
  });

  it("terminate tears the worker down", () => {
    const fake = new FakeWorker();
    const client = createWorkerRpcClient<{ n: number }, unknown>(() => fake as unknown as Worker);
    void client.call({ n: 1 });
    client.terminate();
    expect(fake.terminated).toBe(true);
  });

  it("drops an errored worker and constructs a fresh one on the next call", async () => {
    const workers: FakeWorker[] = [];
    const client = createWorkerRpcClient<{ n: number }, { doubled: number }>(() => {
      const fake = new FakeWorker();
      workers.push(fake);
      return fake as unknown as Worker;
    });
    const failed = client.call({ n: 1 });
    workers[0].emit("error", { message: "boom" });
    await expect(failed).rejects.toThrow("boom");
    expect(workers[0].terminated).toBe(true);

    // A dead-but-cached worker here would swallow this call forever.
    const retried = client.call({ n: 2 });
    expect(workers).toHaveLength(2);
    workers[1].emit("message", { data: { id: 2, doubled: 4 } });
    await expect(retried).resolves.toEqual({ doubled: 4 });
  });

  it("rejects pending calls on a messageerror instead of stranding them", async () => {
    const fake = new FakeWorker();
    const client = createWorkerRpcClient<{ n: number }, unknown>(() => fake as unknown as Worker);
    const pending = client.call({ n: 1 });
    fake.emit("messageerror", {});
    await expect(pending).rejects.toThrow("deserialized");
  });

  it("terminate rejects in-flight calls, honouring its settle contract", async () => {
    const fake = new FakeWorker();
    const client = createWorkerRpcClient<{ n: number }, unknown>(() => fake as unknown as Worker);
    const pending = client.call({ n: 1 });
    client.terminate();
    await expect(pending).rejects.toThrow("terminated");
  });
});
