// SPEC-137 AC-46: base64ToBytesChunked must decode to exactly the same bytes
// as the synchronous base64ToBytes, across slice boundaries and every
// padding case, and must yield (yieldToMainThread) between slices.
import { describe, expect, it, vi } from "vitest";

vi.mock("./idleScheduler", () => ({ yieldToMainThread: vi.fn(() => Promise.resolve()) }));

import { yieldToMainThread } from "./idleScheduler";
import { base64ToBytes, base64ToBytesChunked, bytesToBase64 } from "./packageParts";

const yieldSpy = yieldToMainThread as unknown as ReturnType<typeof vi.fn>;

describe("base64ToBytesChunked", () => {
  it("decodes empty input the same as base64ToBytes", async () => {
    expect(await base64ToBytesChunked("")).toEqual(base64ToBytes(""));
  });

  it("matches base64ToBytes for input shorter than one slice (no yield)", async () => {
    yieldSpy.mockClear();
    const base64 = bytesToBase64(new TextEncoder().encode("hello"));
    expect(await base64ToBytesChunked(base64, 1024)).toEqual(base64ToBytes(base64));
    expect(yieldSpy).not.toHaveBeenCalled();
  });

  it("matches base64ToBytes across many slice boundaries with a small maxSliceChars", async () => {
    for (const length of [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 23, 24, 25, 100]) {
      const bytes = new Uint8Array(length);
      for (let i = 0; i < length; i++) {
        bytes[i] = i % 256;
      }
      const base64 = bytesToBase64(bytes);
      for (const maxSliceChars of [8, 12]) {
        expect(await base64ToBytesChunked(base64, maxSliceChars)).toEqual(base64ToBytes(base64));
      }
    }
  });

  it("handles every padding case (0, 1 and 2 '=' characters)", async () => {
    for (const length of [3, 4, 5]) {
      // 3 bytes -> no padding, 4 bytes -> "==" (1 extra byte), 5 bytes -> "=" (2 extra bytes)
      const bytes = new Uint8Array(length).fill(7);
      const base64 = bytesToBase64(bytes);
      expect(await base64ToBytesChunked(base64, 8)).toEqual(base64ToBytes(base64));
    }
  });

  it("matches base64ToBytes for a large random input", async () => {
    const bytes = new Uint8Array(500_000);
    for (let i = 0; i < bytes.length; i++) {
      bytes[i] = (i * 2654435761) % 256;
    }
    const base64 = bytesToBase64(bytes);
    expect(await base64ToBytesChunked(base64)).toEqual(base64ToBytes(base64));
    expect(await base64ToBytesChunked(base64, 1024)).toEqual(base64ToBytes(base64));
  });

  it("yields between slices, but not after the last one", async () => {
    yieldSpy.mockClear();
    const base64 = bytesToBase64(new Uint8Array(100));
    // 100 bytes -> ~136 base64 chars; maxSliceChars=12 -> 12 slices -> 11 yields.
    await base64ToBytesChunked(base64, 12);
    const expectedSlices = Math.ceil(base64.length / 12);
    expect(yieldSpy).toHaveBeenCalledTimes(expectedSlices - 1);
  });

  it("rounds maxSliceChars down to a multiple of 4 and never loops forever on a tiny value", async () => {
    const base64 = bytesToBase64(new Uint8Array(40));
    expect(await base64ToBytesChunked(base64, 1)).toEqual(base64ToBytes(base64));
  });
});
