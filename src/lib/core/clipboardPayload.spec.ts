// SPEC-099 phase one — the clipboard envelope and guards hoisted out of
// lib/clipboard.ts (whose own spec still passes untouched).
import { describe, expect, it } from "vitest";
import { buildClipboardEnvelope, isFiniteNumber, isSafeColor, parseClipboardEnvelope } from "./clipboardPayload";

const MIME = "application/x-test+json";

describe("buildClipboardEnvelope", () => {
  it("labels the body with mimeType and version first", () => {
    expect(buildClipboardEnvelope(MIME, 1, { blocks: [1] })).toBe(`{"mimeType":"${MIME}","version":1,"blocks":[1]}`);
  });
});

describe("parseClipboardEnvelope", () => {
  it("returns the parsed record for a matching label", () => {
    const raw = buildClipboardEnvelope(MIME, 1, { body: "x" });
    expect(parseClipboardEnvelope(raw, MIME, 1)).toEqual({ mimeType: MIME, version: 1, body: "x" });
  });

  it("is null for non-JSON, non-objects, arrays and a wrong mimeType or version", () => {
    expect(parseClipboardEnvelope("not json", MIME, 1)).toBeNull();
    expect(parseClipboardEnvelope("42", MIME, 1)).toBeNull();
    expect(parseClipboardEnvelope("null", MIME, 1)).toBeNull();
    expect(parseClipboardEnvelope(`[{"mimeType":"${MIME}","version":1}]`, MIME, 1)).toBeNull();
    expect(parseClipboardEnvelope(buildClipboardEnvelope("text/plain", 1, {}), MIME, 1)).toBeNull();
    expect(parseClipboardEnvelope(buildClipboardEnvelope(MIME, 2, {}), MIME, 1)).toBeNull();
    expect(parseClipboardEnvelope(`{"mimeType":"${MIME}","version":"1"}`, MIME, 1)).toBeNull();
  });
});

describe("guards", () => {
  it("isFiniteNumber accepts only real finite numbers", () => {
    expect(isFiniteNumber(3.5)).toBe(true);
    expect(isFiniteNumber(Number.NaN)).toBe(false);
    expect(isFiniteNumber(Number.POSITIVE_INFINITY)).toBe(false);
    expect(isFiniteNumber("3")).toBe(false);
  });

  it("isSafeColor accepts hex and bare keywords only", () => {
    expect(isSafeColor("#a1b2c3")).toBe(true);
    expect(isSafeColor("#abc")).toBe(true);
    expect(isSafeColor("red")).toBe(true);
    expect(isSafeColor("url(javascript:alert(1))")).toBe(false);
    expect(isSafeColor("#12345g")).toBe(false);
    expect(isSafeColor(7)).toBe(false);
  });
});
