import { describe, expect, it } from "vitest";
import * as links from "../links";
import { DANGEROUS_SCHEMES, detectBareUrl, isDangerousUrl, sanitizeHref, urlScheme } from "./urls";

describe("core/urls (SPEC-097's move of SPEC-025's URL helpers)", () => {
  it("parses schemes the forgiving way browsers do", () => {
    expect(urlScheme("  java\tscript:alert(1)")).toBe("javascript");
    expect(urlScheme("https://x.test")).toBe("https");
    expect(urlScheme("relative/path")).toBeNull();
  });

  it("rejects the dangerous schemes and trims the rest", () => {
    expect([...DANGEROUS_SCHEMES].sort()).toEqual(["data", "javascript", "vbscript"]);
    expect(isDangerousUrl("DATA:text/html,x")).toBe(true);
    expect(sanitizeHref("vbscript:x")).toBeNull();
    expect(sanitizeHref("   ")).toBeNull();
    expect(sanitizeHref(" https://example.com ")).toBe("https://example.com");
  });

  it("detects bare URLs with punctuation peeled off", () => {
    expect(detectBareUrl("see www.example.com.", 20)).toEqual({ start: 4, end: 19, href: "https://www.example.com" });
    expect(detectBareUrl("plain words", 11)).toBeNull();
  });

  it("phase one: src/lib/links.ts re-exports the very same functions", () => {
    expect(links.sanitizeHref).toBe(sanitizeHref);
    expect(links.isDangerousUrl).toBe(isDangerousUrl);
    expect(links.urlScheme).toBe(urlScheme);
    expect(links.detectBareUrl).toBe(detectBareUrl);
    expect(links.DANGEROUS_SCHEMES).toBe(DANGEROUS_SCHEMES);
  });
});
