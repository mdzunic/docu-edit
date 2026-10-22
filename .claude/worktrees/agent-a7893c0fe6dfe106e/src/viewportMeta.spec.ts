import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

// SPEC-068 AC-45/46 — the app must never disable browser zoom: setting
// `user-scalable=no` or a `maximum-scale` that pins scaling fails accessibility
// outright, and pinch-to-zoom of the document (AC-32) goes through SPEC-063's
// view transform rather than the browser's page zoom precisely so the page
// zoom can stay available for users who need it. This guards the one place the
// meta tag lives so a later change cannot quietly reintroduce the lock.
describe("viewport meta (SPEC-068 AC-45/46)", () => {
  const html = readFileSync(
    join(dirname(fileURLToPath(import.meta.url)), "..", "index.html"),
    "utf8",
  );
  const viewport = html.match(/<meta\s+name="viewport"\s+content="([^"]*)"/i)?.[1] ?? "";

  it("declares a viewport meta tag", () => {
    expect(viewport).not.toBe("");
    expect(viewport).toContain("width=device-width");
  });

  it("never sets user-scalable=no", () => {
    expect(viewport.replace(/\s/g, "")).not.toContain("user-scalable=no");
    expect(viewport).not.toMatch(/user-scalable\s*=\s*0/i);
  });

  it("never disables zoom with a maximum-scale", () => {
    expect(viewport).not.toMatch(/maximum-scale/i);
  });
});
