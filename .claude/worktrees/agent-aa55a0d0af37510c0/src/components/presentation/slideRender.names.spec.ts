import { describe, it, expect } from "vitest";
import { frameAccessibleName } from "./slideRender";
import type { SlideFrame } from "../../lib/presentation/model";

// AC-52: every top-level frame is exposed with an accessible name. Placeholders
// all shared one — a slide's title and its body were indistinguishable to
// anything reading the name, while the screen plainly said "Click to add title"
// and "Click to add text".
function textFrame(role?: "title" | "subtitle" | "body" | "other"): SlideFrame {
  const frame = { kind: "text", id: "f", xCm: 0, yCm: 0, widthCm: 4, heightCm: 2, paragraphs: [] } as unknown as Record<string, unknown>;
  if (role) frame.role = role;
  return frame as unknown as SlideFrame;
}

describe("frameAccessibleName — placeholders", () => {
  it("names each placeholder by its role", () => {
    expect(frameAccessibleName(textFrame("title"))).toBe("Title placeholder");
    expect(frameAccessibleName(textFrame("subtitle"))).toBe("Subtitle placeholder");
    expect(frameAccessibleName(textFrame("body"))).toBe("Body placeholder");
  });

  it("gives the roles on one slide different names", () => {
    const names = (["title", "subtitle", "body"] as const).map((r) => frameAccessibleName(textFrame(r)));
    expect(new Set(names).size).toBe(names.length);
  });

  it("falls back for an unclassified placeholder, and keeps plain text boxes", () => {
    expect(frameAccessibleName(textFrame("other"))).toBe("Placeholder");
    expect(frameAccessibleName(textFrame())).toBe("Text box");
  });
});
