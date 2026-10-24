// SPEC-098 — which frames a frame-level text session opens on, and how an
// unaligned paragraph lays out there: a text frame's paragraphs, or a shape's
// text (a find hit inside a shape opens a session on it), centred as the stage
// centres it.
import { describe, expect, it } from "vitest";
import type { SlideFrame } from "../../lib/presentation/model";
import { frameTextOwner, unalignedAlign } from "./slideTextSession";

const box = { xCm: 0, yCm: 0, widthCm: 4, heightCm: 2 };
const paragraphs = [{ runs: [{ text: "Press me" }] }];
const text = { id: "t", kind: "text", ...box, paragraphs } as SlideFrame;
const shape = { id: "s", kind: "shape", geometry: "ellipse", ...box, paragraphs } as SlideFrame;
const bareShape = { id: "b", kind: "shape", geometry: "rect", ...box } as SlideFrame;
const table = { id: "tb", kind: "table", ...box, columnWidthsCm: [4], rows: [] } as unknown as SlideFrame;

describe("frameTextOwner", () => {
  it("is the frame itself for a text frame and for a shape holding text", () => {
    expect(frameTextOwner(text)).toBe(text);
    expect(frameTextOwner(shape)).toBe(shape);
  });

  it("is null for a shape with no text, a table (its cells own the text) and no frame", () => {
    expect(frameTextOwner(bareShape)).toBeNull();
    expect(frameTextOwner(table)).toBeNull();
    expect(frameTextOwner(null)).toBeNull();
  });
});

describe("unalignedAlign", () => {
  it("centres a shape's text and leaves everything else left-aligned", () => {
    expect(unalignedAlign(shape)).toBe("center");
    expect(unalignedAlign(text)).toBe("left");
    expect(unalignedAlign(null)).toBe("left");
  });
});
