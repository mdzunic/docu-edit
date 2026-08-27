// SPEC-088 — the self-contained HTML export serialiser. Node-only unit tests
// (no DOM): escaping of untrusted content, determinism, self-containment (no
// external requests), notes inclusion/exclusion at the byte level, the
// transition/animation data attributes, key-table equality with playbackKeys,
// size accounting, and no-mutation of a deep-frozen model.
import { describe, expect, it } from "vitest";
import { exportPresentationHtml, htmlFileName, keyTableJson, esc, validColor } from "./htmlExport";
import { advanceKeys, backKeys, firstKeys, lastKeys } from "./playbackKeys";
import type { PresentationModel, Slide, SlideFrame } from "./model";

function textFrame(id: string, text: string, extra: Partial<Extract<SlideFrame, { kind: "text" }>> = {}): SlideFrame {
  return {
    id,
    kind: "text",
    xCm: 1,
    yCm: 1,
    widthCm: 10,
    heightCm: 3,
    paragraphs: [{ runs: [{ text }] }],
    ...extra,
  };
}

function baseModel(slides: Slide[], extra: Partial<PresentationModel> = {}): PresentationModel {
  return {
    kind: "presentation",
    id: "deck-1",
    title: "My Deck",
    format: "pptx",
    slides,
    slideWidthCm: 25.4,
    slideHeightCm: 14.29,
    ...extra,
  };
}

function slide(id: string, frames: SlideFrame[], extra: Partial<Slide> = {}): Slide {
  return { id, frames, ...extra };
}

describe("exportPresentationHtml — document shape", () => {
  it("is a single self-contained HTML document with inline CSS and JS", () => {
    const html = exportPresentationHtml(baseModel([slide("s1", [textFrame("f1", "Hello")])]), { includeNotes: false });
    expect(html.startsWith("<!doctype html>")).toBe(true);
    expect(html).toContain("<html lang=\"en\">");
    expect(html).toContain("<style>");
    expect(html).toContain("<script>");
    // No external stylesheet/script references.
    expect(html).not.toContain("<link");
    expect(html).not.toContain("src=\"http");
  });

  it("makes no external requests: no http(s) URL anywhere in the bytes", () => {
    const html = exportPresentationHtml(baseModel([slide("s1", [textFrame("f1", "Hi")])]), { includeNotes: false });
    expect(/https?:\/\//i.test(html)).toBe(false);
  });

  it("sets the escaped deck title as <title> and the stage aria-label", () => {
    const html = exportPresentationHtml(baseModel([slide("s1", [textFrame("f1", "x")])], { title: "A & B <c>" }), {
      includeNotes: false,
    });
    expect(html).toContain("<title>A &amp; B &lt;c&gt;</title>");
    expect(html).toContain("aria-label=\"A &amp; B &lt;c&gt;\"");
  });

  it("throws on a deck with no slides", () => {
    expect(() => exportPresentationHtml(baseModel([]), { includeNotes: false })).toThrow();
  });

  it("renders one section per slide with the end screen text", () => {
    const html = exportPresentationHtml(baseModel([slide("s1", []), slide("s2", [])]), { includeNotes: false });
    expect(html.match(/<section /g)?.length).toBe(2);
    expect(html).toContain("End of slide show — click or press any key to exit");
  });
});

describe("exportPresentationHtml — escaping (AC-48/49)", () => {
  it("entity-escapes a script-injection payload so it renders as text", () => {
    const payload = "</script><script>alert(1)</script>";
    const html = exportPresentationHtml(baseModel([slide("s1", [textFrame("f1", payload)])]), { includeNotes: false });
    // The raw payload must NOT appear; its escaped form must.
    expect(html).not.toContain("<script>alert(1)</script>");
    expect(html).toContain("&lt;/script&gt;&lt;script&gt;alert(1)&lt;/script&gt;");
  });

  it("escapes quotes and angle brackets in the esc helper", () => {
    expect(esc("a\"b'c<d>&e")).toBe("a&quot;b&#39;c&lt;d&gt;&amp;e");
  });

  it("re-validates colours against #rrggbb", () => {
    expect(validColor("#AABBCC")).toBe("#aabbcc");
    expect(validColor("red")).toBeUndefined();
    expect(validColor("#12")).toBeUndefined();
    expect(validColor("#abcdef; x:y")).toBeUndefined();
  });

  it("drops an invalid run colour and an unsafe font family from the emitted CSS", () => {
    const frame = textFrame("f1", "styled", {});
    (frame as Extract<SlideFrame, { kind: "text" }>).paragraphs = [
      { runs: [{ text: "styled", color: "javascript:alert(1)", fontFamily: "Arial\"; content:'x" }] },
    ];
    const html = exportPresentationHtml(baseModel([slide("s1", [frame])]), { includeNotes: false });
    expect(html).not.toContain("javascript:alert(1)");
    expect(html).toContain("color:#000000");
    // Family reduced to a safe token set, still leading the system stack.
    expect(html).toContain("font-family:&#39;Arial contentx&#39;, system-ui");
  });
});

describe("exportPresentationHtml — determinism (AC-51)", () => {
  it("produces identical output across two calls on the same model", () => {
    const model = baseModel([
      slide("s1", [textFrame("f1", "one")], { transition: { type: "push", direction: "left" } }),
      slide("s2", [textFrame("f2", "two")], {
        animations: [{ frameId: "f2", effect: "fadeIn", trigger: "click" }],
      }),
    ]);
    expect(exportPresentationHtml(model, { includeNotes: true })).toBe(exportPresentationHtml(model, { includeNotes: true }));
  });
});

describe("exportPresentationHtml — notes (AC-39/41/42)", () => {
  const model = baseModel([slide("s1", [textFrame("f1", "x")], { notes: ["Secret speaker note"] })]);

  it("includes note text only when includeNotes is on", () => {
    const on = exportPresentationHtml(model, { includeNotes: true });
    expect(on).toContain("Secret speaker note");
    expect(on).toContain("id=\"notes\"");
    expect(on).toContain("data-notes=\"1\"");
  });

  it("omits note text entirely when includeNotes is off (AC-42)", () => {
    const off = exportPresentationHtml(model, { includeNotes: false });
    expect(off).not.toContain("Secret speaker note");
    expect(off).not.toContain("id=\"notes\"");
    expect(off).toContain("data-notes=\"0\"");
  });
});

describe("exportPresentationHtml — transition/animation data (AC-29/33)", () => {
  it("emits the transition attributes from the SPEC-085 model", () => {
    const html = exportPresentationHtml(
      baseModel([
        slide("s1", []),
        slide("s2", [], { transition: { type: "push", direction: "right", durationMs: 800, advanceAfterMs: 4000 } }),
      ]),
      { includeNotes: false },
    );
    expect(html).toContain("data-transition=\"push\"");
    expect(html).toContain("data-transition-dir=\"right\"");
    expect(html).toContain("data-transition-ms=\"800\"");
    expect(html).toContain("data-advance-after=\"4000\"");
  });

  it("emits per-frame animation data grouped into builds", () => {
    const f1 = textFrame("f1", "a");
    const f2 = textFrame("f2", "b");
    const html = exportPresentationHtml(
      baseModel([
        slide("s1", [f1, f2], {
          animations: [
            { frameId: "f1", effect: "fadeIn", trigger: "click" },
            { frameId: "f2", effect: "flyIn", trigger: "click", direction: "up", durationMs: 700, delayMs: 100 },
          ],
        }),
      ]),
      { includeNotes: false },
    );
    expect(html).toContain("data-builds=\"2\"");
    // The data-anim JSON carries effect/dir/ms/delay/trigger/build/order per anim.
    const m = html.match(/data-anim="([^"]*)"/);
    expect(m).toBeTruthy();
    const json = JSON.parse((m![1]).replace(/&quot;/g, "\"").replace(/&amp;/g, "&"));
    expect(json).toHaveLength(2);
    expect(json[0]).toMatchObject({ effect: "fadeIn", build: 0, order: 0 });
    expect(json[1]).toMatchObject({ effect: "flyIn", dir: "up", ms: 700, delay: 100, build: 1, order: 1 });
    // Both frame ids resolved to emitted data-fid values.
    expect(json[0].fid).toMatch(/^f\d+$/);
  });
});

describe("exportPresentationHtml — frame rendering", () => {
  it("embeds an image as a data: URI (AC-6/12)", () => {
    const image: SlideFrame = { id: "img1", kind: "image", xCm: 0, yCm: 0, widthCm: 4, heightCm: 4, mediaId: "m1" };
    const html = exportPresentationHtml(
      baseModel([slide("s1", [image])], { media: [{ mediaId: "m1", contentType: "image/png", data: "AAAA" }] }),
      { includeNotes: false },
    );
    expect(html).toContain("src=\"data:image/png;base64,AAAA\"");
    expect(html).toContain("alt=\"\"");
  });

  it("renders a shape frame as inline SVG from the catalogue geometry (AC-13)", () => {
    const shapeFrame: SlideFrame = {
      id: "sh1",
      kind: "shape",
      xCm: 0,
      yCm: 0,
      widthCm: 4,
      heightCm: 4,
      geometry: "rect",
      fill: { type: "solid", color: "#8faadc" },
    };
    const html = exportPresentationHtml(baseModel([slide("s1", [shapeFrame])]), { includeNotes: false });
    expect(html).toContain("<svg class=\"shape-svg\"");
    expect(html).toContain("fill=\"#8faadc\"");
    expect(html).toContain("<path d=\"M 0 0 L");
  });

  it("renders a foreign frame as its labelled placeholder box (AC-15)", () => {
    const foreign: SlideFrame = { id: "fr1", kind: "foreign", xCm: 0, yCm: 0, widthCm: 4, heightCm: 4, xml: "<x/>", label: "Chart" };
    const html = exportPresentationHtml(baseModel([slide("s1", [foreign])]), { includeNotes: false });
    expect(html).toContain("class=\"frame foreign\"");
    expect(html).toContain(">Chart<");
  });

  it("renders group children (AC-14)", () => {
    const group: SlideFrame = {
      id: "g1",
      kind: "group",
      xCm: 0,
      yCm: 0,
      widthCm: 8,
      heightCm: 8,
      xml: "<g/>",
      children: [textFrame("c1", "grouped text")],
    };
    const html = exportPresentationHtml(baseModel([slide("s1", [group])]), { includeNotes: false });
    expect(html).toContain("class=\"frame group\"");
    expect(html).toContain("grouped text");
  });

  it("renders bullet markers and level indent for a bulleted paragraph (AC-11)", () => {
    const frame = textFrame("f1", "item", {});
    (frame as Extract<SlideFrame, { kind: "text" }>).paragraphs = [{ runs: [{ text: "item" }], bullet: "disc", level: 2 }];
    const html = exportPresentationHtml(baseModel([slide("s1", [frame])]), { includeNotes: false });
    expect(html).toContain("class=\"marker\"");
    expect(html).toContain("•");
    expect(html).toContain("margin-left:");
  });
});

describe("exportPresentationHtml — key table equality (AC-20/21)", () => {
  it("embeds the playbackKeys tables verbatim", () => {
    const html = exportPresentationHtml(baseModel([slide("s1", [])]), { includeNotes: false });
    const expected = JSON.stringify({ advance: advanceKeys, back: backKeys, first: firstKeys, last: lastKeys });
    expect(keyTableJson()).toBe(expected);
    expect(html).toContain("var KEYS=" + expected);
  });
});

describe("exportPresentationHtml — size accounting (AC-52)", () => {
  it("stays comfortably under 16MB for ~1MB of media (base64 overhead ~4/3)", () => {
    const data = "A".repeat(1_000_000); // ~1MB of base64 payload
    const image: SlideFrame = { id: "img1", kind: "image", xCm: 0, yCm: 0, widthCm: 4, heightCm: 4, mediaId: "m1" };
    const html = exportPresentationHtml(
      baseModel([slide("s1", [image])], { media: [{ mediaId: "m1", contentType: "image/png", data }] }),
      { includeNotes: false },
    );
    expect(html.length).toBeLessThan(16 * 1024 * 1024);
    expect(html.length).toBeGreaterThan(1_000_000);
  });
});

describe("exportPresentationHtml — purity (AC-54)", () => {
  it("does not mutate a deep-frozen model", () => {
    const model = baseModel([
      slide("s1", [textFrame("f1", "x")], {
        transition: { type: "fade" },
        animations: [{ frameId: "f1", effect: "fadeIn", trigger: "click" }],
        notes: ["n"],
      }),
    ]);
    deepFreeze(model);
    expect(() => exportPresentationHtml(model, { includeNotes: true })).not.toThrow();
  });
});

describe("htmlFileName", () => {
  it("appends .html and falls back on a blank title", () => {
    expect(htmlFileName("My Deck")).toBe("My Deck.html");
    expect(htmlFileName("   ")).toBe("Untitled presentation.html");
  });
});

function deepFreeze<T>(value: T): T {
  if (value && typeof value === "object") {
    Object.freeze(value);
    for (const key of Object.keys(value as object)) {
      deepFreeze((value as Record<string, unknown>)[key]);
    }
  }
  return value;
}
