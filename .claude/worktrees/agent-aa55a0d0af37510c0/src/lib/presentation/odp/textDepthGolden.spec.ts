// Pins SPEC-097's golden-pending ODF details (textDepthGolden.ts): each is
// recorded with its two candidate forms, the writer omits it, and the reader
// preserves a source's markup in either form. When a golden lands its bytes
// are asserted here byte-exact in both directions instead.
import { describe, expect, it } from "vitest";
import type { SlideFrame, TextFrame } from "../model";
import { parseFrameChildForTest } from "./reader";
import { serializeOdpGroupFrame } from "./serialize";
import { ODP_NUMBERING_RESTART, ODP_SHRINK_ON_OVERFLOW, ODP_SLIDE_JUMP_LINK, ODP_TEXT_DEPTH_PENDING } from "./textDepthGolden";

function textFrame(paragraphs: TextFrame["paragraphs"], extra: Partial<TextFrame> = {}): TextFrame {
  return { id: "t", kind: "text", xCm: 1, yCm: 1, widthCm: 10, heightCm: 4, paragraphs, ...extra };
}

function write(frame: SlideFrame): { frame: string; autoStyles: string } {
  return serializeOdpGroupFrame(structuredClone(frame));
}

describe("the golden-pending record (SPEC-097 odp)", () => {
  it("records all three details as pending Impress verifications with two named candidates", () => {
    expect(ODP_TEXT_DEPTH_PENDING).toEqual([ODP_SHRINK_ON_OVERFLOW, ODP_NUMBERING_RESTART, ODP_SLIDE_JUMP_LINK]);
    for (const entry of ODP_TEXT_DEPTH_PENDING) {
      expect(entry.status).toBe("pending");
      expect(entry.verification).toBe("Impress");
      expect(entry.candidates).toHaveLength(2);
      expect(entry.candidates[0]).not.toBe(entry.candidates[1]);
    }
    expect(ODP_SHRINK_ON_OVERFLOW.candidates.join(" | ")).toMatch(/draw:fit-to-size.*\| style:shrink-to-fit/);
    expect(ODP_NUMBERING_RESTART.candidates.join(" | ")).toMatch(/text:continue-numbering.*\| text:start-value/);
    expect(ODP_SLIDE_JUMP_LINK.candidates.join(" | ")).toMatch(/xlink:href="#\{draw:name\}".*\| .*presentation:action="jump"/);
  });
});

describe("the writer omits every golden-pending property", () => {
  it("writes no shrink-on-overflow attribute — a shrink-only body mints no graphic style", () => {
    const out = write(textFrame([{ runs: [{ text: "Long title" }] }], { body: { autofit: "shrink", fontScalePct: 62.5, lineScalePct: 90 } }));
    expect(out.frame).not.toContain("draw:style-name");
    expect(out.autoStyles).toBe("");
    // Beside a modelled body property, still neither candidate.
    const withAnchor = write(textFrame([{ runs: [{ text: "x" }] }], { body: { autofit: "shrink", anchor: "bottom" } }));
    expect(withAnchor.autoStyles).toContain(`draw:textarea-vertical-align="bottom"`);
    expect(withAnchor.autoStyles).not.toMatch(/fit-to-size|shrink-to-fit/);
  });

  it("writes no numbering restart: only the level style's start-value, from the level's first item", () => {
    const out = write(
      textFrame([
        { runs: [{ text: "a" }], bullet: "number", startAt: 4 },
        { runs: [{ text: "b" }], bullet: "number", startAt: 7 },
      ]),
    );
    expect(out.frame).not.toContain("continue-numbering");
    expect(out.frame).not.toMatch(/<text:list-item [^>]*start-value/);
    expect(out.autoStyles.match(/text:start-value="\d+"/g)).toEqual([`text:start-value="4"`]);
    // One list: the restart was not turned into a second list either.
    expect(out.frame.match(/<text:list /g)).toHaveLength(1);
  });

  it("writes no slide-jump link for a run, a frame or an in-deck #… href", () => {
    const out = write(
      textFrame(
        [{ runs: [{ text: "see " }, { text: "slide 2", linkSlide: 2, linkTooltip: "Go" }, { text: " or ", linkHref: "#page3" }] }],
        { linkSlide: 3 },
      ),
    );
    expect(out.frame).not.toContain("text:a");
    expect(out.frame).not.toContain("draw:a");
    expect(out.frame).not.toContain("#");
    expect(out.frame).not.toContain("office:title");
  });
});

describe("the reader preserves golden-pending source markup", () => {
  it("keeps both shrink candidates in odpBodyProps (autofit unmodelled) and re-emits them", () => {
    const styles =
      `<style:style style:name="grShrink" style:family="graphic">` +
      `<style:graphic-properties draw:fit-to-size="shrink-to-fit" style:shrink-to-fit="true"/></style:style>`;
    const frame = parseFrameChildForTest(
      `<draw:frame draw:style-name="grShrink" svg:x="1cm" svg:y="1cm" svg:width="10cm" svg:height="2cm"><draw:text-box><text:p>Title</text:p></draw:text-box></draw:frame>`,
      styles,
    ) as TextFrame;
    expect(frame.body).toBeUndefined();
    expect(frame.odpBodyProps).toBe(`draw:fit-to-size="shrink-to-fit" style:shrink-to-fit="true"`);
    expect(write(frame).autoStyles).toBe(
      `<style:style style:name="de-G1" style:family="graphic"><style:graphic-properties draw:fit-to-size="shrink-to-fit" style:shrink-to-fit="true"/></style:style>`,
    );
  });

  it("reads an in-deck #… text:a as plain text and keeps a #… draw:a verbatim", () => {
    const text = parseFrameChildForTest(
      `<draw:frame svg:x="1cm" svg:y="1cm" svg:width="10cm" svg:height="2cm"><draw:text-box><text:p>Go to <text:a xlink:href="#page3" xlink:type="simple">results</text:a></text:p></draw:text-box></draw:frame>`,
    ) as TextFrame;
    expect(text.paragraphs[0].runs).toEqual([{ text: "Go to " }, { text: "results" }]);
    const jump = parseFrameChildForTest(
      `<draw:a xlink:href="#page3" xlink:type="simple"><draw:frame svg:x="1cm" svg:y="1cm" svg:width="10cm" svg:height="2cm"><draw:text-box><text:p>Box</text:p></draw:text-box></draw:frame></draw:a>`,
    );
    expect(jump?.kind).toBe("foreign");
    // (The test DOM's serialiser may re-prefix xlink, so match the local name.)
    expect(jump?.kind === "foreign" && jump.xml).toMatch(/:href="#page3"/);
    expect([jump?.xCm, jump?.yCm, jump?.widthCm, jump?.heightCm]).toEqual([1, 1, 10, 2]);
  });

  it("reads a per-item text:start-value as that item's startAt and ignores text:continue-numbering", () => {
    const styles = `<text:list-style style:name="Lnum"><text:list-level-style-number text:level="1" style:num-format="1" style:num-suffix="."/></text:list-style>`;
    const frame = parseFrameChildForTest(
      `<draw:frame svg:x="1cm" svg:y="1cm" svg:width="10cm" svg:height="4cm"><draw:text-box>` +
        `<text:list text:style-name="Lnum" text:continue-numbering="true"><text:list-item><text:p>a</text:p></text:list-item>` +
        `<text:list-item text:start-value="4"><text:p>b</text:p></text:list-item></text:list>` +
        `</draw:text-box></draw:frame>`,
      styles,
    ) as TextFrame;
    expect(frame.paragraphs.map((p) => [p.bullet, p.startAt])).toEqual([
      ["number", undefined],
      ["number", 4],
    ]);
  });
});
