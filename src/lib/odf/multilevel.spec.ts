// SPEC-032 — multilevel definitions through .odt (AC-27, AC-28). The
// previous attempt failed AC-28 here: the writer grouped every contiguous
// run of list paragraphs into ONE text:list keyed by the first paragraph's
// instance, so a mid-run restart re-imported as a continuation. These tests
// pin the fix: per-instance text:list segments, continue-list across
// interruptions, and the full write -> re-open fixed point for the QA
// scenario (two independent lists plus a restarted one).
import { describe, expect, it } from "vitest";
import {
  applyMultilevelPreset,
  applyListDefinition,
  caretSelection,
  continueListNumbering,
  restartListNumbering,
  setListNumberingValue,
  type ModelSelection,
} from "../commands";
import { asParagraph, createDocument, type DocumentModel, type Paragraph } from "../documentModel";
import { listMarkers } from "../lists";
import { buildContentXml, writeOdt } from "./writer";
import { parseOdt } from "./reader";

function paragraph(text: string): Paragraph {
  return { type: "paragraph", runs: [{ text, formatting: {} }] };
}

function caret(block: number): ModelSelection {
  return caretSelection({ section: 0, block, offset: 0 });
}

function range(anchorBlock: number, focusBlock: number): ModelSelection {
  return {
    anchor: { section: 0, block: anchorBlock, offset: 0 },
    focus: { section: 0, block: focusBlock, offset: 0 },
  };
}

function docWith(...blocks: Paragraph[]): DocumentModel {
  const model = createDocument("odf-multilevel-fixture");
  model.sections = [{ blocks }];
  return model;
}

/** The QA scenario (AC-28): list A (3 items), list B (2 items), list C
 * (2 items, second restarted), separated by plain paragraphs. */
function qaScenario(): DocumentModel {
  let model = docWith(
    paragraph("a1"),
    paragraph("a2"),
    paragraph("a3"),
    paragraph("gap1"),
    paragraph("b1"),
    paragraph("b2"),
    paragraph("gap2"),
    paragraph("c1"),
    paragraph("c2"),
  );
  model = applyMultilevelPreset(model, range(0, 2), "decimal").model;
  model = applyMultilevelPreset(model, range(4, 5), "decimal").model;
  model = applyMultilevelPreset(model, range(7, 8), "decimal").model;
  model = restartListNumbering(model, caret(8)).model;
  return model;
}

async function reopen(model: DocumentModel): Promise<DocumentModel> {
  const { bytes } = await writeOdt(model);
  const parsed = await parseOdt(bytes);
  const reopened = createDocument("reopened");
  reopened.sections = parsed.sections;
  if (parsed.listDefinitions.length > 0) {
    reopened.listDefinitions = parsed.listDefinitions;
  }
  return reopened;
}

describe("content.xml emission (AC-27)", () => {
  it("writes definitions as text:list-level-style-number with display-levels", () => {
    let model = docWith(paragraph("one"), paragraph("sub"));
    model = applyMultilevelPreset(model, range(0, 1), "decimal").model;
    const second = asParagraph(model.sections[0].blocks[1]);
    second.list = { ...second.list!, level: 2 };
    const xml = buildContentXml(model);
    expect(xml).toContain('<text:list-level-style-number text:level="1"');
    expect(xml).toContain('text:display-levels="1"');
    expect(xml).toContain('text:display-levels="2"');
    expect(xml).toContain('style:num-suffix="."');
  });

  it("a restarted instance becomes a SEPARATE text:list, not one merged run", () => {
    const model = qaScenario();
    const xml = buildContentXml(model);
    // c1 and c2 are adjacent but in different instances: two sibling
    // text:list elements, the second WITHOUT continue-numbering.
    const c1 = xml.indexOf(">c1<");
    const c2 = xml.indexOf(">c2<");
    const between = xml.slice(c1, c2);
    expect(between).toContain("</text:list>");
    expect(between).toContain("<text:list ");
    expect(between).not.toContain("text:continue-numbering");
  });

  it("an interrupted instance's later segment continues via text:continue-list", () => {
    let model = docWith(paragraph("one"), paragraph("two"), paragraph("gap"), paragraph("three"));
    model = applyMultilevelPreset(model, range(0, 1), "decimal").model;
    model = applyMultilevelPreset(model, caret(3), "decimal").model;
    model = continueListNumbering(model, caret(3)).model;
    expect(listMarkers(model)[0]).toEqual(["1.", "2.", "", "3."]);
    const xml = buildContentXml(model);
    expect(xml).toMatch(/text:continue-list="list\d+" text:continue-numbering="true"/);
  });

  it("a set-value restart writes text:start-value on its list item", () => {
    let model = docWith(paragraph("one"), paragraph("two"));
    model = applyMultilevelPreset(model, range(0, 1), "decimal").model;
    model = setListNumberingValue(model, caret(1), 42).model;
    const xml = buildContentXml(model);
    expect(xml).toContain('<text:list-item text:start-value="42">');
  });
});

describe("full .odt round trip (AC-27, AC-28)", () => {
  it("preserves all three numbering sequences of the QA scenario", async () => {
    const model = qaScenario();
    const expected = ["1.", "2.", "3.", "", "1.", "2.", "", "1.", "1."];
    expect(listMarkers(model)[0]).toEqual(expected);
    const reopened = await reopen(model);
    expect(listMarkers(reopened)[0]).toEqual(expected);
    // …and a SECOND cycle stays stable (the reopened model reproduces the
    // same instances on the next save).
    const twice = await reopen(reopened);
    expect(listMarkers(twice)[0]).toEqual(expected);
  });

  it("preserves a continued (interrupted) instance", async () => {
    let model = docWith(paragraph("one"), paragraph("two"), paragraph("gap"), paragraph("three"));
    model = applyMultilevelPreset(model, range(0, 1), "decimal").model;
    model = applyMultilevelPreset(model, caret(3), "decimal").model;
    model = continueListNumbering(model, caret(3)).model;
    const reopened = await reopen(model);
    expect(listMarkers(reopened)[0]).toEqual(["1.", "2.", "", "3."]);
  });

  it("preserves a set-value restart", async () => {
    let model = docWith(paragraph("one"), paragraph("two"));
    model = applyMultilevelPreset(model, range(0, 1), "decimal").model;
    model = setListNumberingValue(model, caret(1), 42).model;
    const reopened = await reopen(model);
    expect(listMarkers(reopened)[0]).toEqual(["1.", "42."]);
  });

  it("preserves a customised definition's per-level format", async () => {
    let model = docWith(paragraph("one"), paragraph("sub"));
    model = applyMultilevelPreset(model, range(0, 1), "decimal").model;
    const second = asParagraph(model.sections[0].blocks[1]);
    second.list = { ...second.list!, level: 2 };
    model = applyListDefinition(model, caret(0), [
      { style: "upperRoman", prefix: "(", suffix: ")" },
      { style: "lowerAlpha", suffix: ".", includeLevels: 2, startAt: 3 },
    ]).model;
    const before = listMarkers(model)[0];
    expect(before).toEqual(["(I)", "I.c."]);
    const reopened = await reopen(model);
    expect(listMarkers(reopened)[0]).toEqual(before);
  });

  it("keeps a mixed bullet/number definition defined across the round trip", async () => {
    let model = docWith(paragraph("one"), paragraph("sub"));
    model = applyMultilevelPreset(model, range(0, 1), "bulletMixed").model;
    const second = asParagraph(model.sections[0].blocks[1]);
    second.list = { ...second.list!, type: "bullet", level: 2 };
    expect(listMarkers(model)[0]).toEqual(["1.", "•"]);
    const reopened = await reopen(model);
    expect(reopened.listDefinitions).toHaveLength(1);
    expect(listMarkers(reopened)[0]).toEqual(["1.", "•"]);
  });

  it("leaves plain SPEC-018 adjacency lists exactly as before", async () => {
    const model = docWith(paragraph("x"), paragraph("y"), paragraph("gap"), paragraph("z"));
    const one = asParagraph(model.sections[0].blocks[0]);
    const two = asParagraph(model.sections[0].blocks[1]);
    const three = asParagraph(model.sections[0].blocks[3]);
    one.list = { type: "number", level: 1 };
    two.list = { type: "number", level: 1 };
    three.list = { type: "number", level: 1 };
    const reopened = await reopen(model);
    expect(reopened.listDefinitions).toBeUndefined();
    expect(listMarkers(reopened)[0]).toEqual(["1.", "2.", "", "1."]);
    expect(asParagraph(reopened.sections[0].blocks[0]).list).toEqual({ type: "number", level: 1 });
  });

  it("round-trips custom tab stops through style:tab-stops (AC-19..23)", async () => {
    const model = docWith(paragraph("tabbed"));
    asParagraph(model.sections[0].blocks[0]).formatting = {
      tabStops: [
        { positionCm: 3, alignment: "left" },
        { positionCm: 4, alignment: "right", leader: "dot" },
        { positionCm: 5, alignment: "center", leader: "hyphen" },
        { positionCm: 6, alignment: "decimal", leader: "underscore" },
      ],
    };
    const { bytes } = await writeOdt(model);
    const parsed = await parseOdt(bytes);
    const stops = asParagraph(parsed.sections[0].blocks[0]).formatting?.tabStops;
    expect(stops).toEqual([
      { positionCm: 3, alignment: "left" },
      { positionCm: 4, alignment: "right", leader: "dot" },
      { positionCm: 5, alignment: "center", leader: "hyphen" },
      { positionCm: 6, alignment: "decimal", leader: "underscore" },
    ]);
  });

  it("drops each bar tab stop with a DropRecord and keeps the rest (SPEC-054 AC-38)", async () => {
    // A bar stop only reaches an .odt save from a .docx-sourced document
    // (ODF has no bar tab, so the .odt reader never produces one); the drop
    // report fires on that cross-format save.
    const model = docWith(paragraph("barred"));
    model.format = "docx";
    asParagraph(model.sections[0].blocks[0]).formatting = {
      tabStops: [
        { positionCm: 3, alignment: "bar" },
        { positionCm: 5, alignment: "left" },
      ],
    };
    const { bytes, dropped } = await writeOdt(model);
    const barDrop = dropped.find((record) => record.kind === "tabStop");
    expect(barDrop).toBeDefined();
    expect(barDrop?.reason).toContain("no bar tab stop");
    // The non-bar stop survives; the bar stop is gone (ODF has no equivalent).
    const parsed = await parseOdt(bytes);
    const stops = asParagraph(parsed.sections[0].blocks[0]).formatting?.tabStops;
    expect(stops).toEqual([{ positionCm: 5, alignment: "left" }]);
  });

  it("round-trips the document default tab interval as style:tab-stop-distance (SPEC-054 AC-37)", async () => {
    const model = docWith(paragraph("spaced"));
    model.defaultStyles.defaultTabStopCm = 1.5;
    const { bytes } = await writeOdt(model);
    const reopened = await parseOdt(bytes);
    expect(reopened.defaultStyles.defaultTabStopCm).toBeCloseTo(1.5, 4);
  });
});
