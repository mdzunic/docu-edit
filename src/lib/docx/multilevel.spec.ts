// SPEC-032 — multilevel definitions through .docx (AC-26, AC-28): distinct
// w:abstractNum entries with composed w:lvlText and per-level w:numFmt, one
// w:num per instance, startOverride restarts, and the full write -> re-open
// fixed point including the QA scenario the previous attempt failed (two
// independent lists plus a restarted one).
import { describe, expect, it } from "vitest";
import {
  applyMultilevelPreset,
  applyListDefinition,
  caretSelection,
  restartListNumbering,
  setListNumberingValue,
  type ModelSelection,
} from "../commands";
import { asParagraph, createDocument, type DocumentModel, type Paragraph } from "../documentModel";
import { listMarkers } from "../lists";
import { parseXml } from "../core/xml";
import {
  buildNumberingXml,
  parseNumberingCatalog,
  planNumbering,
  levelTemplate,
  parseLevelTemplate,
} from "./numbering";
import { parseDocx } from "./reader";
import { writeDocx } from "./writer";

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
  const model = createDocument("docx-multilevel-fixture");
  model.sections = [{ blocks }];
  return model;
}

/** The QA scenario (AC-28): list A (3 items), list B (2 items), list C
 * (2 items with the second explicitly restarted), separated by plain
 * paragraphs. Expected markers: 1./2./3., 1./2., 1./1. */
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

describe("levelTemplate <-> parseLevelTemplate", () => {
  it("composes and parses back prefix/components/suffix", () => {
    const def: import("../core/lists").ListDefinition = {
      defId: "d",
      levels: [
        { style: "decimal" },
        { style: "decimal" },
        { style: "decimal", prefix: "(", suffix: ")", includeLevels: 3 },
      ],
    };
    const template = levelTemplate(def, 3);
    expect(template).toBe("(%1.%2.%3)");
    expect(parseLevelTemplate(template)).toEqual({ prefix: "(", suffix: ")", includeLevels: 3 });
  });
});

describe("planNumbering with definitions (AC-26)", () => {
  it("emits one multilevel abstractNum per definition and one num per instance", () => {
    const model = qaScenario();
    const plan = planNumbering(model);
    const xml = buildNumberingXml(plan);
    const doc = parseXml(xml);
    const abstracts = Array.from(doc.documentElement.children).filter(
      (child) => child.localName === "abstractNum",
    );
    const nums = Array.from(doc.documentElement.children).filter(
      (child) => child.localName === "num",
    );
    // Three definitions (one per list), four instances (C has two: the
    // original and the restarted one).
    expect(abstracts).toHaveLength(3);
    expect(nums).toHaveLength(4);
    for (const abstract of abstracts) {
      expect(abstract.outerHTML ?? "").toBeDefined();
      const lvlTexts = Array.from(abstract.getElementsByTagName("*")).filter(
        (el) => el.localName === "lvlText",
      );
      expect(lvlTexts).toHaveLength(9);
    }
    expect(xml).toContain('w:multiLevelType w:val="multilevel"');
    expect(xml).toContain('<w:lvlText w:val="%1."/>');
    expect(xml).toContain('<w:lvlText w:val="%1.%2."/>');
  });

  it("a set-value restart becomes a startOverride on its instance's num", () => {
    let model = docWith(paragraph("one"), paragraph("two"));
    model = applyMultilevelPreset(model, range(0, 1), "decimal").model;
    model = setListNumberingValue(model, caret(1), 42).model;
    const xml = buildNumberingXml(planNumbering(model));
    expect(xml).toContain('<w:lvlOverride w:ilvl="0"><w:startOverride w:val="42"/></w:lvlOverride>');
  });

  it("write -> parse catalog reproduces the definition (fixed point)", () => {
    let model = docWith(paragraph("one"));
    model = applyMultilevelPreset(model, caret(0), "decimal").model;
    model = applyListDefinition(model, caret(0), [
      { style: "upperRoman", prefix: "(", suffix: ")", startAt: 3, alignment: "right" },
      { style: "lowerAlpha", suffix: ".", includeLevels: 2 },
    ]).model;
    const xml = buildNumberingXml(planNumbering(model));
    const catalog = parseNumberingCatalog(parseXml(xml));
    const list = catalog.listFormattingFor("1", 0);
    expect(list?.defId).toBeDefined();
    const [def] = catalog.definitionsUsed();
    expect(def.levels[0]).toEqual({
      style: "upperRoman",
      prefix: "(",
      suffix: ")",
      startAt: 3,
      alignment: "right",
    });
    expect(def.levels[1]).toEqual({ style: "lowerAlpha", suffix: ".", includeLevels: 2 });
  });
});

describe("full .docx round trip (AC-26, AC-28)", () => {
  it("preserves all three numbering sequences of the QA scenario", async () => {
    const model = qaScenario();
    expect(listMarkers(model)[0]).toEqual(["1.", "2.", "3.", "", "1.", "2.", "", "1.", "1."]);
    const { bytes } = await writeDocx(model);
    const reopened = await parseDocx(bytes);
    const reopenedModel: DocumentModel = {
      ...createDocument("reopened"),
      sections: reopened.sections,
      listDefinitions: reopened.listDefinitions,
    };
    expect(listMarkers(reopenedModel)[0]).toEqual([
      "1.",
      "2.",
      "3.",
      "",
      "1.",
      "2.",
      "",
      "1.",
      "1.",
    ]);
  });

  it("keeps a set-value restart and heading-style links across the round trip", async () => {
    let model = docWith(paragraph("one"), paragraph("two"));
    model = applyMultilevelPreset(model, range(0, 1), "decimal").model;
    model = setListNumberingValue(model, caret(1), 42).model;
    model.listDefinitions![0].styleLinks = { Heading1: 1 };
    expect(listMarkers(model)[0]).toEqual(["1.", "42."]);
    const { bytes } = await writeDocx(model);
    const reopened = await parseDocx(bytes);
    const reopenedModel: DocumentModel = {
      ...createDocument("reopened"),
      sections: reopened.sections,
      listDefinitions: reopened.listDefinitions,
    };
    expect(listMarkers(reopenedModel)[0]).toEqual(["1.", "42."]);
    expect(reopened.listDefinitions[0]?.styleLinks).toEqual({ Heading1: 1 });
  });

  it("keeps a mixed bullet/number definition defined across the round trip", async () => {
    let model = docWith(paragraph("one"), paragraph("sub"));
    model = applyMultilevelPreset(model, range(0, 1), "bulletMixed").model;
    const second = asParagraph(model.sections[0].blocks[1]);
    second.list = { ...second.list!, type: "bullet", level: 2 };
    expect(listMarkers(model)[0]).toEqual(["1.", "•"]);
    const { bytes } = await writeDocx(model);
    const reopened = await parseDocx(bytes);
    expect(reopened.listDefinitions).toHaveLength(1);
    const reopenedModel: DocumentModel = {
      ...createDocument("reopened"),
      sections: reopened.sections,
      listDefinitions: reopened.listDefinitions,
    };
    expect(listMarkers(reopenedModel)[0]).toEqual(["1.", "•"]);
  });

  it("round-trips custom tab stops through w:tabs (AC-19..23)", async () => {
    const model = docWith(paragraph("tabbed"));
    asParagraph(model.sections[0].blocks[0]).formatting = {
      tabStops: [
        { positionCm: 2.99930555555556, alignment: "left" },
        { positionCm: 4.00050000000000, alignment: "right", leader: "dot" },
        { positionCm: 5.00028, alignment: "center", leader: "hyphen" },
        { positionCm: 6.00075, alignment: "decimal", leader: "underscore" },
      ],
    };
    const { bytes } = await writeDocx(model);
    const reopened = await parseDocx(bytes);
    const stops = asParagraph(reopened.sections[0].blocks[0]).formatting?.tabStops;
    expect(stops?.map((stop) => [stop.alignment, stop.leader])).toEqual([
      ["left", undefined],
      ["right", "dot"],
      ["center", "hyphen"],
      ["decimal", "underscore"],
    ]);
  });
});
