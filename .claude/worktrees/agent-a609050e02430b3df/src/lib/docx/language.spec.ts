// SPEC-035 — document/run language (w:lang) and "do not check spelling"
// (w:noProof) through the .docx reader, writer and validator.
import { describe, it, expect } from "vitest";
import { asParagraph, createDocument, type DocumentModel, type RunFormatting } from "../documentModel";
import { parseRunProperties } from "./properties";
import { parseDocx } from "./reader";
import { buildDocumentXml, buildStylesXml, writeDocx } from "./writer";
import { validateDocx } from "./validate";
import { WORD_NAMESPACE } from "./xml";
import { parseXml } from "../core/xml";

function rPr(inner: string): Element {
  return parseXml(`<w:rPr xmlns:w="${WORD_NAMESPACE}">${inner}</w:rPr>`).documentElement;
}

function docWithRun(formatting: RunFormatting, text = "hello"): DocumentModel {
  const model = createDocument("docx-language-test");
  model.sections = [{ blocks: [{ type: "paragraph", runs: [{ text, formatting }] }] }];
  return model;
}

describe("parseRunProperties — lang, noProof (SPEC-035)", () => {
  it("reads w:lang/@w:val", () => {
    expect(parseRunProperties(rPr('<w:lang w:val="fr-FR"/>')).lang).toBe("fr-FR");
  });

  it("reads w:noProof as a boolean toggle, present-with-no-val means true", () => {
    expect(parseRunProperties(rPr("<w:noProof/>")).noProof).toBe(true);
    expect(parseRunProperties(rPr('<w:noProof w:val="0"/>')).noProof).toBe(false);
  });

  it("leaves lang/noProof unset when absent", () => {
    const formatting = parseRunProperties(rPr("<w:b/>"));
    expect(formatting.lang).toBeUndefined();
    expect(formatting.noProof).toBeUndefined();
  });
});

describe("buildDocumentXml — lang mapping on save (SPEC-035)", () => {
  it("writes w:lang after w:vertAlign, per the CT_RPr sequence", () => {
    const xml = buildDocumentXml(docWithRun({ vertAlign: "subscript", lang: "de-DE" }));
    const vertAlign = xml.indexOf("<w:vertAlign");
    const lang = xml.indexOf("<w:lang");
    expect(vertAlign).toBeGreaterThan(-1);
    expect(lang).toBeGreaterThan(vertAlign);
    expect(xml).toContain('<w:lang w:val="de-DE"/>');
  });

  it("writes w:noProof before w:color, per the CT_RPr sequence", () => {
    const xml = buildDocumentXml(docWithRun({ noProof: true, color: "#112233" }));
    const noProof = xml.indexOf("<w:noProof");
    const color = xml.indexOf("<w:color");
    expect(noProof).toBeGreaterThan(-1);
    expect(color).toBeGreaterThan(noProof);
  });

  it("omits both when unset", () => {
    const xml = buildDocumentXml(docWithRun({ bold: true }));
    expect(xml).not.toContain("<w:lang");
    expect(xml).not.toContain("<w:noProof");
  });
});

describe("buildStylesXml — document default language (SPEC-035)", () => {
  it("writes the document default into docDefaults/rPrDefault/rPr/w:lang", () => {
    const model = createDocument("docx-language-doc-default");
    model.defaultStyles.language = "cs-CZ";
    const xml = buildStylesXml(model);
    expect(xml).toContain("<w:docDefaults>");
    const docDefaultsIndex = xml.indexOf("<w:docDefaults>");
    const rPrDefaultEnd = xml.indexOf("</w:rPrDefault>");
    const lang = xml.indexOf('<w:lang w:val="cs-CZ"/>');
    expect(lang).toBeGreaterThan(docDefaultsIndex);
    expect(lang).toBeLessThan(rPrDefaultEnd);
  });
});

describe("writeDocx — validates and round-trips (SPEC-035)", () => {
  it("produces a schema-valid package with lang/noProof on runs and docDefaults", async () => {
    const model = docWithRun({ lang: "fr-FR", noProof: false });
    asParagraph(model.sections[0].blocks[0]).runs.push({ text: " skip", formatting: { noProof: true } });
    const { bytes } = await writeDocx(model);
    expect(await validateDocx(bytes)).toEqual([]);
  });

  it("round-trips a run's explicit language and noProof exactly", async () => {
    const { bytes } = await writeDocx(docWithRun({ lang: "es-ES", noProof: true }));
    const reopened = await parseDocx(bytes);
    const run = asParagraph(reopened.sections[0].blocks[0]).runs[0];
    expect(run.formatting.lang).toBe("es-ES");
    expect(run.formatting.noProof).toBe(true);
  });

  it("round-trips the document default language onto defaultStyles.language", async () => {
    const model = createDocument("docx-language-roundtrip-default");
    model.defaultStyles.language = "de-DE";
    const { bytes } = await writeDocx(model);
    const reopened = await parseDocx(bytes);
    expect(reopened.defaultStyles.language).toBe("de-DE");
  });

  it("does NOT stamp an explicit lang onto a run that never diverged from the document default (AC)", async () => {
    const model = docWithRun({ bold: true });
    model.defaultStyles.language = "de-DE";
    const { bytes } = await writeDocx(model);
    const reopened = await parseDocx(bytes);
    const run = asParagraph(reopened.sections[0].blocks[0]).runs[0];
    expect(run.formatting.lang).toBeUndefined();
    expect(reopened.defaultStyles.language).toBe("de-DE");
  });

  it("a run's own language overrides the document default independently (AC)", async () => {
    const model = docWithRun({ lang: "fr-FR" });
    model.defaultStyles.language = "de-DE";
    const { bytes } = await writeDocx(model);
    const reopened = await parseDocx(bytes);
    expect(reopened.defaultStyles.language).toBe("de-DE");
    expect(asParagraph(reopened.sections[0].blocks[0]).runs[0].formatting.lang).toBe("fr-FR");
  });
});
