// SPEC-035 — document/run language (fo:language/fo:country) and "do not
// check spelling" (the fo:language="zxx" convention) through the .odt
// reader and writer.
import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { asParagraph, createDocument, type DocumentModel, type RunFormatting } from "../documentModel";
import { parseTextProperties } from "./properties";
import { parseOdt } from "./reader";
import { buildContentXml, buildStylesXml, writeOdt } from "./writer";
import { unzip } from "../core/zip";
import { equalBytes } from "../core/equality";
import { parseXml } from "../core/xml";
import { FO_NAMESPACE, STYLE_NAMESPACE } from "../core/odfXml";

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");

function loadFixture(name: string): Uint8Array {
  return new Uint8Array(readFileSync(join(fixturesDir, name)));
}

function textProperties(attrs: string): Element {
  return parseXml(
    `<style:text-properties xmlns:style="${STYLE_NAMESPACE}" xmlns:fo="${FO_NAMESPACE}" ${attrs}/>`,
  ).documentElement;
}

function docWithRun(formatting: RunFormatting, text = "hello"): DocumentModel {
  const model = createDocument("odf-language-test");
  model.sections = [{ blocks: [{ type: "paragraph", runs: [{ text, formatting }] }] }];
  return model;
}

describe("parseTextProperties — language, noProof (SPEC-035)", () => {
  const noFonts = new Map<string, string>();

  it("composes fo:language + fo:country into one tag", () => {
    expect(
      parseTextProperties(textProperties('fo:language="fr" fo:country="FR"'), noFonts).formatting.lang,
    ).toBe("fr-FR");
  });

  it("keeps a language-only tag with no country", () => {
    expect(parseTextProperties(textProperties('fo:language="de"'), noFonts).formatting.lang).toBe("de");
  });

  it('maps fo:language="zxx" to noProof, not a language', () => {
    const formatting = parseTextProperties(textProperties('fo:language="zxx"'), noFonts).formatting;
    expect(formatting.noProof).toBe(true);
    expect(formatting.lang).toBeUndefined();
  });
});

describe("buildContentXml — language mapping on save (SPEC-035)", () => {
  it("writes fo:language + fo:country", () => {
    const xml = buildContentXml(docWithRun({ lang: "es-ES" }));
    expect(xml).toContain('fo:language="es"');
    expect(xml).toContain('fo:country="ES"');
  });

  it('writes noProof as fo:language="zxx", taking precedence over lang', () => {
    const xml = buildContentXml(docWithRun({ lang: "es-ES", noProof: true }));
    expect(xml).toContain('fo:language="zxx"');
    expect(xml).not.toContain('fo:language="es"');
  });
});

describe("buildStylesXml — document default language (SPEC-035)", () => {
  it("writes the model's default language into style:default-style", () => {
    const model = createDocument("odf-language-doc-default");
    model.defaultStyles.language = "cs-CZ";
    const xml = buildStylesXml(model);
    expect(xml).toContain('style:family="paragraph"');
    expect(xml).toContain('fo:language="cs"');
    expect(xml).toContain('fo:country="CZ"');
  });
});

describe("writeOdt/parseOdt — round-trips (SPEC-035)", () => {
  it("round-trips a run's explicit language and noProof exactly", async () => {
    const { bytes } = await writeOdt(docWithRun({ lang: "fr-FR", noProof: false }));
    const reopened = await parseOdt(bytes);
    const run = asParagraph(reopened.sections[0].blocks[0]).runs[0];
    expect(run.formatting.lang).toBe("fr-FR");
  });

  it("round-trips noProof (fo:language=zxx) without inventing a language", async () => {
    const { bytes } = await writeOdt(docWithRun({ noProof: true }));
    const reopened = await parseOdt(bytes);
    const run = asParagraph(reopened.sections[0].blocks[0]).runs[0];
    expect(run.formatting.noProof).toBe(true);
    expect(run.formatting.lang).toBeUndefined();
  });

  it("round-trips the document default language onto defaultStyles.language", async () => {
    const model = createDocument("odf-language-roundtrip-default");
    model.defaultStyles.language = "de-DE";
    const { bytes } = await writeOdt(model);
    const reopened = await parseOdt(bytes);
    expect(reopened.defaultStyles.language).toBe("de-DE");
  });

  it("does NOT stamp an explicit lang onto a run that never diverged from the document default (AC)", async () => {
    const model = docWithRun({ bold: true });
    model.defaultStyles.language = "de-DE";
    const { bytes } = await writeOdt(model);
    const reopened = await parseOdt(bytes);
    const run = asParagraph(reopened.sections[0].blocks[0]).runs[0];
    expect(run.formatting.lang).toBeUndefined();
    expect(reopened.defaultStyles.language).toBe("de-DE");
  });
});

function modelFrom(opened: Awaited<ReturnType<typeof parseOdt>>, title = "Field Report"): DocumentModel {
  const model: DocumentModel = {
    id: "test-doc",
    title,
    sections: opened.sections,
    defaultStyles: opened.defaultStyles,
    format: "odt",
    opaque: opened.opaque,
    sourceXml: opened.sourceXml,
  };
  return model;
}

describe("writeOdt — preserved styles.xml (SPEC-008 byte-stability, SPEC-035)", () => {
  it("leaves styles.xml byte-identical when a re-opened fixture is saved unmodified", async () => {
    const fixture = loadFixture("libreoffice-writer.odt");
    const originalParts = await unzip(fixture);
    const opened = await parseOdt(fixture);
    const { bytes } = await writeOdt(modelFrom(opened), { modified: new Date("2026-08-03T12:34:56Z") });
    const savedParts = await unzip(bytes);
    expect(equalBytes(savedParts.get("styles.xml")!, originalParts.get("styles.xml")!)).toBe(true);
  });

  it("patches only the default-style language when the document language changes", async () => {
    const fixture = loadFixture("libreoffice-writer.odt");
    const opened = await parseOdt(fixture);
    const model = modelFrom(opened);
    model.defaultStyles = { ...model.defaultStyles, language: "cs-CZ" };
    const { bytes } = await writeOdt(model, { modified: new Date("2026-08-03T12:34:56Z") });
    const reopened = await parseOdt(bytes);
    expect(reopened.defaultStyles.language).toBe("cs-CZ");
  });
});
