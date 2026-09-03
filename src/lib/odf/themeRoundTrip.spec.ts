// SPEC-077 — the .odt half of document themes (AC-37..AC-42): the theme
// block in styles.xml, the style:parent-style-name reference carry for
// untransformed colour refs, the two pinned .odt losses (a transformed
// colour ref and a theme font ref both re-open resolved), and the single
// "theme" drop record a save of a referencing document reports.
import { describe, it, expect } from "vitest";
import { parseOdt, type ParsedOdtDocument } from "./reader";
import { writeOdt, withShapeThemeParent } from "./writer";
import { unzip } from "../core/zip";
import { DEFAULT_THEME, resolveColor } from "../core/theme";
import type { DocumentTheme } from "../core/theme";
import { collectDropRecords, dropKindLabel } from "../opaque";
import {
  asParagraph,
  createDefaultStyles,
  type DocumentModel,
  type Paragraph,
  type Shape,
  type Table,
} from "../documentModel";

const THEME: DocumentTheme = {
  ...DEFAULT_THEME,
  name: "Round Trip Scheme",
  colors: { ...DEFAULT_THEME.colors, accent1: "#112233", accent3: "#334455", accent4: "#556677", accent5: "#778899" },
  fonts: { major: "Georgia", minor: "Verdana" },
};

const TRANSFORMED = { theme: "accent1", lumMod: 0.6, lumOff: 0.4 } as const;

function paragraph(text: string): Paragraph {
  return { type: "paragraph", runs: [{ text, formatting: {} }] };
}

/** A fresh (no preserved package) model exercising every AC-38 surface plus
 * the AC-39 losses: an untransformed run colour ref, a transformed ref, a
 * theme font ref, a literal, a paragraph-shading ref, a cell-shading ref
 * beside a literal cell, and a solid-fill shape ref. */
function themedModel(): DocumentModel {
  const table: Table = {
    type: "table",
    rows: [
      {
        cells: [
          { blocks: [paragraph("themed")], shadingFill: { theme: "accent4" } },
          { blocks: [paragraph("literal")], shadingFill: "#abcdef" },
        ],
      },
    ],
    columnWidthsCm: [5, 5],
  };
  const shape: Shape = {
    type: "shape",
    geometry: "rect",
    widthCm: 5,
    heightCm: 3,
    wrap: "inline",
    fill: { kind: "solid", color: { theme: "accent5" } },
  };
  const shaded: Paragraph = {
    type: "paragraph",
    runs: [{ text: "shaded", formatting: {} }],
    formatting: { shadingFill: { theme: "accent3" } },
  };
  return {
    id: "theme-doc",
    title: "Theme Doc",
    defaultStyles: createDefaultStyles(),
    theme: structuredClone(THEME),
    sections: [
      {
        blocks: [
          {
            type: "paragraph",
            runs: [
              { text: "ref ", formatting: { color: { theme: "accent1" } } },
              { text: "tinted ", formatting: { color: { ...TRANSFORMED } } },
              { text: "heading-font ", formatting: { fontFamily: { theme: "major" } } },
              { text: "literal", formatting: { color: "#123456" } },
            ],
          },
          shaded,
          table,
          shape,
        ],
      },
    ],
  };
}

async function saveAndReopen(model: DocumentModel): Promise<{ parsed: ParsedOdtDocument; parts: Map<string, Uint8Array> }> {
  const { bytes } = await writeOdt(model);
  return { parsed: await parseOdt(bytes), parts: await unzip(bytes) };
}

const decoder = new TextDecoder();

describe("SPEC-077 — .odt theme round trip", () => {
  it("stores the theme as de-theme-* styles and round-trips model.theme (AC-37)", async () => {
    const { parsed, parts } = await saveAndReopen(themedModel());
    const stylesXml = decoder.decode(parts.get("styles.xml"));
    expect(stylesXml).toContain('style:name="de-theme" style:family="graphic" style:display-name="Round Trip Scheme"');
    expect(stylesXml).toContain('style:name="de-theme-accent1"');
    expect(stylesXml).toContain('style:name="de-theme-paragraph-accent1"');
    expect(stylesXml).toContain('style:name="de-theme-cell-accent1"');
    expect(parsed.theme?.name).toBe(THEME.name);
    expect(parsed.theme?.colors).toEqual(THEME.colors);
    expect(parsed.theme?.fonts).toEqual(THEME.fonts);
    // The carrier styles are the theme's wire form, not catalogue entries.
    expect(parsed.styles.some((style) => style.id.startsWith("de-theme"))).toBe(false);
  });

  it("carries an untransformed run colour ref via style:parent-style-name beside the literal (AC-38)", async () => {
    const model = themedModel();
    const { parsed, parts } = await saveAndReopen(model);
    const contentXml = decoder.decode(parts.get("content.xml"));
    // The automatic style parents the carrier AND declares the resolved
    // literal, so LibreOffice shows the same colour.
    expect(contentXml).toContain('style:parent-style-name="de-theme-text-accent1"');
    expect(contentXml).toContain('fo:color="#112233"');
    const runs = asParagraph(parsed.sections[0].blocks[0]).runs;
    expect(runs[0].formatting.color).toEqual({ theme: "accent1" });
  });

  it("re-opens a transformed colour ref as its resolved literal and a theme font as its family — the two pinned losses (AC-39)", async () => {
    const { parsed } = await saveAndReopen(themedModel());
    const runs = asParagraph(parsed.sections[0].blocks[0]).runs;
    expect(runs[1].formatting.color).toBe(resolveColor({ ...TRANSFORMED }, THEME));
    expect(runs[2].formatting.fontFamily).toBe("Georgia");
    // A literal is untouched by any of this.
    expect(runs[3].formatting.color).toBe("#123456");
  });

  it("carries paragraph-shading and cell-shading refs and leaves literal shading alone (AC-38, AC-40)", async () => {
    const { parsed, parts } = await saveAndReopen(themedModel());
    const contentXml = decoder.decode(parts.get("content.xml"));
    expect(contentXml).toContain('style:parent-style-name="de-theme-paragraph-accent3"');
    expect(contentXml).toContain('style:parent-style-name="de-theme-cell-accent4"');
    expect(asParagraph(parsed.sections[0].blocks[1]).formatting?.shadingFill).toEqual({ theme: "accent3" });
    const table = parsed.sections[0].blocks[2];
    if (table.type !== "table") {
      throw new Error("expected the table block");
    }
    expect(table.rows[0].cells[0].shadingFill).toEqual({ theme: "accent4" });
    expect(table.rows[0].cells[1].shadingFill).toBe("#abcdef");
  });

  it("carries a shape's solid-fill ref on the graphic style's parent (AC-38)", async () => {
    const { parsed, parts } = await saveAndReopen(themedModel());
    const contentXml = decoder.decode(parts.get("content.xml"));
    expect(contentXml).toContain('style:parent-style-name="de-theme-accent5"');
    const shape = parsed.sections[0].blocks[3];
    if (shape.type !== "shape") {
      throw new Error("expected the shape block");
    }
    expect(shape.fill).toEqual({ kind: "solid", color: { theme: "accent5" } });
  });

  it("emits the de-theme-* carrier block for a THEMELESS document that carries a reference, so no written parent dangles (AC-38, AC-48)", async () => {
    // A blank document never sets model.theme (AC-48), yet a shape preset or
    // a theme swatch stores a reference — the carrier styles the automatic
    // styles parent must still exist, resolved against DEFAULT_THEME.
    const themeless: DocumentModel = {
      id: "themeless-ref",
      title: "Themeless Ref",
      defaultStyles: createDefaultStyles(),
      sections: [
        {
          blocks: [
            { type: "paragraph", runs: [{ text: "accented", formatting: { color: { theme: "accent1" } } }] },
            {
              type: "shape",
              geometry: "rect",
              widthCm: 5,
              heightCm: 3,
              wrap: "inline",
              fill: { kind: "solid", color: { theme: "accent2" } },
            },
          ],
        },
      ],
    };
    const { parsed, parts } = await saveAndReopen(themeless);
    const stylesXml = decoder.decode(parts.get("styles.xml"));
    const contentXml = decoder.decode(parts.get("content.xml"));
    // Every de-theme-* parent written into content.xml resolves to a style
    // styles.xml defines.
    for (const match of contentXml.matchAll(/style:parent-style-name="(de-theme[^"]*)"/g)) {
      expect(stylesXml).toContain(`style:name="${match[1]}"`);
    }
    expect(contentXml).toContain('style:parent-style-name="de-theme-text-accent1"');
    expect(contentXml).toContain('style:parent-style-name="de-theme-accent2"');
    // The references survive the round trip…
    const runs = asParagraph(parsed.sections[0].blocks[0]).runs;
    expect(runs[0].formatting.color).toEqual({ theme: "accent1" });
    const shape = parsed.sections[0].blocks[1];
    if (shape.type !== "shape") {
      throw new Error("expected the shape block");
    }
    expect(shape.fill).toEqual({ kind: "solid", color: { theme: "accent2" } });
    // …resolved against the product default scheme.
    expect(parsed.theme?.colors.accent1).toBe(DEFAULT_THEME.colors.accent1);
  });

  it("writes the literal alone when the parent slot is already taken by a link it did not mint (AC-38 conflict rule)", () => {
    const shape: Shape = {
      type: "shape",
      geometry: "rect",
      widthCm: 5,
      heightCm: 3,
      wrap: "inline",
      fill: { kind: "solid", color: { theme: "accent5" } },
    };
    // The writer-minted boilerplate parent yields to the reference…
    expect(
      withShapeThemeParent('<style:style style:name="shfr1" style:family="graphic" style:parent-style-name="Graphics"/>', shape),
    ).toContain('style:parent-style-name="de-theme-accent5"');
    // …but a data-bearing parent keeps its slot: the resolved literal in
    // the properties then stands alone (stated, not silent — see the
    // "theme" drop record test below).
    const foreign = '<style:style style:name="shfr1" style:family="graphic" style:parent-style-name="MyShapeBase"/>';
    expect(withShapeThemeParent(foreign, shape)).toBe(foreign);
    // A transformed ref never claims the parent at all.
    const boilerplate = '<style:style style:name="shfr1" style:family="graphic" style:parent-style-name="Graphics"/>';
    expect(
      withShapeThemeParent(boilerplate, { ...shape, fill: { kind: "solid", color: { ...TRANSFORMED } } }),
    ).toBe(boilerplate);
  });

  it("splices the theme block into a preserved styles.xml and keeps an unthemed untouched save byte-identical (AC-37)", async () => {
    // Unthemed: a save → reopen → save must keep styles.xml byte-identical.
    const unthemed: DocumentModel = {
      id: "plain",
      title: "Plain",
      defaultStyles: createDefaultStyles(),
      sections: [{ blocks: [paragraph("hello")] }],
    };
    const first = await writeOdt(unthemed);
    const parsedPlain = await parseOdt(first.bytes);
    const reopenedPlain: DocumentModel = {
      ...unthemed,
      sections: parsedPlain.sections,
      defaultStyles: parsedPlain.defaultStyles,
      styles: parsedPlain.styles,
      format: "odt",
      opaque: parsedPlain.opaque,
      sourceXml: parsedPlain.sourceXml,
    };
    const second = await writeOdt(reopenedPlain);
    const firstStyles = decoder.decode((await unzip(first.bytes)).get("styles.xml"));
    const secondStyles = decoder.decode((await unzip(second.bytes)).get("styles.xml"));
    expect(secondStyles).toBe(firstStyles);

    // Themed: a preserved save carries the block exactly once, updated to
    // the model's scheme.
    const themed = themedModel();
    const { bytes } = await writeOdt(themed);
    const parsed = await parseOdt(bytes);
    const reopened: DocumentModel = {
      ...themed,
      sections: parsed.sections,
      defaultStyles: parsed.defaultStyles,
      styles: parsed.styles,
      theme: parsed.theme,
      format: "odt",
      opaque: parsed.opaque,
      sourceXml: parsed.sourceXml,
    };
    const resaved = await writeOdt(reopened);
    const resavedStyles = decoder.decode((await unzip(resaved.bytes)).get("styles.xml"));
    expect((resavedStyles.match(/style:name="de-theme"/g) ?? []).length).toBe(1);
    expect((resavedStyles.match(/style:name="de-theme-accent1"/g) ?? []).length).toBe(1);
    expect(resavedStyles).toContain('style:display-name="Round Trip Scheme"');
  });

  it("reports one 'theme' drop record naming the theme for a referencing save, none otherwise (AC-41, AC-42)", async () => {
    // The dialog renders kinds generically via dropKindLabel (AC-41).
    expect(dropKindLabel("theme")).toBe("Document theme");

    const { dropped } = await writeOdt(themedModel());
    const themeDrops = dropped.filter((record) => record.kind === "theme");
    expect(themeDrops).toHaveLength(1);
    expect(themeDrops[0].location).toContain('"Round Trip Scheme"');
    expect(themeDrops[0].reason).toContain("no theme part");

    // A same-format .odt save of a referencing document reports it too.
    const { bytes } = await writeOdt(themedModel());
    const parsed = await parseOdt(bytes);
    const reopened: DocumentModel = {
      id: "again",
      title: "Again",
      sections: parsed.sections,
      defaultStyles: parsed.defaultStyles,
      theme: parsed.theme,
      format: "odt",
      opaque: parsed.opaque,
      sourceXml: parsed.sourceXml,
    };
    const resaved = await writeOdt(reopened);
    expect(resaved.dropped.filter((record) => record.kind === "theme")).toHaveLength(1);

    // A theme with no reference anywhere loses nothing.
    const literalOnly: DocumentModel = {
      id: "literal",
      title: "Literal",
      defaultStyles: createDefaultStyles(),
      theme: structuredClone(THEME),
      sections: [{ blocks: [{ type: "paragraph", runs: [{ text: "plain", formatting: { color: "#112233" } }] }] }],
    };
    const literalSave = await writeOdt(literalOnly);
    expect(literalSave.dropped.filter((record) => record.kind === "theme")).toHaveLength(0);

    // A Markdown save reports no new kind for the same referencing model.
    expect(collectDropRecords(themedModel(), "md").some((record) => record.kind === "theme")).toBe(false);
  });
});
