// Table look and style banding for .odt (SPEC-076 AC-68, AC-69, AC-70).
//
// ODF splits the same two halves .docx does, in its own vocabulary:
//   - the TABLE says which regions it uses, as six boolean attributes on
//     `table:table` — `table:use-first-row-styles`, `…-last-row-…`,
//     `…-first-column-…`, `…-last-column-…`, `…-banding-rows-…` and
//     `…-banding-columns-…`. Unlike OOXML's `w:noHBand`, these are positive,
//     so they map straight onto `TableLook`.
//   - the STYLE says what each region looks like, as a `table:table-template`
//     whose six children reference generated `table-cell` styles. The table
//     points at it with `table:template-name`.
//
// ODF's bands are odd/even only: there is no band-size attribute anywhere in
// the template, so a `bandedRows.size` other than 1 cannot cross (AC-70 — the
// writer emits size-1 banding and opaque.ts reports the loss).
import { childElements, escapeXml, firstChildElement } from "../core/xml";
import { borderShorthand, parseBorderShorthand } from "./borders";
import type { CellBorders } from "../core/tables";
import type { NamedStyle, TableConditionalFormat, TableLook, TableStyleBanding } from "../documentModel";

export const TABLE_NAMESPACE = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const STYLE_NAMESPACE = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const FO_NAMESPACE = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";

/** The six regions, each paired with its `table:use-*-styles` attribute, its
 * `table:table-template` child element and the `TableStyleBanding` field it
 * carries. One list drives both directions and both halves. */
const REGIONS: {
  look: keyof TableLook;
  use: string;
  child: string;
  band: keyof TableStyleBanding;
}[] = [
  { look: "firstRow", use: "use-first-row-styles", child: "first-row", band: "firstRow" },
  { look: "lastRow", use: "use-last-row-styles", child: "last-row", band: "lastRow" },
  { look: "firstColumn", use: "use-first-column-styles", child: "first-column", band: "firstColumn" },
  { look: "lastColumn", use: "use-last-column-styles", child: "last-column", band: "lastColumn" },
  { look: "bandedRows", use: "use-banding-rows-styles", child: "odd-rows", band: "bandedRows" },
  { look: "bandedColumns", use: "use-banding-columns-styles", child: "odd-columns", band: "bandedColumns" },
];

const BORDER_SIDES: { key: keyof CellBorders; fo: string }[] = [
  { key: "top", fo: "fo:border-top" },
  { key: "left", fo: "fo:border-left" },
  { key: "bottom", fo: "fo:border-bottom" },
  { key: "right", fo: "fo:border-right" },
];

/** The generated cell style one region of one template uses (AC-68). Derived
 * from the style id and the region's own element name, so the same style
 * always generates the same names and a same-format save is byte-stable. */
export function templateCellStyleName(styleId: string, child: string): string {
  return `${styleId}.${child}`;
}

/** Every cell-style name `tableTemplateXml` can generate for one style — the
 * set a patched styles.xml must clear before re-emitting the style, so a
 * previous save's artefacts (or a since-removed region's) never linger
 * beside the fresh ones (SPEC-076 AC-68). */
export function templateCellStyleNames(styleId: string): string[] {
  return REGIONS.map((region) => templateCellStyleName(styleId, region.child));
}

/** The `table:use-*-styles` attributes for a table's `look` (AC-68). An
 * absent look writes nothing at all, which is ODF's own default of "no
 * template regions in use". */
export function tableUseStyleAttrs(look: TableLook | undefined): string {
  if (!look) {
    return "";
  }
  return REGIONS.map((region) => ` table:${region.use}="${look[region.look] ? "true" : "false"}"`).join("");
}

/** A table's `look` from its `table:table` element (AC-69), or undefined
 * when it declares none of the six attributes — "no banding at all", not
 * "every flag off" (AC-23). */
export function parseTableUseStyles(tableEl: Element): TableLook | undefined {
  const look: TableLook = {};
  let sawAny = false;
  for (const region of REGIONS) {
    const value = tableEl.getAttributeNS(TABLE_NAMESPACE, region.use);
    if (value === null) {
      continue;
    }
    sawAny = true;
    if (value === "true") {
      look[region.look] = true;
    }
  }
  return sawAny ? look : undefined;
}

function conditionalOf(
  banding: TableStyleBanding,
  key: keyof TableStyleBanding,
): TableConditionalFormat | undefined {
  const value = banding[key];
  if (!value) {
    return undefined;
  }
  return key === "bandedRows" || key === "bandedColumns"
    ? { shading: (value as { shading: string }).shading }
    : (value as TableConditionalFormat);
}

/** The generated `table-cell` styles a template's regions reference, plus the
 * `table:table-template` itself (AC-68) — returned together because both go
 * into `office:styles` and neither is meaningful without the other. */
export function tableTemplateXml(style: NamedStyle): { cellStyles: string; template: string } {
  const banding = style.banding;
  if (!banding) {
    return { cellStyles: "", template: "" };
  }
  let cellStyles = "";
  let children = "";
  for (const region of REGIONS) {
    const format = conditionalOf(banding, region.band);
    if (!format) {
      continue;
    }
    const name = templateCellStyleName(style.id, region.child);
    const properties = [
      format.shading !== undefined ? ` fo:background-color="${escapeXml(format.shading)}"` : "",
      ...BORDER_SIDES.map((side) => {
        const value = format.borders?.[side.key];
        return value ? ` ${side.fo}="${escapeXml(borderShorthand(value))}"` : "";
      }),
    ].join("");
    const text = format.bold ? '<style:text-properties fo:font-weight="bold"/>' : "";
    cellStyles +=
      `<style:style style:name="${escapeXml(name)}" style:family="table-cell">` +
      (properties.length > 0 ? `<style:table-cell-properties${properties}/>` : "") +
      text +
      "</style:style>";
    children += `<table:${region.child} table:style-name="${escapeXml(name)}"/>`;
  }
  if (children.length === 0) {
    return { cellStyles: "", template: "" };
  }
  return {
    cellStyles,
    // The table: namespace is declared on the element itself, the same rule
    // odf/writer.ts's TABLE_XMLNS follows: styles.xml's root declares only
    // office/style/text/fo, and a document with no banding gains nothing.
    template:
      `<table:table-template xmlns:table="${TABLE_NAMESPACE}" table:name="${escapeXml(style.id)}">` +
      `${children}</table:table-template>`,
  };
}

function parseCellStyle(el: Element): TableConditionalFormat | undefined {
  const format: TableConditionalFormat = {};
  const cellProps = firstChildElement(el, STYLE_NAMESPACE, "table-cell-properties");
  if (cellProps) {
    const fill = cellProps.getAttributeNS(FO_NAMESPACE, "background-color");
    if (fill) {
      format.shading = fill;
    }
    const borders: CellBorders = {};
    for (const side of BORDER_SIDES) {
      const parsed = parseBorderShorthand(cellProps.getAttributeNS(FO_NAMESPACE, side.fo.slice("fo:".length)));
      if (parsed) {
        borders[side.key] = parsed;
      }
    }
    if (Object.keys(borders).length > 0) {
      format.borders = borders;
    }
  }
  const textProps = firstChildElement(el, STYLE_NAMESPACE, "text-properties");
  if (textProps?.getAttributeNS(FO_NAMESPACE, "font-weight") === "bold") {
    format.bold = true;
  }
  return Object.keys(format).length > 0 ? format : undefined;
}

/** Every `table:table-template` under `office:styles` as banding, keyed by
 * `table:name` (AC-69). Band sizes always read back as 1: ODF has nowhere to
 * put another value, which is exactly what AC-70 reports on the way out. */
export function parseTableTemplates(officeStyles: Element | undefined): Map<string, TableStyleBanding> {
  const templates = new Map<string, TableStyleBanding>();
  if (!officeStyles) {
    return templates;
  }
  // The generated cell styles live alongside the templates, keyed by name.
  const cellStyles = new Map<string, TableConditionalFormat>();
  for (const styleEl of childElements(officeStyles, STYLE_NAMESPACE, "style")) {
    if (styleEl.getAttributeNS(STYLE_NAMESPACE, "family") !== "table-cell") {
      continue;
    }
    const name = styleEl.getAttributeNS(STYLE_NAMESPACE, "name");
    const format = name ? parseCellStyle(styleEl) : undefined;
    if (name && format) {
      cellStyles.set(name, format);
    }
  }
  for (const templateEl of childElements(officeStyles, TABLE_NAMESPACE, "table-template")) {
    const name = templateEl.getAttributeNS(TABLE_NAMESPACE, "name");
    if (!name) {
      continue;
    }
    const banding: TableStyleBanding = {};
    for (const region of REGIONS) {
      const child = firstChildElement(templateEl, TABLE_NAMESPACE, region.child);
      const styleName = child?.getAttributeNS(TABLE_NAMESPACE, "style-name") ?? undefined;
      const format = styleName ? cellStyles.get(styleName) : undefined;
      if (!format) {
        continue;
      }
      if (region.band === "bandedRows" || region.band === "bandedColumns") {
        if (format.shading !== undefined) {
          banding[region.band] = { size: 1, shading: format.shading };
        }
      } else {
        banding[region.band] = format;
      }
    }
    if (Object.keys(banding).length > 0) {
      templates.set(name, banding);
    }
  }
  return templates;
}
