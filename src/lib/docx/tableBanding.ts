// Table look and style banding for .docx (SPEC-076 AC-66, AC-67, AC-71).
//
// Two halves of one wire format:
//   - the TABLE says which regions it uses — `w:tblPr/w:tblLook`, whose
//     `w:noHBand`/`w:noVBand` are the INVERSE of the model's bandedRows /
//     bandedColumns;
//   - the STYLE says what each region looks like — `w:tblStylePr` elements
//     of type firstRow/lastRow/firstCol/lastCol/band1Horz/band1Vert, plus
//     the band sizes in `w:tblPr/w:tblStyleRowBandSize` and
//     `w:tblStyleColBandSize`.
//
// Anything outside that vocabulary is what keeps a style `unmodeled` and
// re-emitted verbatim (SPEC-042 AC-18, restated by AC-71); a style whose
// conditional formatting fits entirely inside `TableStyleBanding` is modelled
// and editable. `parseTableStyleBanding` is the single place that judgement
// is made, so the reader and the "is Modify available" question can never
// disagree.
import { WORD_NAMESPACE, wordAttr } from "./xml";
import { childElements, firstChildElement } from "../core/xml";
import { borderElementXml, parseBorderElement } from "./borders";
import { cssColorToOoxml, ooxmlColorToCss } from "../core/colors";
import { resolveColor, type DocumentTheme } from "../core/theme";
import type { CellBorders } from "../core/tables";
import type { TableConditionalFormat, TableLook, TableStyleBanding } from "../documentModel";

/** The six `w:tblStylePr` region types this spec models, paired with the
 * `TableStyleBanding` field each maps onto. Any OTHER `w:type` is what makes
 * a style unmodeled (AC-71). */
const REGION_TYPES: { wire: string; key: keyof TableStyleBanding }[] = [
  { wire: "firstRow", key: "firstRow" },
  { wire: "lastRow", key: "lastRow" },
  { wire: "firstCol", key: "firstColumn" },
  { wire: "lastCol", key: "lastColumn" },
  { wire: "band1Horz", key: "bandedRows" },
  { wire: "band1Vert", key: "bandedColumns" },
];

/** The four sides in `w:tcBorders` order, with the wire element each uses.
 * Left/right write as `w:start`/`w:end` — the same names cellBordersXml has
 * always emitted — and read from either spelling (table RTL is out of scope,
 * so logical and physical sides coincide). */
const BORDER_SIDES: { key: keyof CellBorders; wire: string; legacy?: string }[] = [
  { key: "top", wire: "top" },
  { key: "left", wire: "start", legacy: "left" },
  { key: "bottom", wire: "bottom" },
  { key: "right", wire: "end", legacy: "right" },
];

const on = (flag: boolean | undefined) => (flag ? "1" : "0");

/** `w:tblLook` for a table's `look` (AC-66). `w:noHBand`/`w:noVBand` invert
 * `bandedRows`/`bandedColumns`: Word states which bands to SUPPRESS. The
 * legacy `w:val` bitmask is emitted alongside the explicit attributes, which
 * is what Word itself writes and what older consumers read. */
export function tableLookXml(look: TableLook | undefined): string {
  if (!look) {
    return "";
  }
  const mask =
    (look.firstRow ? 0x20 : 0) +
    (look.lastRow ? 0x40 : 0) +
    (look.firstColumn ? 0x80 : 0) +
    (look.lastColumn ? 0x100 : 0) +
    (look.bandedRows ? 0 : 0x200) +
    (look.bandedColumns ? 0 : 0x400);
  return (
    `<w:tblLook w:val="${mask.toString(16).toUpperCase().padStart(4, "0")}"` +
    ` w:firstRow="${on(look.firstRow)}" w:lastRow="${on(look.lastRow)}"` +
    ` w:firstColumn="${on(look.firstColumn)}" w:lastColumn="${on(look.lastColumn)}"` +
    ` w:noHBand="${on(!look.bandedRows)}" w:noVBand="${on(!look.bandedColumns)}"/>`
  );
}

function flagOn(el: Element, name: string): boolean {
  const value = wordAttr(el, name);
  return value === "1" || value === "true" || value === "on";
}

/** A table's `look` from its `w:tblPr` (AC-67), or undefined when the table
 * carries no `w:tblLook` at all — which is "no banding", not "all off"
 * (AC-23/D-10: an imported table with no look flags renders as it does
 * today). A `w:tblLook` with every flag off returns an empty object, which
 * resolves the same way but round-trips the element the source had. */
export function parseTableLook(tblPr: Element | undefined): TableLook | undefined {
  const el = tblPr ? firstChildElement(tblPr, WORD_NAMESPACE, "tblLook") : undefined;
  if (!el) {
    return undefined;
  }
  const look: TableLook = {};
  if (flagOn(el, "firstRow")) {
    look.firstRow = true;
  }
  if (flagOn(el, "lastRow")) {
    look.lastRow = true;
  }
  if (flagOn(el, "firstColumn")) {
    look.firstColumn = true;
  }
  if (flagOn(el, "lastColumn")) {
    look.lastColumn = true;
  }
  // The inverse again: no w:noHBand means the row bands ARE used.
  if (!flagOn(el, "noHBand")) {
    look.bandedRows = true;
  }
  if (!flagOn(el, "noVBand")) {
    look.bandedColumns = true;
  }
  return look;
}

function conditionalXml(format: TableConditionalFormat | undefined, theme: DocumentTheme): string {
  if (!format) {
    return "";
  }
  const rPr = format.bold ? "<w:rPr><w:b/></w:rPr>" : "";
  const sides = BORDER_SIDES.map((side) => {
    const value = format.borders?.[side.key];
    return value ? borderElementXml(side.wire, value, theme) : "";
  }).join("");
  const borders = sides.length > 0 ? `<w:tcBorders>${sides}</w:tcBorders>` : "";
  const fill = cssColorToOoxml(format.shading === undefined ? undefined : resolveColor(format.shading, theme));
  const shd = fill !== undefined ? `<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>` : "";
  const tcPr = borders + shd;
  return rPr + (tcPr.length > 0 ? `<w:tcPr>${tcPr}</w:tcPr>` : "");
}

/** The band sizes for a table style's own `w:tblPr` (AC-66) — one element per
 * band the style declares, size 1 included. Writing the default explicitly
 * costs nothing and keeps the style self-describing for a consumer that does
 * not know OOXML's default. */
export function tableStyleBandSizesXml(banding: TableStyleBanding | undefined): string {
  const rowSize = banding?.bandedRows?.size;
  const colSize = banding?.bandedColumns?.size;
  return (
    (rowSize !== undefined ? `<w:tblStyleRowBandSize w:val="${Math.max(1, Math.floor(rowSize))}"/>` : "") +
    (colSize !== undefined ? `<w:tblStyleColBandSize w:val="${Math.max(1, Math.floor(colSize))}"/>` : "")
  );
}

/** The `w:tblStylePr` elements for a style's banding (AC-66), in Word's own
 * region order. A region the model does not carry emits nothing. */
export function tableStylePrXml(banding: TableStyleBanding | undefined, theme: DocumentTheme): string {
  if (!banding) {
    return "";
  }
  return REGION_TYPES.map((region) => {
    // Every region — band or edge — is a full conditional format (AC-37);
    // a band's `size` is not here but in tableStyleBandSizesXml.
    const format: TableConditionalFormat | undefined = banding[region.key];
    const body = conditionalXml(format, theme);
    return body.length > 0 ? `<w:tblStylePr w:type="${region.wire}">${body}</w:tblStylePr>` : "";
  }).join("");
}

function parseConditional(el: Element): TableConditionalFormat | undefined {
  const format: TableConditionalFormat = {};
  const rPr = firstChildElement(el, WORD_NAMESPACE, "rPr");
  if (rPr && firstChildElement(rPr, WORD_NAMESPACE, "b")) {
    format.bold = true;
  }
  const tcPr = firstChildElement(el, WORD_NAMESPACE, "tcPr");
  if (tcPr) {
    const shd = firstChildElement(tcPr, WORD_NAMESPACE, "shd");
    const fill = shd ? ooxmlColorToCss(wordAttr(shd, "fill")) : undefined;
    if (fill !== undefined) {
      format.shading = fill;
    }
    const bordersEl = firstChildElement(tcPr, WORD_NAMESPACE, "tcBorders");
    if (bordersEl) {
      const borders: CellBorders = {};
      for (const side of BORDER_SIDES) {
        const sideEl =
          firstChildElement(bordersEl, WORD_NAMESPACE, side.wire) ??
          (side.legacy ? firstChildElement(bordersEl, WORD_NAMESPACE, side.legacy) : undefined);
        const parsed = parseBorderElement(sideEl, wordAttr);
        if (parsed) {
          borders[side.key] = parsed;
        }
      }
      if (Object.keys(borders).length > 0) {
        format.borders = borders;
      }
    }
  }
  return Object.keys(format).length > 0 ? format : undefined;
}

function bandSize(styleEl: Element, tag: string): number {
  const el = firstChildElement(firstChildElement(styleEl, WORD_NAMESPACE, "tblPr") ?? styleEl, WORD_NAMESPACE, tag);
  const value = el ? Number(wordAttr(el, "val") ?? "") : NaN;
  return Number.isFinite(value) && value >= 1 ? Math.floor(value) : 1;
}

/** A table style's banding, and whether the style's conditional formatting
 * is FULLY expressible as `TableStyleBanding` (AC-67, AC-71).
 *
 * `modelled` false means at least one `w:tblStylePr` names a region outside
 * the six — the style keeps `unmodeled: true` and its verbatim rawXml, and
 * Modify stays unavailable for it. `modelled` true with no banding at all is
 * an ordinary table style with no conditional formatting, which is also
 * fully modelled. */
export function parseTableStyleBanding(styleEl: Element): {
  banding?: TableStyleBanding;
  modelled: boolean;
} {
  const banding: TableStyleBanding = {};
  let modelled = true;
  for (const el of childElements(styleEl, WORD_NAMESPACE, "tblStylePr")) {
    const wire = wordAttr(el, "type") ?? "";
    const region = REGION_TYPES.find((entry) => entry.wire === wire);
    if (!region) {
      modelled = false;
      continue;
    }
    const format = parseConditional(el);
    if (!format) {
      continue;
    }
    if (region.key === "bandedRows" || region.key === "bandedColumns") {
      // A band is a full conditional format plus its size (AC-37) — bold and
      // borders on a band region are modelled since SPEC-076's style editor
      // edits them, so they no longer force the style verbatim.
      const size = bandSize(styleEl, region.key === "bandedRows" ? "tblStyleRowBandSize" : "tblStyleColBandSize");
      banding[region.key] = { size, ...format };
    } else {
      banding[region.key] = format;
    }
  }
  return Object.keys(banding).length > 0 ? { banding, modelled } : { modelled };
}
