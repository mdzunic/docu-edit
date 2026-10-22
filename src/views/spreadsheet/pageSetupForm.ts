// The Page Setup dialog's fields (SPEC-136 "The dialog") — framework-free:
// a setup and its print area and titles as the dialog's text fields (margins
// in the SPEC-064 measurement unit), validation (core/pageSetup's
// textAreaTooNarrow / MARGINS_TOO_LARGE_MESSAGE, the 10–400% scale, ranges
// through parseRange), and back to a patch. A blank fit field is automatic
// (0), a blank first page number automatic.
import { MARGINS_TOO_LARGE_MESSAGE, textAreaTooNarrow, type PageSize } from "../../lib/core/pageSetup";
import { formatLength, parseLengthToCm, type MeasurementUnit } from "../../lib/core/units";
import { colToLetters, parseRange, type GridRange } from "../../lib/spreadsheet/address";
import type { HeaderFooterSections, SheetMargins, SheetPageSetup } from "../../lib/spreadsheet/model";
import {
  CUSTOM_PAPER_NAME,
  DEFAULT_SHEET_MARGINS,
  effectiveMargins,
  MAX_PRINT_SCALE,
  MIN_PRINT_SCALE,
  paperForName,
  paperNameOf,
  PRINT_SCALE_MESSAGE,
  XLSX_PAPER_SIZES,
  type PageSetupPatch,
  type PrintTitles,
} from "../../lib/spreadsheet/pageSetup";

export const PAPER_CHOICES: readonly string[] = [...XLSX_PAPER_SIZES.map((entry) => entry.name), CUSTOM_PAPER_NAME];
export const MARGIN_FIELDS = ["top", "right", "bottom", "left", "header", "footer"] as const;
export type MarginField = (typeof MARGIN_FIELDS)[number];
export const INVALID_RANGE_MESSAGE = "Enter a range such as A1:H20";
export const INVALID_TITLES_MESSAGE = "Enter rows such as $1:$2 or columns such as $A:$B";
export const INVALID_NUMBER_MESSAGE = "Enter a whole number";
export const INVALID_MARGIN_MESSAGE = "Enter a valid margin";

/** The header/footer presets: the code each fills the centre with. */
export const HEADER_FOOTER_PRESETS: readonly { id: string; label: string; code: string }[] = [
  { id: "none", label: "None", code: "" },
  { id: "page", label: "Page 1", code: "Page &P" },
  { id: "page-of", label: "Page 1 of ?", code: "Page &P of &N" },
  { id: "sheet", label: "Sheet name", code: "&A" },
  { id: "file", label: "File name", code: "&F" },
  { id: "date", label: "Date", code: "&D" },
];

/** The insert buttons: the code each inserts. */
export const HEADER_FOOTER_INSERTS: readonly { id: string; label: string; code: string }[] = [
  { id: "page", label: "Page", code: "&P" },
  { id: "pages", label: "Pages", code: "&N" },
  { id: "date", label: "Date", code: "&D" },
  { id: "time", label: "Time", code: "&T" },
  { id: "file", label: "File", code: "&F" },
  { id: "sheet", label: "Sheet", code: "&A" },
];

/** The dialog's margin fields, in its order, and its header/footer parts and sections. */
export const MARGIN_LABELS: readonly { field: MarginField; label: string }[] = [
  { field: "top", label: "Top" },
  { field: "bottom", label: "Bottom" },
  { field: "left", label: "Left" },
  { field: "right", label: "Right" },
  { field: "header", label: "Header" },
  { field: "footer", label: "Footer" },
];
export const HF_PARTS: readonly { part: "header" | "footer"; label: string }[] = [
  { part: "header", label: "Header" },
  { part: "footer", label: "Footer" },
];
export const HF_SECTIONS: readonly { id: keyof SectionFields; label: string }[] = [
  { id: "left", label: "Left" },
  { id: "center", label: "Centre" },
  { id: "right", label: "Right" },
];

export interface SectionFields {
  left: string;
  center: string;
  right: string;
}

export interface PageSetupForm {
  orientation: "portrait" | "landscape";
  paper: string;
  /** The custom paper a "Custom" choice keeps (portrait, cm). */
  customPaper: PageSize;
  scaleMode: "percent" | "fit";
  scale: string;
  fitWide: string;
  fitTall: string;
  firstNumber: string;
  margins: Record<MarginField, string>;
  centerHorizontally: boolean;
  centerVertically: boolean;
  header: SectionFields;
  footer: SectionFields;
  differentFirst: boolean;
  differentOddEven: boolean;
  printArea: string;
  titleRows: string;
  titleCols: string;
  gridLines: boolean;
  headings: boolean;
  pageOrder: "down" | "over";
}

const MARGIN_KEYS: Record<MarginField, keyof SheetMargins> = {
  top: "topCm",
  right: "rightCm",
  bottom: "bottomCm",
  left: "leftCm",
  header: "headerCm",
  footer: "footerCm",
};

function sections(value: HeaderFooterSections | undefined): SectionFields {
  return { left: value?.left ?? "", center: value?.center ?? "", right: value?.right ?? "" };
}

function rangeText(range: GridRange): string {
  const a = `$${colToLetters(range.left)}$${range.top + 1}`;
  return range.top === range.bottom && range.left === range.right ? a : `${a}:$${colToLetters(range.right)}$${range.bottom + 1}`;
}

/** The dialog's fields for a sheet's setup, its effective (portrait) paper,
 * its print area and titles, in `unit`. */
export function formFromSetup(setup: SheetPageSetup | undefined, paper: PageSize, unit: MeasurementUnit, area: readonly GridRange[] | null, titles: PrintTitles | null): PageSetupForm {
  const margins = effectiveMargins(setup);
  const fit = setup?.fitTo ?? (setup?.odsScaleToPages !== undefined ? { width: 1, height: setup.odsScaleToPages } : undefined);
  return {
    orientation: setup?.orientation === "landscape" ? "landscape" : "portrait",
    paper: paperNameOf(setup?.paper, paper),
    customPaper: { widthCm: paper.widthCm, heightCm: paper.heightCm },
    scaleMode: fit === undefined ? "percent" : "fit",
    scale: String(setup?.scale ?? 100),
    fitWide: fit === undefined || fit.width === 0 ? (fit === undefined ? "1" : "") : String(fit.width),
    fitTall: fit === undefined || fit.height === 0 ? (fit === undefined ? "1" : "") : String(fit.height),
    firstNumber: setup?.firstPageNumber === undefined ? "" : String(setup.firstPageNumber),
    margins: Object.fromEntries(MARGIN_FIELDS.map((field) => [field, formatLength(margins[MARGIN_KEYS[field]], unit)])) as Record<MarginField, string>,
    centerHorizontally: setup?.centerHorizontally === true,
    centerVertically: setup?.centerVertically === true,
    header: sections(setup?.headerFooter?.odd.header),
    footer: sections(setup?.headerFooter?.odd.footer),
    differentFirst: setup?.headerFooter?.differentFirst === true,
    differentOddEven: setup?.headerFooter?.differentOddEven === true,
    printArea: area === null ? "" : area.map(rangeText).join(","),
    titleRows: titles?.rows === undefined ? "" : `$${titles.rows[0] + 1}:$${titles.rows[1] + 1}`,
    titleCols: titles?.cols === undefined ? "" : `$${colToLetters(titles.cols[0])}:$${colToLetters(titles.cols[1])}`,
    gridLines: setup?.printGridLines === true,
    headings: setup?.printHeadings === true,
    pageOrder: setup?.pageOrder === "overThenDown" ? "over" : "down",
  };
}

/** The portrait paper the form names. */
export function formPaper(form: PageSetupForm): NonNullable<SheetPageSetup["paper"]> {
  return paperForName(form.paper) ?? { widthCm: form.customPaper.widthCm, heightCm: form.customPaper.heightCm };
}

function wholeOrBlank(text: string, min: number, max: number): number | null | undefined {
  const trimmed = text.trim();
  if (trimmed === "") {
    return undefined;
  }
  if (!/^\d+$/.test(trimmed)) {
    return null;
  }
  const value = Number(trimmed);
  return value < min || value > max ? null : value;
}

/** The print area field's ranges (comma-separated, each through parseRange);
 * null for a blank field, "invalid" for one that does not parse. */
export function parsePrintAreaField(text: string): GridRange[] | null | "invalid" {
  const trimmed = text.trim();
  if (trimmed === "") {
    return null;
  }
  const ranges = trimmed.split(",").map((part) => parseRange(part.trim().replace(/^[^!]*!/, "")));
  return ranges.some((range) => range === null) ? "invalid" : (ranges as GridRange[]);
}

/** A print titles field: "$1:$2" (rows) or "$A:$B" (columns), checked
 * through parseRange; undefined when blank, null when invalid. */
export function parseTitlesField(text: string, axis: "row" | "col"): [number, number] | undefined | null {
  const trimmed = text.trim().replace(/^[^!]*!/, "");
  if (trimmed === "") {
    return undefined;
  }
  const match = axis === "row" ? /^\$?(\d+)(?::\$?(\d+))?$/.exec(trimmed) : /^\$?([A-Za-z]+)(?::\$?([A-Za-z]+))?$/.exec(trimmed);
  if (match === null) {
    return null;
  }
  const range = axis === "row" ? parseRange(`A${match[1]}:A${match[2] ?? match[1]}`) : parseRange(`${match[1]}1:${match[2] ?? match[1]}1`);
  if (range === null) {
    return null;
  }
  return axis === "row" ? [range.top, range.bottom] : [range.left, range.right];
}

export interface FormProblem {
  field: string;
  message: string;
}

/** The first problem with the form, or null. */
export function validateForm(form: PageSetupForm, unit: MeasurementUnit): FormProblem | null {
  if (form.scaleMode === "percent" && wholeOrBlank(form.scale, MIN_PRINT_SCALE, MAX_PRINT_SCALE) == null) {
    return { field: "scale", message: PRINT_SCALE_MESSAGE };
  }
  if (form.scaleMode === "fit" && (wholeOrBlank(form.fitWide, 0, 32767) === null || wholeOrBlank(form.fitTall, 0, 32767) === null)) {
    return { field: "fit", message: INVALID_NUMBER_MESSAGE };
  }
  if (wholeOrBlank(form.firstNumber, 1, 2147483647) === null) {
    return { field: "firstNumber", message: INVALID_NUMBER_MESSAGE };
  }
  const margins: Partial<SheetMargins> = {};
  for (const field of MARGIN_FIELDS) {
    const cm = parseLengthToCm(form.margins[field], unit);
    if (cm === null || cm < 0) {
      return { field: `margin-${field}`, message: INVALID_MARGIN_MESSAGE };
    }
    margins[MARGIN_KEYS[field]] = cm;
  }
  const paper = formPaper(form);
  const oriented = form.orientation === "landscape" ? { widthCm: paper.heightCm, heightCm: paper.widthCm } : paper;
  if (textAreaTooNarrow(oriented, margins as SheetMargins)) {
    return { field: "margins", message: MARGINS_TOO_LARGE_MESSAGE };
  }
  if (parsePrintAreaField(form.printArea) === "invalid") {
    return { field: "printArea", message: INVALID_RANGE_MESSAGE };
  }
  if (parseTitlesField(form.titleRows, "row") === null || parseTitlesField(form.titleCols, "col") === null) {
    return { field: "titles", message: INVALID_TITLES_MESSAGE };
  }
  return null;
}

function sectionsOf(fields: SectionFields): HeaderFooterSections | undefined {
  const out: HeaderFooterSections = {};
  for (const key of ["left", "center", "right"] as const) {
    if (fields[key] !== "") {
      out[key] = fields[key];
    }
  }
  return Object.keys(out).length === 0 ? undefined : out;
}

/** What a valid form sets: the setup patch, the print area and titles. */
export function setupFromForm(
  form: PageSetupForm,
  unit: MeasurementUnit,
  base: SheetPageSetup | undefined,
): { patch: PageSetupPatch; printArea: GridRange[] | null; titles: PrintTitles | null } {
  const margins = Object.fromEntries(MARGIN_FIELDS.map((field) => [MARGIN_KEYS[field], parseLengthToCm(form.margins[field], unit) ?? DEFAULT_SHEET_MARGINS[MARGIN_KEYS[field]]])) as unknown as SheetMargins;
  const hf = base?.headerFooter;
  const header = sectionsOf(form.header);
  const footer = sectionsOf(form.footer);
  const odd = { ...(header ? { header } : {}), ...(footer ? { footer } : {}) };
  const headerFooter = {
    ...hf,
    odd,
    ...(form.differentFirst ? { differentFirst: true as const } : { differentFirst: undefined }),
    ...(form.differentOddEven ? { differentOddEven: true as const } : { differentOddEven: undefined }),
  };
  const fitWide = wholeOrBlank(form.fitWide, 0, 32767) ?? 0;
  const fitTall = wholeOrBlank(form.fitTall, 0, 32767) ?? 0;
  const patch: PageSetupPatch = {
    orientation: form.orientation === "landscape" ? "landscape" : undefined,
    paper: formPaper(form),
    scale: form.scaleMode === "percent" ? (wholeOrBlank(form.scale, MIN_PRINT_SCALE, MAX_PRINT_SCALE) ?? 100) : undefined,
    fitTo: form.scaleMode === "fit" ? { width: fitWide, height: fitTall } : undefined,
    odsScaleToPages: undefined,
    margins,
    centerHorizontally: form.centerHorizontally ? true : undefined,
    centerVertically: form.centerVertically ? true : undefined,
    printGridLines: form.gridLines ? true : undefined,
    printHeadings: form.headings ? true : undefined,
    pageOrder: form.pageOrder === "over" ? "overThenDown" : undefined,
    firstPageNumber: wholeOrBlank(form.firstNumber, 1, 2147483647) ?? undefined,
    headerFooter: JSON.parse(JSON.stringify(headerFooter)) as SheetPageSetup["headerFooter"],
  };
  // ods: a scale-to-pages setup the dialog left alone stays as it was.
  if (base?.odsScaleToPages !== undefined && form.scaleMode === "fit" && fitWide === 1 && fitTall === base.odsScaleToPages && base.fitTo === undefined) {
    patch.fitTo = undefined;
    patch.odsScaleToPages = base.odsScaleToPages;
  }
  const area = parsePrintAreaField(form.printArea);
  const rows = parseTitlesField(form.titleRows, "row") ?? undefined;
  const cols = parseTitlesField(form.titleCols, "col") ?? undefined;
  return {
    patch,
    printArea: area === "invalid" ? null : area,
    titles: rows === undefined && cols === undefined ? null : { ...(rows ? { rows } : {}), ...(cols ? { cols } : {}) },
  };
}
