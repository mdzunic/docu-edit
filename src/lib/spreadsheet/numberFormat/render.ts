// Rendering a cell value through a format code (SPEC-118 "Sections",
// "Rendering numbers", "Rendering dates and times", "Text, booleans and
// errors", "General", "editText"). formatValue never throws: an invalid or
// unsupported code renders as General.
//
// - Sections. Without conditions: one section → every number; two → (≥ 0;
//   < 0); three → (> 0; < 0; = 0); a fourth (or, with fewer, a final section
//   containing @) → text. With conditions on the first and/or second section
//   the first whose condition holds wins, the first unconditioned numeric
//   section is the "else", and no match is `overflow: "hash"`. The automatic
//   "-" is dropped only in section 2 of a code whose first two sections carry
//   no condition (the classic negative section writes its own sign).
// - Numbers: |v| × 100 per `%`, ÷ 1000 per scaling comma, taken at 15
//   significant digits and rounded half away from zero (decimal.ts); digits
//   fill the integer placeholders right to left, `0` pads with "0", `#` with
//   nothing, `?` with U+2007 FIGURE SPACE; the locale's group and decimal
//   characters stand in for the code's invariant `,` and `.`. Scientific,
//   fractions (best rational approximation, ties → the smaller denominator)
//   and `General` in place per the Reference.
// - Dates and times: serial × 86 400 s rounded half away from zero to the
//   section's fractional-second digits, carries reaching the date; Excel's
//   calendar (1900-02-29, day 0, its weekdays) under `dateQuirks: "excel"`,
//   proleptic Gregorian for an ods-born workbook; names from localeNames.
// - Fitting. With `widthPx`, General shrinks (formatGeneral) and any other
//   number or date that does not fit is `#` repeated to the width; text never
//   shrinks or hashes.
import type { Cell, CellStyle, CellValue, SpreadsheetModel } from "../model";
import { isCellError } from "../model";
import { CELL_PADDING_PX, DEFAULT_COL_WIDTH_PX, estimateTextWidth } from "../geometry";
import { civilFromDays, daysFromCivil, serialToParts } from "../serial";
import { INDEXED_COLORS } from "../styles";
import { formatCodeOf } from "./builtin";
import {
  decimalParts,
  fixedCandidate,
  fixedDecimals,
  generalNumberText,
  plainNotation,
  roundDigits,
  roundToPlaces,
  scientificCandidate,
  significand,
} from "./decimal";
import {
  localeNames,
  localeTagOfLcid,
  SPREADSHEET_LOCALE_NAMES,
  SPREADSHEET_LOCALES,
  type SpreadsheetLocale,
  type SpreadsheetLocaleTag,
} from "./locale";
import {
  LCID_SYSTEM_LONG_DATE,
  LCID_SYSTEM_LONG_TIME,
  parseFormatCode,
  type FormatSection,
  type FormatToken,
  type ParsedFormat,
} from "./parse";

export interface FormatContext {
  locale: SpreadsheetLocale;
  dateSystem: "1900" | "1904";
  /** The cell's width (column or merge); absent = full width, no fitting. */
  widthPx?: number;
  /** The MDW of the cell font; default 7. */
  digitWidthPx?: number;
  /** Default "excel"; the view passes "proleptic" for format === "ods". */
  dateQuirks?: "excel" | "proleptic";
  /** The [ColorN] palette; default INDEXED_COLORS (model.indexedColors when set). */
  palette?: readonly string[];
}

export interface FormattedValue {
  text: string;
  /** "#rrggbb" from [Red] / [ColorN]. */
  color?: string;
  /** For horizontal "general". */
  alignHint?: "left" | "right" | "center";
  /** The text is the # fill. */
  overflow?: "hash";
  /** *x: expand `char` at `index` to fill the width. */
  fill?: { char: string; index: number };
}

export interface EditTextContext extends FormatContext {
  /** The cell's effective style; absent = Normal. */
  style?: CellStyle;
}

/** The context a workbook's cells format in, in `locale`: its date system,
 * Excel's calendar quirks for an xlsx-born workbook and the proleptic
 * Gregorian calendar for an ods-born one, and its [ColorN] palette. */
export function workbookFormatContext(
  model: Pick<SpreadsheetModel, "dateSystem" | "format" | "indexedColors">,
  locale: SpreadsheetLocale,
): FormatContext {
  const ctx: FormatContext = { locale, dateSystem: model.dateSystem, dateQuirks: model.format === "ods" ? "proleptic" : "excel" };
  if (model.indexedColors !== undefined) {
    ctx.palette = model.indexedColors;
  }
  return ctx;
}

const FIGURE_SPACE = "\u2007";
const DEFAULT_DIGIT_WIDTH_PX = 7;
const SECONDS_PER_DAY = 86_400;
/** The eight colour names' hex. */
const NAMED_COLORS: Readonly<Record<string, string>> = {
  Black: "#000000",
  Blue: "#0000ff",
  Cyan: "#00ffff",
  Green: "#00ff00",
  Magenta: "#ff00ff",
  Red: "#ff0000",
  White: "#ffffff",
  Yellow: "#ffff00",
};

// --- the output being built ----------------------------------------------------

/** A section's text as it is assembled, with its fill position. */
class Output {
  text = "";
  fill: { char: string; index: number } | undefined;

  add(text: string): void {
    this.text += text;
  }

  addFill(char: string): void {
    this.fill = { char, index: this.text.length };
  }

  /** Prepends text (the automatic minus), keeping the fill in place. */
  prepend(text: string): void {
    this.text = text + this.text;
    if (this.fill) {
      this.fill = { char: this.fill.char, index: this.fill.index + text.length };
    }
  }
}

// --- shared helpers ----------------------------------------------------------------

/** The px a cell's text may use: its width less the padding and gridline. */
function availablePx(widthPx: number): number {
  return widthPx - 2 * CELL_PADDING_PX - 1;
}

function digitWidth(ctx: FormatContext): number {
  return ctx.digitWidthPx !== undefined && ctx.digitWidthPx > 0 ? ctx.digitWidthPx : DEFAULT_DIGIT_WIDTH_PX;
}

/** `#` × max(1, floor(available / MDW)) — Excel's overflow fill; a cell
 * with no width gets the default column's. */
function hashText(ctx: FormatContext): string {
  const width = ctx.widthPx ?? DEFAULT_COL_WIDTH_PX;
  const count = Math.max(1, Math.floor(availablePx(width) / digitWidth(ctx)));
  return "#".repeat(count);
}

function hashed(ctx: FormatContext, color?: string): FormattedValue {
  const value: FormattedValue = { text: hashText(ctx), alignHint: "right", overflow: "hash" };
  if (color !== undefined) {
    value.color = color;
  }
  return value;
}

/** Whether a text (plus one fill character) fits the context's width. */
function fits(text: string, fill: FormattedValue["fill"], ctx: FormatContext): boolean {
  if (ctx.widthPx === undefined) {
    return true;
  }
  return estimateTextWidth(fill ? text + fill.char : text, digitWidth(ctx)) <= availablePx(ctx.widthPx);
}

/** A section colour as "#rrggbb": the eight names, or [ColorN] → palette[N + 7]. */
function sectionColor(section: FormatSection, ctx: FormatContext): string | undefined {
  const color = section.color;
  if (color === undefined) {
    return undefined;
  }
  const named = NAMED_COLORS[color];
  if (named !== undefined) {
    return named;
  }
  const index = /^Color(\d+)$/.exec(color);
  const palette = ctx.palette ?? INDEXED_COLORS;
  return index ? palette[Number(index[1]) + 7] : undefined;
}

function finish(out: Output, alignHint: FormattedValue["alignHint"], color: string | undefined): FormattedValue {
  const value: FormattedValue = { text: out.text, alignHint };
  if (color !== undefined) {
    value.color = color;
  }
  if (out.fill !== undefined) {
    value.fill = out.fill;
  }
  return value;
}

/** The code parsed, or null when it renders as General (absent, "General",
 * invalid or unsupported). */
function parsedCode(code: string | undefined): ParsedFormat | null {
  if (code === undefined || code.trim() === "" || /^general$/i.test(code.trim())) {
    return null;
  }
  try {
    const parsed = parseFormatCode(code);
    return parsed.unsupported ? null : parsed;
  } catch {
    return null;
  }
}

/** The index of the text section: the fourth, or a final section containing
 * @ when there are fewer; -1 when there is none. */
function textSectionIndex(sections: readonly FormatSection[]): number {
  if (sections.length === 4) {
    return 3;
  }
  const last = sections.length - 1;
  return sections[last]!.kind === "text" ? last : -1;
}

function conditionHolds(condition: NonNullable<FormatSection["condition"]>, v: number): boolean {
  switch (condition.op) {
    case "<":
      return v < condition.value;
    case "<=":
      return v <= condition.value;
    case ">":
      return v > condition.value;
    case ">=":
      return v >= condition.value;
    case "=":
      return v === condition.value;
    default:
      return v !== condition.value;
  }
}

/** The section a number renders through and whether it gets the automatic
 * minus; null when a conditioned code has no section for it (→ hash). */
function chooseSection(numeric: readonly FormatSection[], v: number): { section: FormatSection; autoMinus: boolean } | null {
  const negative = v < 0;
  if (numeric.slice(0, 2).some((candidate) => candidate.condition !== undefined)) {
    const matched = numeric.find((candidate) => candidate.condition !== undefined && conditionHolds(candidate.condition, v));
    const section = matched ?? numeric.find((candidate) => candidate.condition === undefined);
    return section ? { section, autoMinus: negative } : null;
  }
  if (numeric.length === 1) {
    return { section: numeric[0]!, autoMinus: negative };
  }
  if (numeric.length === 2) {
    return negative ? { section: numeric[1]!, autoMinus: false } : { section: numeric[0]!, autoMinus: false };
  }
  if (negative) {
    return { section: numeric[1]!, autoMinus: false };
  }
  return { section: v > 0 ? numeric[0]! : numeric[2]!, autoMinus: false };
}

/** The numeric sections of a code (all but its text section). */
function numericSections(parsed: ParsedFormat): FormatSection[] {
  const textIndex = textSectionIndex(parsed.sections);
  return parsed.sections.filter((_, index) => index !== textIndex);
}

/** The section a number would render through (null: none). */
function sectionForNumber(parsed: ParsedFormat, v: number): FormatSection | null {
  const numeric = numericSections(parsed);
  return numeric.length === 0 ? null : (chooseSection(numeric, v)?.section ?? null);
}

/** Whether a section renders a date or time (its tokens, or a system
 * long-date/long-time LCID). */
function isDateSection(section: FormatSection): boolean {
  return section.kind === "date" || section.lcid === LCID_SYSTEM_LONG_DATE || section.lcid === LCID_SYSTEM_LONG_TIME;
}

// --- numbers ---------------------------------------------------------------------

type Placeholder = "0" | "#" | "?";

/** Where a number section's parts are, analysed once per (cached) section. */
interface NumberLayout {
  percent: number;
  scale: number;
  grouping: boolean;
  /** Commas consumed as grouping or scaling (printed as nothing). */
  silentCommas: Set<number>;
  point: number;
  /** Integer placeholders (the whole part of a fraction, a mantissa's
   * integer digits); fraction placeholders (a mantissa's decimals). */
  intSlots: number[];
  fracSlots: number[];
  exponent: number;
  expSlots: number[];
  numSlots: number[];
  slash: number;
  denSlots: number[];
  denominator: number | undefined;
}

const layouts = new WeakMap<FormatSection, NumberLayout>();

function digitIndices(tokens: readonly FormatToken[], from: number, to: number): number[] {
  const out: number[] = [];
  for (let i = Math.max(0, from); i < Math.min(to, tokens.length); i += 1) {
    if (tokens[i]!.type === "digit") {
      out.push(i);
    }
  }
  return out;
}

/** Classifies the commas of an integer region [from, to): between two
 * integer placeholders → grouping; after the last → scaling; else literal. */
function integerCommas(tokens: readonly FormatToken[], slots: readonly number[], from: number, to: number, layout: NumberLayout): void {
  if (slots.length === 0) {
    return;
  }
  const first = slots[0]!;
  const last = slots.at(-1)!;
  for (let i = from; i < to; i += 1) {
    if (tokens[i]!.type !== "comma" || i < first) {
      continue;
    }
    layout.silentCommas.add(i);
    if (i < last) {
      layout.grouping = true;
    } else {
      layout.scale += 1;
    }
  }
}

function analyse(section: FormatSection): NumberLayout {
  let layout = layouts.get(section);
  if (layout) {
    return layout;
  }
  const tokens = section.tokens;
  layout = {
    percent: tokens.filter((token) => token.type === "percent").length,
    scale: 0,
    grouping: false,
    silentCommas: new Set(),
    point: -1,
    intSlots: [],
    fracSlots: [],
    exponent: -1,
    expSlots: [],
    numSlots: [],
    slash: -1,
    denSlots: [],
    denominator: undefined,
  };
  if (section.kind === "fraction") {
    const slash = tokens.findIndex((token) => token.type === "slash");
    layout.slash = slash;
    let start = slash;
    while (start > 0 && tokens[start - 1]!.type === "digit") {
      start -= 1;
    }
    layout.numSlots = digitIndices(tokens, start, slash);
    layout.intSlots = digitIndices(tokens, 0, start);
    integerCommas(tokens, layout.intSlots, 0, start, layout);
    const after = tokens[slash + 1];
    if (after?.type === "denominator") {
      layout.denominator = after.value;
    } else {
      let end = slash + 1;
      while (end < tokens.length && tokens[end]!.type === "digit") {
        end += 1;
      }
      layout.denSlots = digitIndices(tokens, slash + 1, end);
    }
  } else {
    layout.exponent = tokens.findIndex((token) => token.type === "exponent");
    const mantissaEnd = layout.exponent >= 0 ? layout.exponent : tokens.length;
    layout.point = tokens.findIndex((token, index) => token.type === "point" && index < mantissaEnd);
    const intEnd = layout.point >= 0 ? layout.point : mantissaEnd;
    layout.intSlots = digitIndices(tokens, 0, intEnd);
    layout.fracSlots = layout.point >= 0 ? digitIndices(tokens, layout.point + 1, mantissaEnd) : [];
    layout.expSlots = layout.exponent >= 0 ? digitIndices(tokens, layout.exponent + 1, tokens.length) : [];
    integerCommas(tokens, layout.intSlots, 0, intEnd, layout);
    if (layout.exponent >= 0) {
      layout.grouping = false; // a mantissa is never grouped
      layout.scale = 0;
    }
    const lastFraction = layout.fracSlots.at(-1);
    if (lastFraction !== undefined && layout.exponent < 0) {
      // Commas right after the last decimal also scale (`0.0,`).
      for (let i = lastFraction + 1; i < tokens.length && tokens[i]!.type === "comma"; i += 1) {
        layout.silentCommas.add(i);
        layout.scale += 1;
      }
    }
  }
  layouts.set(section, layout);
  return layout;
}

function placeholderOf(token: FormatToken | undefined): Placeholder {
  return token?.type === "digit" ? token.placeholder : "#";
}

/** What an unfilled placeholder shows. */
function pad(placeholder: Placeholder): string {
  return placeholder === "0" ? "0" : placeholder === "?" ? FIGURE_SPACE : "";
}

/** Integer digits into placeholders, right to left: unfilled ones pad, extra
 * leading digits precede the first; `group` (when grouping) every three
 * digits. Returns each slot's text. */
function integerSlotTexts(placeholders: readonly Placeholder[], digits: string, group: string | null): string[] {
  const out = placeholders.map(() => "");
  let remaining = digits.length;
  for (let j = placeholders.length - 1; j >= 0; j -= 1) {
    if (remaining > 0) {
      out[j] = digits[remaining - 1]!;
      remaining -= 1;
    } else {
      out[j] = pad(placeholders[j]!);
    }
  }
  if (remaining > 0 && out.length > 0) {
    out[0] = digits.slice(0, remaining) + out[0];
  }
  if (group !== null) {
    let seen = 0;
    for (let j = out.length - 1; j >= 0; j -= 1) {
      let rebuilt = "";
      const chars = Array.from(out[j]!);
      for (let c = chars.length - 1; c >= 0; c -= 1) {
        const ch = chars[c]!;
        if (ch >= "0" && ch <= "9") {
          rebuilt = (seen > 0 && seen % 3 === 0 ? ch + group : ch) + rebuilt;
          seen += 1;
        } else {
          rebuilt = ch + rebuilt;
        }
      }
      out[j] = rebuilt;
    }
  }
  return out;
}

/** Fraction digits into placeholders, left to right; trailing zeros under
 * `#` drop and under `?` become figure spaces. */
function fractionSlotTexts(placeholders: readonly Placeholder[], digits: string): string[] {
  const out = placeholders.map((_, k) => digits[k] ?? "0");
  for (let k = placeholders.length - 1; k >= 0; k -= 1) {
    if (out[k] !== "0" || placeholders[k] === "0") {
      break;
    }
    out[k] = pad(placeholders[k]!);
  }
  return out;
}

/** Grouped integer digits with no placeholders to hold them. */
function groupDigits(digits: string, group: string | null): string {
  return group === null ? digits : digits.replace(/\B(?=(\d{3})+(?!\d))/g, group);
}

/** Emits a number section's non-numeric tokens; `slot(i)` renders the
 * numeric token at index i (or returns undefined to treat it as literal). */
function emitTokens(
  section: FormatSection,
  out: Output,
  ctx: FormatContext,
  magnitude: number,
  slot: (index: number, token: FormatToken) => string | undefined,
): void {
  section.tokens.forEach((token, index) => {
    const special = slot(index, token);
    if (special !== undefined) {
      out.add(special);
      return;
    }
    switch (token.type) {
      case "literal":
        out.add(token.text);
        break;
      case "space":
        out.add(" ");
        break;
      case "fill":
        out.addFill(token.char);
        break;
      case "currency":
        out.add(token.symbol);
        break;
      case "percent":
        out.add("%");
        break;
      case "comma":
        out.add(",");
        break;
      case "point":
        out.add(".");
        break;
      case "slash":
        out.add("/");
        break;
      case "digit":
        out.add(token.placeholder);
        break;
      case "general":
        out.add(generalNumberText(magnitude, ctx.locale.decimal));
        break;
      case "denominator":
        out.add(token.text);
        break;
      case "exponent":
        out.add(token.text);
        break;
      default:
        break;
    }
  });
}

/** |v| scaled by the section's % and scaling commas. */
function scaled(layout: NumberLayout, magnitude: number): number {
  return (magnitude * 100 ** layout.percent) / 1000 ** layout.scale;
}

function renderFixed(section: FormatSection, magnitude: number, ctx: FormatContext, out: Output): void {
  const layout = analyse(section);
  const tokens = section.tokens;
  const { int, frac } = roundToPlaces(scaled(layout, magnitude), layout.fracSlots.length);
  const group = layout.grouping ? ctx.locale.group : null;
  const intTexts = integerSlotTexts(layout.intSlots.map((i) => placeholderOf(tokens[i])), int, group);
  const fracTexts = fractionSlotTexts(layout.fracSlots.map((i) => placeholderOf(tokens[i])), frac);
  const intAt = new Map(layout.intSlots.map((index, k) => [index, intTexts[k]!]));
  const fracAt = new Map(layout.fracSlots.map((index, k) => [index, fracTexts[k]!]));
  emitTokens(section, out, ctx, magnitude, (index) => {
    if (intAt.has(index)) {
      return intAt.get(index);
    }
    if (fracAt.has(index)) {
      return fracAt.get(index);
    }
    if (index === layout.point) {
      return (layout.intSlots.length === 0 ? groupDigits(int, group) : "") + ctx.locale.decimal;
    }
    if (layout.silentCommas.has(index)) {
      return "";
    }
    return undefined;
  });
  if (layout.intSlots.length === 0 && layout.point < 0 && int !== "") {
    out.add(groupDigits(int, group));
  }
}

/** The mantissa and exponent of a scientific section: with n > 1 integer
 * placeholders containing a `#` the exponent is a multiple of n, otherwise
 * the mantissa has one integer digit (none without integer placeholders). */
function scientificParts(layout: NumberLayout, tokens: readonly FormatToken[], x: number): { int: string; frac: string; exponent: number } {
  const n = layout.intSlots.length;
  const places = layout.fracSlots.length;
  if (x === 0) {
    return { int: "", frac: "0".repeat(places), exponent: 0 };
  }
  const step = n > 1 && layout.intSlots.some((i) => placeholderOf(tokens[i]) === "#") ? n : 1;
  const sig = significand(x);
  let exponentOfFirst = sig.exponent;
  let digits = sig.digits;
  for (;;) {
    const shown = n === 0 ? exponentOfFirst + 1 : step > 1 ? Math.floor(exponentOfFirst / step) * step : exponentOfFirst;
    const intLength = exponentOfFirst - shown + 1;
    const padded = digits.padEnd(intLength, "0");
    const rounded = roundDigits(padded, intLength + places);
    if (!rounded.carried) {
      const all = rounded.digits.padEnd(intLength + places, "0");
      return { int: all.slice(0, intLength).replace(/^0+/, ""), frac: all.slice(intLength), exponent: shown };
    }
    // The rounding gained a digit (9.99 → 10.0): renormalise one place up.
    exponentOfFirst += 1;
    digits = "1";
  }
}

function renderScientific(section: FormatSection, magnitude: number, ctx: FormatContext, out: Output): void {
  const layout = analyse(section);
  const tokens = section.tokens;
  const { int, frac, exponent } = scientificParts(layout, tokens, scaled(layout, magnitude));
  const intTexts = integerSlotTexts(layout.intSlots.map((i) => placeholderOf(tokens[i])), int, null);
  const fracTexts = fractionSlotTexts(layout.fracSlots.map((i) => placeholderOf(tokens[i])), frac);
  const intAt = new Map(layout.intSlots.map((index, k) => [index, intTexts[k]!]));
  const fracAt = new Map(layout.fracSlots.map((index, k) => [index, fracTexts[k]!]));
  const expToken = tokens[layout.exponent];
  const expText = expToken?.type === "exponent" ? expToken.text : "E+";
  const zeros = layout.expSlots.filter((i) => placeholderOf(tokens[i]) === "0").length;
  const sign = exponent < 0 ? "-" : expText[1] === "+" ? "+" : "";
  const exponentText = `${expText[0]}${sign}${String(Math.abs(exponent)).padStart(Math.max(1, zeros), "0")}`;
  const expSlots = new Set(layout.expSlots);
  emitTokens(section, out, ctx, magnitude, (index) => {
    if (intAt.has(index)) {
      return intAt.get(index);
    }
    if (fracAt.has(index)) {
      return fracAt.get(index);
    }
    if (index === layout.point) {
      return (layout.intSlots.length === 0 ? int : "") + ctx.locale.decimal;
    }
    if (index === layout.exponent) {
      return exponentText;
    }
    if (expSlots.has(index) || layout.silentCommas.has(index)) {
      return "";
    }
    return undefined;
  });
}

function abs(value: bigint): bigint {
  return value < 0n ? -value : value;
}

/** The best rational approximation p/q of n/d (exact, as BigInts) with
 * q ≤ maxDen: continued-fraction convergents and semiconvergents; a tie goes
 * to the smaller denominator. */
function bestRational(n: bigint, d: bigint, maxDen: bigint): [bigint, bigint] {
  let [p0, q0, p1, q1] = [0n, 1n, 1n, 0n];
  let [num, den] = [n, d];
  while (den !== 0n) {
    const a = num / den;
    const q2 = q0 + a * q1;
    if (q2 > maxDen) {
      break;
    }
    [p0, q0, p1, q1] = [p1, q1, p0 + a * p1, q2];
    [num, den] = [den, num - a * den];
  }
  if (den === 0n) {
    return [p1, q1]; // exact
  }
  const k = (maxDen - q0) / q1;
  const [pa, qa] = [p0 + k * p1, q0 + k * q1]; // the best semiconvergent
  const [pb, qb] = [p1, q1]; // the last convergent
  // |pa/qa − n/d| vs |pb/qb − n/d|, cross-multiplied.
  const errorA = abs(pa * d - n * qa) * qb;
  const errorB = abs(pb * d - n * qb) * qa;
  if (errorA < errorB || (errorA === errorB && qa < qb)) {
    return [pa, qa];
  }
  return [pb, qb];
}

function renderFraction(section: FormatSection, magnitude: number, ctx: FormatContext, out: Output): void {
  const layout = analyse(section);
  const tokens = section.tokens;
  const parts = decimalParts(scaled(layout, magnitude));
  const scale = 10n ** BigInt(parts.frac.length);
  const integerPart = layout.intSlots.length > 0;
  let whole = integerPart ? BigInt(parts.int || "0") : 0n;
  const fractionNumerator = integerPart ? BigInt(parts.frac || "0") : BigInt(parts.int || "0") * scale + BigInt(parts.frac || "0");
  let numerator: bigint;
  let denominator: bigint;
  if (layout.denominator !== undefined) {
    denominator = BigInt(layout.denominator);
    // round half away from zero: floor((2·n·D + scale) / (2·scale))
    numerator = (2n * fractionNumerator * denominator + scale) / (2n * scale);
  } else {
    const maxDen = 10n ** BigInt(Math.max(1, layout.denSlots.length)) - 1n;
    [numerator, denominator] = bestRational(fractionNumerator, scale, maxDen);
  }
  if (integerPart && numerator !== 0n && numerator === denominator) {
    whole += 1n;
    numerator = 0n;
  }
  const blank = integerPart && numerator === 0n;
  const group = layout.grouping ? ctx.locale.group : null;
  const wholeDigits = whole === 0n ? (blank ? "0" : "") : whole.toString();
  const intTexts = integerSlotTexts(layout.intSlots.map((i) => placeholderOf(tokens[i])), wholeDigits, group);
  const intAt = new Map(layout.intSlots.map((index, k) => [index, intTexts[k]!]));
  const numPlaceholders = layout.numSlots.map((i) => placeholderOf(tokens[i]));
  const numDigits = numerator.toString();
  const numText = numPlaceholders.slice(0, Math.max(0, numPlaceholders.length - numDigits.length)).map(pad).join("") + numDigits;
  const denPlaceholders = layout.denSlots.map((i) => placeholderOf(tokens[i]));
  const denDigits = denominator.toString();
  const denText =
    layout.denominator !== undefined
      ? denDigits
      : denDigits + denPlaceholders.slice(denDigits.length).map(pad).join("");
  const numSlots = new Set(layout.numSlots);
  const denSlots = new Set(layout.denSlots);
  const firstNum = layout.numSlots[0];
  const firstDen = layout.denominator !== undefined ? layout.slash + 1 : layout.denSlots[0];
  const fractionWidth = layout.numSlots.length + 1 + (layout.denominator !== undefined ? denDigits.length : layout.denSlots.length);
  emitTokens(section, out, ctx, magnitude, (index) => {
    if (intAt.has(index)) {
      return intAt.get(index);
    }
    if (blank && (numSlots.has(index) || denSlots.has(index) || index === layout.slash || (layout.denominator !== undefined && index === layout.slash + 1))) {
      return index === firstNum ? " ".repeat(fractionWidth) : "";
    }
    if (numSlots.has(index)) {
      return index === firstNum ? numText : "";
    }
    if (denSlots.has(index) || (layout.denominator !== undefined && index === layout.slash + 1)) {
      return index === firstDen ? denText : "";
    }
    if (index === layout.slash) {
      return "/";
    }
    if (layout.silentCommas.has(index)) {
      return "";
    }
    return undefined;
  });
}

/** A number section (any kind but date) for |v|. */
function renderNumberSection(section: FormatSection, magnitude: number, ctx: FormatContext, out: Output): void {
  switch (section.kind) {
    case "scientific":
      renderScientific(section, magnitude, ctx, out);
      return;
    case "fraction":
      renderFraction(section, magnitude, ctx, out);
      return;
    case "number":
      renderFixed(section, magnitude, ctx, out);
      return;
    case "text":
      // A text section reached by a number shows the number as General.
      emitTokens(section, out, ctx, magnitude, (_index, token) =>
        token.type === "text" ? generalNumberText(magnitude, ctx.locale.decimal) : undefined,
      );
      return;
    default:
      emitTokens(section, out, ctx, magnitude, () => undefined); // general, literal
  }
}

// --- dates and times ---------------------------------------------------------------

const systemSections = new Map<string, FormatSection>();

/** The first section of a locale's own code (longDate/longTime). */
function systemSection(code: string): FormatSection {
  let section = systemSections.get(code);
  if (!section) {
    section = parseFormatCode(code).sections[0]!;
    systemSections.set(code, section);
  }
  return section;
}

function mod(a: number, n: number): number {
  return ((a % n) + n) % n;
}

/** The ISO 8601 week number of a calendar date. */
function isoWeek(year: number, month: number, day: number): number {
  const days = daysFromCivil(year, month, day);
  const weekday = mod(days + 3, 7) + 1; // Monday 1 … Sunday 7 (1970-01-01 was a Thursday)
  const thursday = days + (4 - weekday);
  const january1 = daysFromCivil(civilFromDays(thursday).year, 1, 1);
  return Math.floor((thursday - january1) / 7) + 1;
}

function pad2(value: number): string {
  return String(value).padStart(2, "0");
}

/** The calendar date of a whole day count, or null when it is out of range
 * (→ hash). Excel's calendar: day 0 is 1900-01-00 and the weekday counts
 * from serial 1 = Sunday; proleptic: the real date and its weekday. */
function calendarDay(
  days: number,
  ctx: FormatContext,
): { year: number; month: number; day: number; weekday: number } | null {
  const excel = (ctx.dateQuirks ?? "excel") === "excel";
  let date: { year: number; month: number; day: number };
  if (excel && ctx.dateSystem === "1900" && days === 0) {
    date = { year: 1900, month: 1, day: 0 };
  } else {
    date = serialToParts(days, ctx.dateSystem);
  }
  if (date.year > 9999 || date.year < 1) {
    return null;
  }
  let weekday: number;
  if (ctx.dateSystem === "1904") {
    weekday = mod(days + 5, 7);
  } else if (excel || days >= 60) {
    weekday = mod(days + 6, 7);
  } else {
    weekday = mod(days, 7); // proleptic: 1899-12-31 (day 0) was a Sunday
  }
  return { ...date, weekday };
}

/** A date/time section for serial v, or null (→ hash). */
function renderDate(section: FormatSection, v: number, ctx: FormatContext, out: Output): boolean {
  const tokens =
    section.lcid === LCID_SYSTEM_LONG_DATE
      ? systemSection(ctx.locale.longDate).tokens
      : section.lcid === LCID_SYSTEM_LONG_TIME
        ? systemSection(ctx.locale.longTime).tokens
        : section.tokens;
  const excel = (ctx.dateQuirks ?? "excel") === "excel";
  const hasElapsed = tokens.some((token) => token.type === "elapsed");
  if (v < 0 && (excel || hasElapsed)) {
    return false;
  }
  const k = Math.min(3, Math.max(0, ...tokens.map((token) => (token.type === "subsecond" ? token.digits : 0))));
  const perSecond = 10 ** k;
  const scaledTotal = v * SECONDS_PER_DAY * perSecond;
  const total = Math.sign(scaledTotal) * Math.round(Math.abs(scaledTotal));
  const perDay = SECONDS_PER_DAY * perSecond;
  const days = Math.floor(total / perDay);
  const remainder = total - days * perDay;
  const sub = remainder % perSecond;
  const secondOfDay = (remainder - sub) / perSecond;
  const hour = Math.floor(secondOfDay / 3600);
  const minute = Math.floor(secondOfDay / 60) % 60;
  const second = secondOfDay % 60;
  const totalSeconds = Math.floor(total / perSecond);
  const showsDate = tokens.some(
    (token) => token.type === "date" && token.part !== "hour" && token.part !== "minute" && token.part !== "second",
  );
  let date: ReturnType<typeof calendarDay> = null;
  if (showsDate || !hasElapsed) {
    date = calendarDay(days, ctx);
    if (date === null) {
      return false;
    }
  }
  const namesTag = (section.lcid !== undefined ? localeTagOfLcid(section.lcid) : undefined) ?? ctx.locale.tag;
  const names = localeNames(namesTag);
  const twelveHour = tokens.some((token) => token.type === "ampm");
  const genitive = tokens.some((token) => token.type === "date" && token.part === "day" && token.length <= 2);
  const subText = String(sub).padStart(k, "0");
  for (const token of tokens) {
    switch (token.type) {
      case "date": {
        const d = date!;
        switch (token.part) {
          case "year":
            out.add(token.length <= 2 ? pad2(d.year % 100) : String(d.year).padStart(4, "0"));
            break;
          case "month":
            if (token.length === 1) {
              out.add(String(d.month));
            } else if (token.length === 2) {
              out.add(pad2(d.month));
            } else if (token.length === 3) {
              out.add(names.monthsShort[d.month - 1]!);
            } else if (token.length === 5) {
              out.add(Array.from(names.monthsLong[d.month - 1]!)[0]!);
            } else {
              out.add((genitive ? names.monthsGenitive : names.monthsLong)[d.month - 1]!);
            }
            break;
          case "day":
            if (token.length <= 2) {
              out.add(token.length === 1 ? String(d.day) : pad2(d.day));
            } else {
              out.add((token.length === 3 ? names.daysShort : names.daysLong)[d.weekday]!);
            }
            break;
          case "quarter":
            out.add(`Q${Math.floor((d.month - 1) / 3) + 1}`);
            break;
          case "week":
            out.add(String(isoWeek(d.year, d.month, d.day)));
            break;
          case "hour": {
            const shown = twelveHour ? hour % 12 || 12 : hour;
            out.add(token.length === 1 ? String(shown) : pad2(shown));
            break;
          }
          case "minute":
            out.add(token.length === 1 ? String(minute) : pad2(minute));
            break;
          case "second":
            out.add(token.length === 1 ? String(second) : pad2(second));
            break;
        }
        break;
      }
      case "elapsed": {
        const value = token.unit === "h" ? Math.floor(totalSeconds / 3600) : token.unit === "m" ? Math.floor(totalSeconds / 60) : totalSeconds;
        out.add(String(value).padStart(token.length, "0"));
        break;
      }
      case "ampm":
        out.add(hour < 12 ? token.am : token.pm);
        break;
      case "subsecond":
        out.add(ctx.locale.decimal + subText.slice(0, token.digits));
        break;
      case "literal":
        out.add(token.text);
        break;
      case "space":
        out.add(" ");
        break;
      case "fill":
        out.addFill(token.char);
        break;
      case "currency":
        out.add(token.symbol);
        break;
      default:
        break;
    }
  }
  return true;
}

// --- text ---------------------------------------------------------------------------

/** A string through a text section: @ is the string. */
function renderText(section: FormatSection, text: string, out: Output): void {
  for (const token of section.tokens) {
    switch (token.type) {
      case "text":
        out.add(text);
        break;
      case "literal":
        out.add(token.text);
        break;
      case "space":
        out.add(" ");
        break;
      case "fill":
        out.addFill(token.char);
        break;
      case "currency":
        out.add(token.symbol);
        break;
      case "percent":
        out.add("%");
        break;
      default:
        break;
    }
  }
}

// --- the entry points --------------------------------------------------------------

/** The scientific General candidates fitting: mantissa decimals 5 → 0. */
function generalFit(n: number, ctx: FormatContext, full: string): FormattedValue {
  const decimal = ctx.locale.decimal;
  if (fits(full, undefined, ctx)) {
    return { text: full, alignHint: "right" };
  }
  const sig = significand(n);
  const magnitude = Math.abs(n);
  if (!/E/.test(full) && magnitude >= 1e-9 && magnitude < 1e11) {
    // (1) Fixed notation: drop decimals one at a time while the integer part
    // fits — from one fewer than the full text shows.
    const shown = full.split(decimal)[1]?.length ?? 0;
    for (let decimals = Math.min(shown, fixedDecimals(sig)) - 1; decimals >= 0; decimals -= 1) {
      const candidate = fixedCandidate(sig, decimals, decimal);
      if (candidate === null) {
        break;
      }
      if (fits(candidate, undefined, ctx)) {
        return { text: candidate, alignHint: "right" };
      }
    }
  }
  // (2) Scientific, mantissa decimals 5 down to 0.
  for (let decimals = 5; decimals >= 0; decimals -= 1) {
    const candidate = scientificCandidate(sig, decimals, decimal);
    if (fits(candidate, undefined, ctx)) {
      return { text: candidate, alignHint: "right" };
    }
  }
  // (3) The # fill.
  return hashed(ctx);
}

/** A value in the General format (SPEC-116's generalText, with the locale's
 * decimal separator and — given `widthPx` — Excel's width fitting: fewer
 * decimals, then scientific, then `#`). */
export function formatGeneral(value: CellValue, ctx: FormatContext): FormattedValue {
  if (value === null) {
    return { text: "" };
  }
  if (typeof value === "string") {
    return { text: value, alignHint: "left" };
  }
  if (typeof value === "boolean") {
    return { text: value ? "TRUE" : "FALSE", alignHint: "center" };
  }
  if (isCellError(value)) {
    return { text: value.error, alignHint: "center" };
  }
  const full = generalNumberText(value, ctx.locale.decimal);
  if (ctx.widthPx === undefined || !Number.isFinite(value) || value === 0) {
    return fits(full, undefined, ctx) ? { text: full, alignHint: "right" } : hashed(ctx);
  }
  return generalFit(value, ctx, full);
}

function formatNumber(n: number, parsed: ParsedFormat, ctx: FormatContext): FormattedValue {
  const numeric = numericSections(parsed);
  if (numeric.length === 0 || !Number.isFinite(n)) {
    return formatGeneral(n, ctx); // only a text section: the number shows as General
  }
  const choice = chooseSection(numeric, n);
  if (choice === null) {
    return hashed(ctx);
  }
  const { section, autoMinus } = choice;
  const color = sectionColor(section, ctx);
  if (numeric.length === 1 && section.kind === "general" && section.tokens.length === 1) {
    // `[Red]General`: General, fitted, in the section's colour.
    const general = formatGeneral(n, ctx);
    return color === undefined ? general : { ...general, color };
  }
  const out = new Output();
  if (isDateSection(section)) {
    if (!renderDate(section, n, ctx, out)) {
      return hashed(ctx, color);
    }
  } else {
    renderNumberSection(section, Math.abs(n), ctx, out);
    if (autoMinus) {
      out.prepend("-");
    }
  }
  if (!fits(out.text, out.fill, ctx)) {
    return hashed(ctx, color);
  }
  return finish(out, "right", color);
}

/** A cell value through a format code (see the module header). Never
 * throws: an absent, invalid or unsupported code renders as General. */
export function formatValue(value: CellValue, code: string | undefined, ctx: FormatContext): FormattedValue {
  if (value === null) {
    return { text: "" };
  }
  if (typeof value === "boolean") {
    return { text: value ? "TRUE" : "FALSE", alignHint: "center" };
  }
  if (isCellError(value)) {
    return { text: value.error, alignHint: "center" };
  }
  const parsed = parsedCode(code);
  if (parsed === null) {
    return formatGeneral(value, ctx);
  }
  try {
    if (typeof value === "string") {
      const textIndex = textSectionIndex(parsed.sections);
      if (textIndex < 0) {
        return { text: value, alignHint: "left" };
      }
      const section = parsed.sections[textIndex]!;
      const out = new Output();
      renderText(section, value, out);
      return finish(out, "left", sectionColor(section, ctx));
    }
    return formatNumber(value, parsed, ctx);
  } catch {
    // A code the grammar accepts but no rendering rule anticipated: General,
    // as for an invalid one — a format never takes a cell down.
    return formatGeneral(value, ctx);
  }
}

/** A formatted value as plain text: the text with one fill character
 * inserted (CSV, the clipboard, the library preview). */
export function plainText(fv: FormattedValue): string {
  if (!fv.fill) {
    return fv.text;
  }
  return fv.text.slice(0, fv.fill.index) + fv.fill.char + fv.text.slice(fv.fill.index);
}

/** Renders a serial through one of the locale's own codes (null: out of
 * the calendar's range). */
function localeDateText(serial: number, code: string, ctx: FormatContext): string | null {
  const out = new Output();
  return renderDate(systemSection(code), serial, { ...ctx, widthPx: undefined }, out) ? out.text : null;
}

/** What the formula bar and the cell editor show for a cell (SPEC-118
 * "editText"): dates as the locale's short date (+ long time when the
 * serial has a time part), time-only values as the long time, percentages
 * ×100 with "%", other numbers in plain notation; TRUE/FALSE, error codes,
 * and text (with a leading ' when the style is quote-prefixed). Never reads
 * a formula. */
export function editText(cell: Pick<Cell, "value">, ctx: EditTextContext): string {
  const value = cell.value;
  if (value === null) {
    return "";
  }
  if (typeof value === "string") {
    return ctx.style?.quotePrefix ? `'${value}` : value;
  }
  if (typeof value === "boolean") {
    return value ? "TRUE" : "FALSE";
  }
  if (isCellError(value)) {
    return value.error;
  }
  const decimal = ctx.locale.decimal;
  const parsed = parsedCode(formatCodeOf(ctx.style ?? {}, ctx.locale));
  const section = parsed === null ? null : sectionForNumber(parsed, value);
  if (section !== null && Number.isFinite(value) && isDateSection(section)) {
    const tokens =
      section.lcid === LCID_SYSTEM_LONG_DATE
        ? systemSection(ctx.locale.longDate).tokens
        : section.lcid === LCID_SYSTEM_LONG_TIME
          ? systemSection(ctx.locale.longTime).tokens
          : section.tokens;
    const showsDate = tokens.some(
      (token) => token.type === "date" && token.part !== "hour" && token.part !== "minute" && token.part !== "second",
    );
    const timeOfDay = mod(Math.round(value * SECONDS_PER_DAY), SECONDS_PER_DAY);
    const date = localeDateText(value, ctx.locale.shortDate, ctx);
    const time = localeDateText(value, ctx.locale.longTime, ctx);
    if (date !== null && time !== null) {
      if (showsDate) {
        return timeOfDay === 0 ? date : `${date} ${time}`;
      }
      return value < 1 ? time : `${date} ${time}`;
    }
    return plainNotation(value, decimal); // a hash case
  }
  if (parsed !== null && parsed.sections.some((s) => s.kind !== "date" && s.tokens.some((token) => token.type === "percent"))) {
    return `${plainNotation(value * 100, decimal)}%`;
  }
  return plainNotation(value, decimal);
}

/** The sample a locale picker shows: 1,234.56 and 2026-10-08 (serial 46303)
 * as that locale writes them. */
const SAMPLE_NUMBER = 1234.56;
const SAMPLE_DATE = 46303;

/** A locale's picker label — its name and a live sample formatted by
 * formatValue: "English (United States) — 1,234.56 · 10/8/2026". */
export function localeOptionLabel(tag: SpreadsheetLocaleTag): string {
  const locale = SPREADSHEET_LOCALES[tag];
  const ctx: FormatContext = { locale, dateSystem: "1900" };
  const number = formatValue(SAMPLE_NUMBER, "#,##0.00", ctx).text;
  const date = formatValue(SAMPLE_DATE, locale.shortDate, ctx).text;
  return `${SPREADSHEET_LOCALE_NAMES[tag]} — ${number} · ${date}`;
}
