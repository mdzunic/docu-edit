// The shared-string table on save (SPEC-122 "Shared strings", and the
// "Strings" paragraph of "The worksheet writer"). THIS FILE is the normative
// source for the `_xHHHH_` text encoding and for how `xl/sharedStrings.xml`
// grows when a sheet is rebuilt — append-only, because every clean sheet
// part still references the source items by number.
//
// - encodeXlsxText maps every character XML 1.0 cannot carry (the C0
//   controls but tab and line feed, U+FFFE, U+FFFF, a lone surrogate) and CR
//   (which any XML parser would normalise away) to `_xHHHH_` (upper-case
//   hex), and escapes a literal `_xHHHH_` as `_x005F_xHHHH_` — exactly the
//   underscores the reader's decodeXmlEscapes would otherwise consume, so the
//   decoding reads back the text it was given (property-tested) — then
//   XML-escapes `&`, `<` and `>`.
// - SharedStringsWriter, one per save: `indexFor(cell)` is `cell.sstIndex`
//   when present (an unchanged string keeps its item, rich runs and phonetic
//   runs intact); else, for a plain string, the first source `<si>` that is a
//   lone `<t>` with the same text (a lookup built once, on first need, with
//   one scanXml pass); else an earlier append; else a new `<si><t>…</t></si>`.
//   Rich text is appended as `<si><r><rPr>…</rPr><t>…</t></r>…</si>` — the
//   run font in styles.xml's font XML (xlsx/styles.ts) with `rFont` naming
//   the face — and deduplicated among appends. `<t>` gets
//   `xml:space="preserve"` when its text has leading or trailing whitespace
//   or a line break. Items are never removed.
// - The part is rewritten only when something was appended: the source text
//   with the appends spliced before `</sst>`, `uniqueCount` = the item count
//   and `count` removed (it is optional, and correct only with every clean
//   part's references counted). Without a source part the whole part is
//   generated; writer.ts adds its relationship and Override.
import type { DocumentTheme } from "../../core/theme";
import { scanXml } from "../../core/xmlScanner";
import type { Cell, TextRun } from "../model";
import { applySplices, escapeElementText, setAttribute, startTagAt, type Splice } from "../xmlSplice";
import { isSpreadsheetMl, xlsxFontProperties } from "./styles";

const XML_DECLARATION = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`;
const MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

/** `_xHHHH_` escapes decoded to their code unit. `_x005F_` is a literal
 * underscore, so `_x005F_x0041_` reads as `_x0041_` (the match after an
 * escape resumes past it). The exact inverse of encodeXlsxText. */
export function decodeXmlEscapes(text: string): string {
  if (!text.includes("_x")) {
    return text;
  }
  return text.replace(/_x([0-9A-Fa-f]{4})_/g, (_, hex: string) => String.fromCharCode(Number.parseInt(hex, 16)));
}

/** Text that needs the slow path: an encoded character, a surrogate (lone
 * ones are encoded, pairs kept) or something shaped like an escape. */
const NEEDS_ENCODING = /[\x00-\x08\x0B-\x1F\uD800-\uDFFF￾￿]|_x[0-9A-Fa-f]{4}/;

function isHexDigit(code: number): boolean {
  return (code >= 0x30 && code <= 0x39) || (code >= 0x41 && code <= 0x46) || (code >= 0x61 && code <= 0x66);
}

/** Whether the code unit at `i` is written as `_xHHHH_`: a character XML 1.0
 * cannot carry (a lone surrogate included), or CR. */
function isEncodedAt(text: string, i: number): boolean {
  const code = text.charCodeAt(i);
  if (code < 0x20) {
    return code !== 0x09 && code !== 0x0a;
  }
  if (code === 0xfffe || code === 0xffff) {
    return true;
  }
  if (code >= 0xd800 && code <= 0xdbff) {
    const next = text.charCodeAt(i + 1);
    return !(next >= 0xdc00 && next <= 0xdfff);
  }
  if (code >= 0xdc00 && code <= 0xdfff) {
    const previous = i > 0 ? text.charCodeAt(i - 1) : 0;
    return !(previous >= 0xd800 && previous <= 0xdbff);
  }
  return false;
}

/** Whether the underscore at `i` must be escaped: it starts `_xHHHH` and the
 * next character WRITTEN begins with an underscore — a literal one, or an
 * escape — so the decoder would read the six characters plus it as one. */
function startsLiteralEscape(text: string, i: number): boolean {
  if (text.charCodeAt(i + 1) !== 0x78 /* x */ || i + 6 >= text.length) {
    return false;
  }
  for (let j = i + 2; j < i + 6; j += 1) {
    if (!isHexDigit(text.charCodeAt(j))) {
      return false;
    }
  }
  return text.charCodeAt(i + 6) === 0x5f || isEncodedAt(text, i + 6);
}

function escapeCode(code: number): string {
  return `_x${code.toString(16).toUpperCase().padStart(4, "0")}_`;
}

/** A cell string as SpreadsheetML text content (see the module header): the
 * `_xHHHH_` encoding, then `&`, `<` and `>` escaped. */
export function encodeXlsxText(text: string): string {
  if (!NEEDS_ENCODING.test(text)) {
    return escapeElementText(text);
  }
  let out = "";
  for (let i = 0; i < text.length; i += 1) {
    const code = text.charCodeAt(i);
    if (code === 0x5f && startsLiteralEscape(text, i)) {
      out += "_x005F_";
    } else if (isEncodedAt(text, i)) {
      out += escapeCode(code);
    } else {
      out += text[i];
    }
  }
  return escapeElementText(out);
}

/** Whether a `<t>` holding `text` needs `xml:space="preserve"`. */
function needsPreserve(text: string): boolean {
  return /^[ \t\n\r]|[ \t\n\r]$|[\n\r]/.test(text);
}

/** A `<t>` element (with the prefix the part writes its elements with). */
export function xlsxTextElement(text: string, prefix = ""): string {
  return `<${prefix}t${needsPreserve(text) ? ` xml:space="preserve"` : ""}>${encodeXlsxText(text)}</${prefix}t>`;
}

/** What one scan of the source part yields. */
interface SourceTable {
  /** The number of `<si>` items. */
  count: number;
  /** Text → the first item that is a lone `<t>` with that text. */
  plain: Map<string, number>;
  /** The root start tag and close tag, for the rewrite. */
  root: { start: number; end: number; selfClosing: boolean; name: string; closeStart: number } | null;
  /** The prefix the root (and so every appended item) is written with. */
  prefix: string;
}

const EMPTY_TABLE: SourceTable = { count: 0, plain: new Map(), root: null, prefix: "" };

/** Scans a shared-strings part once: item count, plain-item lookup, root
 * offsets. */
function scanSourceTable(xml: string): SourceTable {
  const table: SourceTable = { count: 0, plain: new Map(), root: null, prefix: "" };
  let depth = 0;
  /** Inside a level-2 `<si>`: its text so far, and whether it is still a lone `<t>`. */
  let item: { text: string; children: number; lone: boolean; inT: boolean } | null = null;
  const finishItem = (): void => {
    if (item !== null && item.lone && item.children === 1) {
      const text = decodeXmlEscapes(item.text);
      if (!table.plain.has(text)) {
        table.plain.set(text, table.count);
      }
    }
    table.count += 1;
    item = null;
  };
  scanXml(xml, (event) => {
    if (event.type === "open") {
      const level = depth + 1;
      if (level === 1) {
        table.root = { start: event.start, end: event.end, selfClosing: event.selfClosing, name: event.name, closeStart: -1 };
        table.prefix = event.name.includes(":") ? `${event.name.slice(0, event.name.indexOf(":"))}:` : "";
      } else if (level === 2 && isSpreadsheetMl(event.ns) && event.local === "si") {
        item = { text: "", children: 0, lone: true, inT: false };
        if (event.selfClosing) {
          item.lone = false;
          finishItem();
        }
      } else if (level === 3 && item !== null) {
        item.children += 1;
        if (!isSpreadsheetMl(event.ns) || event.local !== "t") {
          item.lone = false;
        } else {
          item.inT = !event.selfClosing;
        }
      } else if (level > 3 && item !== null) {
        item.lone = false;
      }
      if (!event.selfClosing) {
        depth = level;
      }
      return;
    }
    if (event.type === "close") {
      if (depth === 1 && table.root !== null) {
        table.root.closeStart = event.start;
      } else if (depth === 2 && item !== null) {
        finishItem();
      } else if (depth === 3 && item !== null) {
        item.inT = false;
      }
      depth -= 1;
      return;
    }
    if (item !== null) {
      if (item.inT) {
        item.text += event.text;
      } else if (event.text.trim() !== "") {
        item.lone = false; // mixed content directly in <si>
      }
    }
  });
  return table;
}

/** The workbook's shared-string table for one save (see the module header). */
export class SharedStringsWriter {
  private table: SourceTable | null = null;
  private readonly appended: string[] = [];
  private readonly appendedIndex = new Map<string, number>();

  /** `source` is the retained `xl/sharedStrings.xml` text (undefined: the
   * workbook has none); `theme` resolves a rich run's theme typeface. */
  constructor(
    private readonly source: string | undefined,
    private readonly theme: DocumentTheme | undefined,
  ) {}

  private sourceTable(): SourceTable {
    if (this.table === null) {
      this.table = this.source === undefined ? EMPTY_TABLE : scanSourceTable(this.source);
    }
    return this.table;
  }

  /** The item index a string cell is written with (see the module header).
   * A cell whose value is not a string is indexed by its text. */
  indexFor(cell: Pick<Cell, "value" | "runs" | "sstIndex">): number {
    if (cell.sstIndex !== undefined && this.source !== undefined) {
      return cell.sstIndex;
    }
    const table = this.sourceTable();
    const text = typeof cell.value === "string" ? cell.value : String(cell.value ?? "");
    const rich = cell.runs !== undefined && cell.runs.length > 0;
    if (!rich) {
      const found = table.plain.get(text);
      if (found !== undefined) {
        return found;
      }
    }
    const p = table.prefix;
    const item = rich ? `<${p}si>${cell.runs!.map((run) => this.runXml(run, p)).join("")}</${p}si>` : `<${p}si>${xlsxTextElement(text, p)}</${p}si>`;
    const known = this.appendedIndex.get(item);
    if (known !== undefined) {
      return known;
    }
    const index = table.count + this.appended.length;
    this.appended.push(item);
    this.appendedIndex.set(item, index);
    return index;
  }

  private runXml(run: TextRun, p: string): string {
    const font = run.font !== undefined && Object.keys(run.font).length > 0 ? `<${p}rPr>${xlsxFontProperties(run.font, this.theme, "rFont", p)}</${p}rPr>` : "";
    return `<${p}r>${font}${xlsxTextElement(run.text, p)}</${p}r>`;
  }

  /** Whether an item was appended — the part must then be written. */
  get changed(): boolean {
    return this.appended.length > 0;
  }

  /** The part text when something was appended (undefined: the source part,
   * or its absence, stands). */
  xml(): string | undefined {
    if (this.appended.length === 0) {
      return undefined;
    }
    const table = this.sourceTable();
    const items = this.appended.join("");
    const uniqueCount = String(table.count + this.appended.length);
    const root = table.root;
    if (this.source === undefined || root === null) {
      return `${XML_DECLARATION}<sst xmlns="${MAIN_NS}" uniqueCount="${uniqueCount}">${items}</sst>`;
    }
    const tag = setAttribute(setAttribute(startTagAt(this.source, root.start), "uniqueCount", uniqueCount), "count", null);
    const splices: Splice[] = [];
    if (root.selfClosing) {
      splices.push({ start: root.start, end: root.end, text: `${tag.replace(/\s*\/>$/, ">")}${items}</${root.name}>` });
    } else {
      splices.push({ start: root.start, end: root.end, text: tag });
      splices.push({ start: root.closeStart, end: root.closeStart, text: items });
    }
    return applySplices(this.source, splices);
  }
}
