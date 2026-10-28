// Text from a PDF's content streams (SPEC-114 — the normative source for
// `core/pdfText.ts`). It READS content streams and never renders anything:
// the text-showing operators are interpreted only far enough to learn which
// characters a page shows and where lines and words break.
//
// What it can do: decode a font's bytes through its /ToUnicode CMap; for a
// simple font without one, through /Encoding's base (WinAnsi, MacRoman or
// Standard) plus /Differences, glyph names read by the Adobe Glyph List's
// rules; reconstruct a word break where a TJ offset exceeds
// WORD_GAP_THOUSANDTHS or text resumes past a gap, and a line break where
// the text-line matrix moves down; and concatenate pages in order.
//
// What it cannot do, stated once: a Type0/CID font with no /ToUnicode has no
// byte-to-character relation outside the font program, so its text is
// unextractable on that page; a scanned page has no text operators at all;
// and glyphs placed by a renderer's metrics may join or split words
// differently than a viewer shows. Coverage says which of these happened —
// never an empty string that lies.
import {
  parsePdf,
  PdfEncryptedError,
  resolve,
  streamData,
  type PdfDocument,
  type PdfPageRef,
  type PdfValue,
} from "./pdfDocument";
import { COVERAGE_REASONS, type ExtractedText, type TextCoverage } from "./extractedText";

/** A TJ adjustment (thousandths of a text-space unit, negative = rightward)
 * past this much is a word break. The same fraction of an em is the gap
 * past which text resuming further along a line is a new word. */
export const WORD_GAP_THOUSANDTHS = 200;
/** The most pages one PDF contributes; later pages are not read. */
export const MAX_PDF_TEXT_PAGES = 2000;
/** How deeply form XObjects may nest inside one another. */
const MAX_FORM_DEPTH = 8;
/** A simple font's glyphs that no table names may be stray dingbats; past
 * this share of a page's glyphs they count as text lost. */
const UNDECODABLE_TOLERANCE = 0.1;

const PDF_NULL: PdfValue = { t: "null" };

// --- the content-stream lexer -------------------------------------------------

/** One operand of a content-stream operator. */
type Operand =
  | { k: "num"; v: number }
  | { k: "str"; v: Uint8Array }
  | { k: "name"; v: string }
  | { k: "arr"; v: Operand[] }
  | { k: "other" };

function isWhite(byte: number): boolean {
  return byte === 0 || byte === 9 || byte === 10 || byte === 12 || byte === 13 || byte === 32;
}

function isDelim(byte: number): boolean {
  return (
    byte === 0x28 || byte === 0x29 || byte === 0x3c || byte === 0x3e || byte === 0x5b ||
    byte === 0x5d || byte === 0x7b || byte === 0x7d || byte === 0x2f || byte === 0x25
  );
}

function hexValue(byte: number): number {
  if (byte >= 0x30 && byte <= 0x39) return byte - 0x30;
  if (byte >= 0x41 && byte <= 0x46) return byte - 0x37;
  if (byte >= 0x61 && byte <= 0x66) return byte - 0x57;
  return -1;
}

/** Reads a content stream (or a CMap, which shares its syntax), calling
 * `onOp` with each operator and its operands. Inline image data (BI … ID
 * … EI) is skipped whole and reported as a "BI" operator; arrays nest;
 * dictionaries are consumed and passed as opaque operands. Malformed input
 * ends the read rather than throwing. */
function readOps(bytes: Uint8Array, onOp: (op: string, args: Operand[]) => void): void {
  let pos = 0;
  const len = bytes.length;
  const stack: Operand[][] = [[]];
  let dictDepth = 0;
  const push = (operand: Operand): void => {
    stack[stack.length - 1].push(operand);
  };
  while (pos < len) {
    const byte = bytes[pos];
    if (isWhite(byte)) {
      pos++;
    } else if (byte === 0x25) {
      while (pos < len && bytes[pos] !== 10 && bytes[pos] !== 13) pos++;
    } else if (byte === 0x28) {
      const out: number[] = [];
      let depth = 1;
      pos++;
      while (pos < len && depth > 0) {
        const c = bytes[pos++];
        if (c === 0x5c) {
          const e = bytes[pos++];
          if (e === 0x6e) out.push(10);
          else if (e === 0x72) out.push(13);
          else if (e === 0x74) out.push(9);
          else if (e === 0x62) out.push(8);
          else if (e === 0x66) out.push(12);
          else if (e === 13) {
            if (bytes[pos] === 10) pos++;
          } else if (e === 10) {
            // a line continuation
          } else if (e >= 0x30 && e <= 0x37) {
            let value = e - 0x30;
            for (let i = 0; i < 2 && bytes[pos] >= 0x30 && bytes[pos] <= 0x37; i++) value = value * 8 + bytes[pos++] - 0x30;
            out.push(value & 0xff);
          } else if (e !== undefined) out.push(e);
        } else if (c === 0x28) {
          depth++;
          out.push(c);
        } else if (c === 0x29) {
          depth--;
          if (depth > 0) out.push(c);
        } else out.push(c);
      }
      push({ k: "str", v: Uint8Array.from(out) });
    } else if (byte === 0x3c && bytes[pos + 1] === 0x3c) {
      dictDepth++;
      stack.push([]);
      pos += 2;
    } else if (byte === 0x3e && bytes[pos + 1] === 0x3e) {
      if (dictDepth > 0) {
        dictDepth--;
        stack.pop();
        push({ k: "other" });
      }
      pos += 2;
    } else if (byte === 0x3c) {
      const out: number[] = [];
      let high = -1;
      pos++;
      while (pos < len && bytes[pos] !== 0x3e) {
        const value = hexValue(bytes[pos++]);
        if (value < 0) continue;
        if (high < 0) high = value;
        else {
          out.push(high * 16 + value);
          high = -1;
        }
      }
      if (high >= 0) out.push(high * 16);
      pos++;
      push({ k: "str", v: Uint8Array.from(out) });
    } else if (byte === 0x5b) {
      stack.push([]);
      pos++;
    } else if (byte === 0x5d) {
      if (stack.length > 1 + dictDepth) {
        const items = stack.pop() as Operand[];
        push({ k: "arr", v: items });
      }
      pos++;
    } else if (byte === 0x2f) {
      pos++;
      let name = "";
      while (pos < len && !isWhite(bytes[pos]) && !isDelim(bytes[pos])) {
        if (bytes[pos] === 0x23 && hexValue(bytes[pos + 1]) >= 0 && hexValue(bytes[pos + 2]) >= 0) {
          name += String.fromCharCode(hexValue(bytes[pos + 1]) * 16 + hexValue(bytes[pos + 2]));
          pos += 3;
        } else name += String.fromCharCode(bytes[pos++]);
      }
      push({ k: "name", v: name });
    } else if (byte === 0x7b || byte === 0x7d || byte === 0x29 || byte === 0x3e) {
      pos++;
    } else {
      let word = "";
      while (pos < len && !isWhite(bytes[pos]) && !isDelim(bytes[pos])) word += String.fromCharCode(bytes[pos++]);
      if (/^[+-]?(\d+\.?\d*|\.\d+)$/.test(word)) {
        push({ k: "num", v: Number(word) });
      } else if (word === "true" || word === "false" || word === "null") {
        push({ k: "other" });
      } else if (stack.length === 1) {
        if (word === "BI") {
          pos = skipInlineImage(bytes, pos);
          onOp("BI", []);
        } else {
          onOp(word, stack[0]);
        }
        stack[0] = [];
      } else {
        // An operator inside an unclosed array or dictionary: drop the
        // structure and resync on the operator.
        stack.length = 1;
        dictDepth = 0;
        stack[0] = [];
      }
    }
  }
}

/** Skips an inline image's parameters and data: past "ID", one whitespace,
 * then up to an "EI" with whitespace before it and whitespace (or the end)
 * after. Returns the offset after "EI". */
function skipInlineImage(bytes: Uint8Array, from: number): number {
  let pos = from;
  while (pos + 1 < bytes.length && !(bytes[pos] === 0x49 && bytes[pos + 1] === 0x44 && isWhite(bytes[pos - 1]) && isWhite(bytes[pos + 2] ?? 32))) pos++;
  pos += 3;
  while (pos + 1 < bytes.length) {
    if (isWhite(bytes[pos - 1]) && bytes[pos] === 0x45 && bytes[pos + 1] === 0x49 && (pos + 2 >= bytes.length || isWhite(bytes[pos + 2]))) {
      return pos + 2;
    }
    pos++;
  }
  return bytes.length;
}

// --- encodings and glyph names ----------------------------------------------------

/** The Adobe Glyph List names of printable ASCII, 0x20–0x7E. */
const ASCII_NAMES = (
  "space exclam quotedbl numbersign dollar percent ampersand quotesingle parenleft parenright asterisk plus comma " +
  "hyphen period slash zero one two three four five six seven eight nine colon semicolon less equal greater question " +
  "at A B C D E F G H I J K L M N O P Q R S T U V W X Y Z bracketleft backslash bracketright asciicircum underscore " +
  "grave a b c d e f g h i j k l m n o p q r s t u v w x y z braceleft bar braceright asciitilde"
).split(" ");

/** The Adobe Glyph List names of Latin-1, 0xA0–0xFF. */
const LATIN1_NAMES = (
  "nbspace exclamdown cent sterling currency yen brokenbar section dieresis copyright ordfeminine guillemotleft " +
  "logicalnot sfthyphen registered macron degree plusminus twosuperior threesuperior acute mu paragraph " +
  "periodcentered cedilla onesuperior ordmasculine guillemotright onequarter onehalf threequarters questiondown " +
  "Agrave Aacute Acircumflex Atilde Adieresis Aring AE Ccedilla Egrave Eacute Ecircumflex Edieresis Igrave Iacute " +
  "Icircumflex Idieresis Eth Ntilde Ograve Oacute Ocircumflex Otilde Odieresis multiply Oslash Ugrave Uacute " +
  "Ucircumflex Udieresis Yacute Thorn germandbls agrave aacute acircumflex atilde adieresis aring ae ccedilla egrave " +
  "eacute ecircumflex edieresis igrave iacute icircumflex idieresis eth ntilde ograve oacute ocircumflex otilde " +
  "odieresis divide oslash ugrave uacute ucircumflex udieresis yacute thorn ydieresis"
).split(" ");

/** The rest of the names the three standard encodings use, plus the common
 * ligatures, by code point. */
const OTHER_NAMES: Record<string, number> = {
  Euro: 0x20ac, quotesinglbase: 0x201a, florin: 0x0192, quotedblbase: 0x201e, ellipsis: 0x2026, dagger: 0x2020,
  daggerdbl: 0x2021, circumflex: 0x02c6, perthousand: 0x2030, Scaron: 0x0160, guilsinglleft: 0x2039, OE: 0x0152,
  Zcaron: 0x017d, quoteleft: 0x2018, quoteright: 0x2019, quotedblleft: 0x201c, quotedblright: 0x201d,
  bullet: 0x2022, endash: 0x2013, emdash: 0x2014, tilde: 0x02dc, trademark: 0x2122, scaron: 0x0161,
  guilsinglright: 0x203a, oe: 0x0153, zcaron: 0x017e, Ydieresis: 0x0178, fi: 0xfb01, fl: 0xfb02, ff: 0xfb00,
  ffi: 0xfb03, ffl: 0xfb04, fraction: 0x2044, dotlessi: 0x0131, Lslash: 0x0141, lslash: 0x0142, breve: 0x02d8,
  dotaccent: 0x02d9, ring: 0x02da, hungarumlaut: 0x02dd, ogonek: 0x02db, caron: 0x02c7, minus: 0x2212,
  notequal: 0x2260, infinity: 0x221e, lessequal: 0x2264, greaterequal: 0x2265, partialdiff: 0x2202,
  summation: 0x2211, product: 0x220f, pi: 0x03c0, integral: 0x222b, Omega: 0x2126, radical: 0x221a,
  approxequal: 0x2248, Delta: 0x2206, lozenge: 0x25ca, nonbreakingspace: 0x00a0, middot: 0x00b7,
};

const GLYPH_NAMES = new Map<string, string>();
ASCII_NAMES.forEach((name, i) => GLYPH_NAMES.set(name, String.fromCharCode(0x20 + i)));
LATIN1_NAMES.forEach((name, i) => GLYPH_NAMES.set(name, String.fromCharCode(0xa0 + i)));
for (const [name, code] of Object.entries(OTHER_NAMES)) {
  GLYPH_NAMES.set(name, String.fromCodePoint(code));
}

/** A glyph name as text, by the Adobe Glyph List's rules: a known name; a
 * `uniXXXX…` run of UTF-16 code units; a `uXXXX`–`uXXXXXX` code point; a
 * suffixed name (`a.sc`) as its base; a ligature (`f_i`) as its parts.
 * Null for a name no rule reads (an anonymous `g42`). */
export function glyphNameToText(name: string): string | null {
  const known = GLYPH_NAMES.get(name);
  if (known !== undefined) {
    return known;
  }
  const dot = name.indexOf(".");
  if (dot > 0) {
    return glyphNameToText(name.slice(0, dot));
  }
  if (name.includes("_")) {
    const parts = name.split("_").map(glyphNameToText);
    return parts.every((part) => part !== null) ? parts.join("") : null;
  }
  const uni = /^uni((?:[0-9A-F]{4})+)$/.exec(name);
  if (uni) {
    const units = uni[1].match(/.{4}/g) ?? [];
    return String.fromCharCode(...units.map((unit) => parseInt(unit, 16)));
  }
  const u = /^u([0-9A-F]{4,6})$/.exec(name);
  if (u) {
    const code = parseInt(u[1], 16);
    return code <= 0x10ffff && (code < 0xd800 || code > 0xdfff) ? String.fromCodePoint(code) : null;
  }
  return null;
}

/** ASCII under a base encoding: printable 0x20–0x7E as themselves. */
function asciiTable(): Array<string | undefined> {
  const table: Array<string | undefined> = Array.from({ length: 256 }, () => undefined);
  for (let code = 0x20; code <= 0x7e; code++) {
    table[code] = String.fromCharCode(code);
  }
  return table;
}

/** WinAnsiEncoding: Windows code page 1252. 0xA0 is a space and 0xAD a
 * hyphen, and the unused codes past 0x40 are bullets (ISO 32000, D.2). */
function winAnsiTable(): Array<string | undefined> {
  const table = asciiTable();
  const block = ["€", "•", "‚", "ƒ", "„", "…", "†", "‡", "ˆ", "‰", "Š", "‹", "Œ", "•", "Ž", "•",
    "•", "‘", "’", "“", "”", "•", "–", "—", "˜", "™", "š", "›", "œ", "•", "ž", "Ÿ"];
  block.forEach((char, i) => {
    table[0x80 + i] = char;
  });
  table[0x7f] = "•";
  for (let code = 0xa1; code <= 0xff; code++) {
    table[code] = String.fromCharCode(code);
  }
  table[0xa0] = " ";
  table[0xad] = "-";
  return table;
}

/** MacRomanEncoding as ISO 32000 gives it: 0xCA is a space, 0xDB the
 * currency sign, 0xF0 unencoded. */
function macRomanTable(): Array<string | undefined> {
  const table = asciiTable();
  const high = [
    "Ä", "Å", "Ç", "É", "Ñ", "Ö", "Ü", "á", "à", "â", "ä", "ã", "å", "ç", "é", "è",
    "ê", "ë", "í", "ì", "î", "ï", "ñ", "ó", "ò", "ô", "ö", "õ", "ú", "ù", "û", "ü",
    "†", "°", "¢", "£", "§", "•", "¶", "ß", "®", "©", "™", "´", "¨", "≠", "Æ", "Ø",
    "∞", "±", "≤", "≥", "¥", "µ", "∂", "∑", "∏", "π", "∫", "ª", "º", "Ω", "æ", "ø",
    "¿", "¡", "¬", "√", "ƒ", "≈", "∆", "«", "»", "…", " ", "À", "Ã", "Õ", "Œ", "œ",
    "–", "—", "“", "”", "‘", "’", "÷", "◊", "ÿ", "Ÿ", "⁄", "¤", "‹", "›", "ﬁ", "ﬂ",
    "‡", "·", "‚", "„", "‰", "Â", "Ê", "Á", "Ë", "È", "Í", "Î", "Ï", "Ì", "Ó", "Ô",
    "", "Ò", "Ú", "Û", "Ù", "ı", "ˆ", "˜", "¯", "˘", "˙", "˚", "¸", "˝", "˛", "ˇ",
  ];
  high.forEach((char, i) => {
    table[0x80 + i] = char === "" ? undefined : char;
  });
  return table;
}

/** StandardEncoding, Type 1's built-in: ASCII with typographic quotes at
 * 0x27 and 0x60, and its own upper half. */
function standardTable(): Array<string | undefined> {
  const table = asciiTable();
  table[0x27] = "’";
  table[0x60] = "‘";
  const high: Record<number, string> = {
    0xa1: "¡", 0xa2: "¢", 0xa3: "£", 0xa4: "⁄", 0xa5: "¥", 0xa6: "ƒ", 0xa7: "§", 0xa8: "¤", 0xa9: "'",
    0xaa: "“", 0xab: "«", 0xac: "‹", 0xad: "›", 0xae: "ﬁ", 0xaf: "ﬂ", 0xb1: "–", 0xb2: "†", 0xb3: "‡",
    0xb4: "·", 0xb6: "¶", 0xb7: "•", 0xb8: "‚", 0xb9: "„", 0xba: "”", 0xbb: "»", 0xbc: "…", 0xbd: "‰",
    0xbf: "¿", 0xc1: "`", 0xc2: "´", 0xc3: "ˆ", 0xc4: "˜", 0xc5: "¯", 0xc6: "˘", 0xc7: "˙", 0xc8: "¨",
    0xca: "˚", 0xcb: "¸", 0xcd: "˝", 0xce: "˛", 0xcf: "ˇ", 0xd0: "—", 0xe1: "Æ", 0xe3: "ª", 0xe8: "Ł",
    0xe9: "Ø", 0xea: "Œ", 0xeb: "º", 0xf1: "æ", 0xf5: "ı", 0xf8: "ł", 0xf9: "ø", 0xfa: "œ", 0xfb: "ß",
  };
  for (const [code, char] of Object.entries(high)) {
    table[Number(code)] = char;
  }
  return table;
}

/** The three base encodings by their /Encoding names. */
export const BASE_ENCODINGS: Record<"WinAnsiEncoding" | "MacRomanEncoding" | "StandardEncoding", ReadonlyArray<string | undefined>> = {
  WinAnsiEncoding: winAnsiTable(),
  MacRomanEncoding: macRomanTable(),
  StandardEncoding: standardTable(),
};

function baseEncoding(name: string | undefined): ReadonlyArray<string | undefined> {
  return name === "WinAnsiEncoding" || name === "MacRomanEncoding" || name === "StandardEncoding"
    ? BASE_ENCODINGS[name]
    : BASE_ENCODINGS.StandardEncoding;
}

// --- /ToUnicode -----------------------------------------------------------------------

/** A parsed /ToUnicode CMap: its code lengths and code → text. */
export interface ToUnicodeMap {
  codespaces: Array<{ bytes: number; lo: number; hi: number }>;
  map: Map<number, string>;
}

/** The most entries one bfrange may add — a hostile range stays bounded. */
const MAX_RANGE_ENTRIES = 0x10000;

function bytesToNumber(bytes: Uint8Array): number {
  let value = 0;
  for (const byte of bytes) {
    value = value * 256 + byte;
  }
  return value;
}

function utf16be(bytes: Uint8Array): string {
  let text = "";
  for (let i = 0; i + 1 < bytes.length; i += 2) {
    text += String.fromCharCode((bytes[i] << 8) | bytes[i + 1]);
  }
  return text;
}

/** Parses a /ToUnicode CMap: codespace ranges, bfchar and both bfrange
 * forms (an incrementing destination, or an array of them). A destination
 * written as a glyph name is read by the glyph-name rules. */
export function parseToUnicode(bytes: Uint8Array): ToUnicodeMap {
  const result: ToUnicodeMap = { codespaces: [], map: new Map() };
  const dest = (operand: Operand | undefined): string | null =>
    operand?.k === "str" ? utf16be(operand.v) : operand?.k === "name" ? glyphNameToText(operand.v) : null;
  readOps(bytes, (op, args) => {
    if (op === "endcodespacerange") {
      for (let i = 0; i + 1 < args.length; i += 2) {
        const lo = args[i];
        const hi = args[i + 1];
        if (lo.k === "str" && hi.k === "str" && lo.v.length > 0) {
          result.codespaces.push({ bytes: lo.v.length, lo: bytesToNumber(lo.v), hi: bytesToNumber(hi.v) });
        }
      }
    } else if (op === "endbfchar") {
      for (let i = 0; i + 1 < args.length; i += 2) {
        const src = args[i];
        const text = dest(args[i + 1]);
        if (src.k === "str" && text !== null) {
          result.map.set(bytesToNumber(src.v), text);
        }
      }
    } else if (op === "endbfrange") {
      for (let i = 0; i + 2 < args.length; i += 3) {
        const lo = args[i];
        const hi = args[i + 1];
        const target = args[i + 2];
        if (lo.k !== "str" || hi.k !== "str") {
          continue;
        }
        const from = bytesToNumber(lo.v);
        const to = Math.min(bytesToNumber(hi.v), from + MAX_RANGE_ENTRIES - 1);
        if (target.k === "arr") {
          target.v.forEach((entry, offset) => {
            const text = dest(entry);
            if (text !== null && from + offset <= to) {
              result.map.set(from + offset, text);
            }
          });
        } else if (target.k === "str" && target.v.length >= 2) {
          const base = utf16be(target.v);
          const head = base.slice(0, -1);
          const last = base.charCodeAt(base.length - 1);
          for (let code = from; code <= to; code++) {
            result.map.set(code, head + String.fromCharCode(last + (code - from)));
          }
        }
      }
    }
  });
  return result;
}

// --- fonts ------------------------------------------------------------------------------

/** One shown code: its text (null when the font cannot say) and advance. */
interface ShownCode {
  text: string | null;
  /** Advance in glyph space thousandths (simple fonts' /Widths, CID /W). */
  width: number;
  /** Code 32 in a single-byte font — the one word spacing applies to. */
  space: boolean;
}

interface FontInfo {
  /** False for a Type0 font with no /ToUnicode and no Unicode CMap: no
   * byte-to-character relation exists outside the font program. */
  extractable: boolean;
  codes(bytes: Uint8Array): ShownCode[];
  /** Glyph space → text space for widths (FontMatrix a; 1/1000 by default). */
  widthScale: number;
}

function nameOf(doc: PdfDocument, value: PdfValue | undefined): string | undefined {
  const resolved = value === undefined ? PDF_NULL : resolve(doc, value);
  return resolved.t === "name" ? resolved.v : undefined;
}

function numberOf(doc: PdfDocument, value: PdfValue | undefined, fallback: number): number {
  const resolved = value === undefined ? PDF_NULL : resolve(doc, value);
  return resolved.t === "num" ? resolved.v : fallback;
}

function dictOf(doc: PdfDocument, value: PdfValue | undefined): Map<string, PdfValue> | null {
  const resolved = value === undefined ? PDF_NULL : resolve(doc, value);
  return resolved.t === "dict" ? resolved.v : resolved.t === "stream" ? resolved.dict : null;
}

function arrayOf(doc: PdfDocument, value: PdfValue | undefined): PdfValue[] {
  const resolved = value === undefined ? PDF_NULL : resolve(doc, value);
  return resolved.t === "array" ? resolved.v : [];
}

async function toUnicodeOf(doc: PdfDocument, font: Map<string, PdfValue>): Promise<ToUnicodeMap | null> {
  const value = font.get("ToUnicode");
  const stream = value === undefined ? PDF_NULL : resolve(doc, value);
  if (stream.t !== "stream") {
    return null;
  }
  try {
    return parseToUnicode(await streamData(doc, stream));
  } catch {
    return null;
  }
}

/** The CID widths of a descendant font: /DW and the /W array's two forms. */
function cidWidths(doc: PdfDocument, cidFont: Map<string, PdfValue> | null): (cid: number) => number {
  const defaultWidth = cidFont ? numberOf(doc, cidFont.get("DW"), 1000) : 1000;
  const widths = new Map<number, number>();
  const list = cidFont ? arrayOf(doc, cidFont.get("W")) : [];
  for (let i = 0; i < list.length; ) {
    const first = resolve(doc, list[i]);
    const next = list[i + 1] === undefined ? PDF_NULL : resolve(doc, list[i + 1]);
    if (first.t !== "num") {
      i++;
    } else if (next.t === "array") {
      next.v.forEach((width, offset) => {
        const w = resolve(doc, width);
        if (w.t === "num" && offset < MAX_RANGE_ENTRIES) widths.set(first.v + offset, w.v);
      });
      i += 2;
    } else {
      const last = next.t === "num" ? next.v : first.v;
      const w = list[i + 2] === undefined ? PDF_NULL : resolve(doc, list[i + 2]);
      for (let cid = first.v; cid <= Math.min(last, first.v + MAX_RANGE_ENTRIES); cid++) {
        if (w.t === "num") widths.set(cid, w.v);
      }
      i += 3;
    }
  }
  return (cid) => widths.get(cid) ?? defaultWidth;
}

/** Splits a Type0 string into codes by the CMap's codespace ranges (the
 * shortest range a prefix falls in), two bytes by default (Identity-H). */
function splitCodes(bytes: Uint8Array, codespaces: ToUnicodeMap["codespaces"]): number[][] {
  const out: number[][] = [];
  const lengths = [...new Set(codespaces.map((space) => space.bytes))].toSorted((a, b) => a - b);
  for (let i = 0; i < bytes.length; ) {
    let length = 0;
    for (const candidate of lengths) {
      if (i + candidate > bytes.length) break;
      const value = bytesToNumber(bytes.subarray(i, i + candidate));
      if (codespaces.some((space) => space.bytes === candidate && value >= space.lo && value <= space.hi)) {
        length = candidate;
        break;
      }
    }
    if (length === 0) {
      length = lengths.length > 0 ? Math.min(lengths[0], bytes.length - i) : Math.min(2, bytes.length - i);
    }
    out.push([bytesToNumber(bytes.subarray(i, i + length)), length]);
    i += length;
  }
  return out;
}

async function loadFont(doc: PdfDocument, font: Map<string, PdfValue>): Promise<FontInfo> {
  const subtype = nameOf(doc, font.get("Subtype"));
  const toUnicode = await toUnicodeOf(doc, font);
  if (subtype === "Type0") {
    const encoding = nameOf(doc, font.get("Encoding")) ?? "";
    const descendant = dictOf(doc, arrayOf(doc, font.get("DescendantFonts"))[0]);
    const width = cidWidths(doc, descendant);
    // A Uni…-UCS2/UTF16 CMap's codes ARE UTF-16: readable without a map.
    const unicodeCodes = /^Uni.*-(UCS2|UTF16)-[HV]$/.test(encoding);
    const codespaces = toUnicode && toUnicode.codespaces.length > 0 ? toUnicode.codespaces : [{ bytes: 2, lo: 0, hi: 0xffff }];
    return {
      extractable: toUnicode !== null || unicodeCodes,
      widthScale: 0.001,
      codes: (bytes) =>
        splitCodes(bytes, codespaces).map(([code, length]) => ({
          text: toUnicode?.map.get(code) ?? (unicodeCodes ? String.fromCharCode(code) : null),
          width: width(code),
          space: length === 1 && code === 32,
        })),
    };
  }
  // A simple font: one byte per code.
  const encodingValue = font.get("Encoding") === undefined ? PDF_NULL : resolve(doc, font.get("Encoding") as PdfValue);
  const table = [...baseEncoding(encodingValue.t === "name" ? encodingValue.v : nameOf(doc, encodingValue.t === "dict" ? encodingValue.v.get("BaseEncoding") : undefined))];
  if (encodingValue.t === "dict") {
    let code = 0;
    for (const entry of arrayOf(doc, encodingValue.v.get("Differences"))) {
      const item = resolve(doc, entry);
      if (item.t === "num") {
        code = item.v;
      } else if (item.t === "name") {
        if (code >= 0 && code < 256) table[code] = glyphNameToText(item.v) ?? undefined;
        code++;
      }
    }
  }
  const firstChar = numberOf(doc, font.get("FirstChar"), 0);
  const widths = arrayOf(doc, font.get("Widths")).map((value) => {
    const resolved = resolve(doc, value);
    return resolved.t === "num" ? resolved.v : 0;
  });
  const matrix = arrayOf(doc, font.get("FontMatrix")).map((value) => {
    const resolved = resolve(doc, value);
    return resolved.t === "num" ? resolved.v : 0;
  });
  const widthScale = subtype === "Type3" && matrix.length === 6 ? matrix[0] : 0.001;
  // No /Widths (a standard 14 font): half an em is the working guess.
  const fallbackWidth = subtype === "Type3" ? 0.5 / Math.max(widthScale, 1e-6) : 500;
  return {
    extractable: true,
    widthScale,
    codes: (bytes) =>
      Array.from(bytes, (code) => ({
        text: toUnicode?.map.get(code) ?? table[code] ?? null,
        width: widths[code - firstChar] ?? fallbackWidth,
        space: code === 32,
      })),
  };
}

// --- interpretation -----------------------------------------------------------------------

type Matrix = [number, number, number, number, number, number];

function multiply(m: Matrix, n: Matrix): Matrix {
  return [
    m[0] * n[0] + m[1] * n[2],
    m[0] * n[1] + m[1] * n[3],
    m[2] * n[0] + m[3] * n[2],
    m[2] * n[1] + m[3] * n[3],
    m[4] * n[0] + m[5] * n[2] + n[4],
    m[4] * n[1] + m[5] * n[3] + n[5],
  ];
}

/** `text` without C0 control characters (some maps send glyphs to U+0000);
 * tab, line feed and carriage return stay. */
function withoutControls(text: string): string {
  let out = "";
  for (const char of text) {
    const code = char.charCodeAt(0);
    if (code >= 0x20 || code === 9 || code === 10 || code === 13) out += char;
  }
  return out;
}

/** The number operand at `i`, 0 when absent or not a number. */
function numberOperand(args: Operand[], i: number): number {
  const operand = args[i];
  return operand?.k === "num" ? operand.v : 0;
}

/** What one page showed. */
interface PageResult {
  text: string;
  /** Characters decoded (whitespace excluded). */
  decoded: number;
  letters: number;
  /** Glyphs a simple font could not name. */
  undecodable: number;
  /** Glyphs shown in fonts with no character map at all. */
  unmapped: number;
  paintsImage: boolean;
  failed: boolean;
}

interface TextParams {
  font: FontInfo | null;
  size: number;
  charSpacing: number;
  wordSpacing: number;
  scale: number;
  leading: number;
}

/** The state of one page's interpretation, shared by its form XObjects. */
class PageReader {
  private parts: string[] = [];
  private tail = "";
  private started = false;
  private lastY = 0;
  private penX = 0;
  result: PageResult = { text: "", decoded: 0, letters: 0, undecodable: 0, unmapped: 0, paintsImage: false, failed: false };
  private fonts = new Map<PdfValue, Promise<FontInfo>>();
  private visitedForms = new Set<PdfValue>();

  constructor(private readonly doc: PdfDocument) {}

  private append(text: string): void {
    this.parts.push(text);
    this.tail = text;
  }

  private space(): void {
    if (this.parts.length > 0 && !/\s$/.test(this.tail)) this.append(" ");
  }

  private newline(): void {
    if (this.parts.length === 0) return;
    while (this.parts.length > 0 && this.tail === " ") {
      this.parts.pop();
      this.tail = this.parts[this.parts.length - 1] ?? "";
    }
    if (!this.tail.endsWith("\n")) this.append("\n");
  }

  text(): string {
    return this.parts.join("").replace(/[ \t]+\n/g, "\n").trim();
  }

  private async font(resources: Map<string, PdfValue> | null, name: string): Promise<FontInfo | null> {
    const fonts = resources ? dictOf(this.doc, resources.get("Font")) : null;
    const ref = fonts?.get(name);
    if (ref === undefined) return null;
    let loaded = this.fonts.get(ref);
    if (!loaded) {
      const dict = dictOf(this.doc, ref);
      loaded = dict ? loadFont(this.doc, dict) : Promise.resolve({ extractable: false, widthScale: 0.001, codes: (bytes) => Array.from(bytes, () => ({ text: null, width: 500, space: false })) });
      this.fonts.set(ref, loaded);
    }
    return loaded;
  }

  /** Interprets one content stream under `resources`. */
  async run(content: Uint8Array, resources: Map<string, PdfValue> | null, ctm: Matrix, depth: number): Promise<void> {
    const ops: Array<[string, Operand[]]> = [];
    readOps(content, (op, args) => {
      ops.push([op, args]);
    });
    let params: TextParams = { font: null, size: 0, charSpacing: 0, wordSpacing: 0, scale: 1, leading: 0 };
    const saved: TextParams[] = [];
    let tm: Matrix = [1, 0, 0, 1, 0, 0];
    let tlm: Matrix = [1, 0, 0, 1, 0, 0];
    const num = numberOperand;
    const moveLine = (tx: number, ty: number): void => {
      tlm = multiply([1, 0, 0, 1, tx, ty], tlm);
      tm = [...tlm];
    };
    const effectiveSize = (): number => Math.abs(params.size) * (Math.hypot(tm[2], tm[3]) || 1);
    const beforeShow = (): void => {
      const x = tm[4];
      const y = tm[5];
      const size = effectiveSize() || 1;
      if (this.started) {
        if (y < this.lastY - size * 0.5 || y > this.lastY + size * 1.2) {
          this.newline();
        } else if (x > this.penX + (size * WORD_GAP_THOUSANDTHS) / 1000 || x < this.penX - size * 2) {
          this.space();
        }
      }
      this.started = true;
    };
    const advance = (tx: number): void => {
      tm[4] += tx * tm[0];
      tm[5] += tx * tm[1];
    };
    const show = (bytes: Uint8Array): void => {
      const font = params.font;
      if (!font) return;
      for (const shown of font.codes(bytes)) {
        if (!font.extractable) {
          this.result.unmapped++;
        } else if (shown.text === null || shown.text === "") {
          this.result.undecodable++;
        } else {
          const text = withoutControls(shown.text);
          if (text !== "") {
            this.append(text);
            for (const char of text) {
              if (!/\s/.test(char)) {
                this.result.decoded++;
                if (/\p{L}/u.test(char)) this.result.letters++;
              }
            }
          }
        }
        const glyph = shown.width * font.widthScale * params.size;
        advance((glyph + params.charSpacing + (shown.space ? params.wordSpacing : 0)) * params.scale);
      }
      this.penX = tm[4];
      this.lastY = tm[5];
    };
    for (const [op, args] of ops) {
      switch (op) {
        case "q":
          saved.push({ ...params });
          break;
        case "Q":
          params = saved.pop() ?? params;
          break;
        case "BT":
          tm = [1, 0, 0, 1, 0, 0];
          tlm = [1, 0, 0, 1, 0, 0];
          break;
        case "Tf": {
          const name = args[0];
          params.font = name?.k === "name" ? await this.font(resources, name.v) : null;
          params.size = num(args, 1);
          break;
        }
        case "Tc":
          params.charSpacing = num(args, 0);
          break;
        case "Tw":
          params.wordSpacing = num(args, 0);
          break;
        case "Tz":
          params.scale = num(args, 0) / 100;
          break;
        case "TL":
          params.leading = num(args, 0);
          break;
        case "Td":
          moveLine(num(args, 0), num(args, 1));
          break;
        case "TD":
          params.leading = -num(args, 1);
          moveLine(num(args, 0), num(args, 1));
          break;
        case "Tm":
          tlm = [num(args, 0), num(args, 1), num(args, 2), num(args, 3), num(args, 4), num(args, 5)];
          tm = [...tlm];
          break;
        case "T*":
          moveLine(0, -params.leading);
          break;
        case "Tj":
          if (args[0]?.k === "str") {
            beforeShow();
            show(args[0].v);
          }
          break;
        case "'":
          moveLine(0, -params.leading);
          if (args[0]?.k === "str") {
            beforeShow();
            show(args[0].v);
          }
          break;
        case '"':
          params.wordSpacing = num(args, 0);
          params.charSpacing = num(args, 1);
          moveLine(0, -params.leading);
          if (args[2]?.k === "str") {
            beforeShow();
            show(args[2].v);
          }
          break;
        case "TJ": {
          const items = args[0]?.k === "arr" ? args[0].v : [];
          beforeShow();
          for (const item of items) {
            if (item.k === "str") {
              show(item.v);
            } else if (item.k === "num") {
              if (-item.v > WORD_GAP_THOUSANDTHS) this.space();
              advance((-item.v / 1000) * params.size * params.scale);
              this.penX = tm[4];
            }
          }
          break;
        }
        case "BI":
          this.result.paintsImage = true;
          break;
        case "Do": {
          const name = args[0];
          if (name?.k === "name") await this.xobject(resources, name.v, ctm, depth);
          break;
        }
        default:
          break;
      }
    }
  }

  /** An XObject: an image marks the page as painting one; a form's own
   * content is read under its own resources, cycles and depth bounded. */
  private async xobject(resources: Map<string, PdfValue> | null, name: string, ctm: Matrix, depth: number): Promise<void> {
    const xobjects = resources ? dictOf(this.doc, resources.get("XObject")) : null;
    const ref = xobjects?.get(name);
    if (ref === undefined) return;
    const stream = resolve(this.doc, ref);
    if (stream.t !== "stream") return;
    const subtype = nameOf(this.doc, stream.dict.get("Subtype"));
    if (subtype === "Image") {
      this.result.paintsImage = true;
    } else if (subtype === "Form" && depth < MAX_FORM_DEPTH && !this.visitedForms.has(ref)) {
      this.visitedForms.add(ref);
      const own = dictOf(this.doc, stream.dict.get("Resources")) ?? resources;
      await this.run(await streamData(this.doc, stream), own, ctm, depth + 1);
      this.visitedForms.delete(ref);
    }
  }
}

/** A page's content streams, concatenated with a line break between. */
async function pageContent(doc: PdfDocument, page: PdfPageRef): Promise<Uint8Array> {
  const pageDict = dictOf(doc, { t: "ref", num: page.ref.num, gen: page.ref.gen });
  const contents = pageDict?.get("Contents");
  if (contents === undefined) return new Uint8Array(0);
  const resolved = resolve(doc, contents);
  const streams = resolved.t === "array" ? resolved.v.map((item) => resolve(doc, item)) : [resolved];
  const parts: Uint8Array[] = [];
  for (const stream of streams) {
    if (stream.t === "stream") {
      parts.push(await streamData(doc, stream), Uint8Array.of(10));
    }
  }
  const total = parts.reduce((sum, part) => sum + part.length, 0);
  const out = new Uint8Array(total);
  let at = 0;
  for (const part of parts) {
    out.set(part, at);
    at += part.length;
  }
  return out;
}

async function readPage(doc: PdfDocument, page: PdfPageRef): Promise<PageResult> {
  const reader = new PageReader(doc);
  try {
    const resources = dictOf(doc, page.inherited.get("Resources"));
    await reader.run(await pageContent(doc, page), resources, [1, 0, 0, 1, 0, 0], 0);
  } catch {
    reader.result.failed = true;
  }
  reader.result.text = reader.text();
  return reader.result;
}

// --- coverage -----------------------------------------------------------------------------

type PageKind = "text" | "partial" | "scanned" | "unmapped" | "blank" | "failed";

/** How one page counts: text it holds, or why it holds none. A page that
 * paints an image and whose only text has no letter (a page number stamped
 * on a scan) is a scanned page — that number is not a text layer. */
function pageKind(page: PageResult): PageKind {
  if (page.failed) return "failed";
  if (page.decoded === 0) {
    return page.unmapped > 0 ? "unmapped" : page.paintsImage ? "scanned" : "blank";
  }
  if (page.letters === 0 && page.paintsImage) return "scanned";
  const shown = page.decoded + page.undecodable + page.unmapped;
  return page.unmapped > 0 || page.undecodable > shown * UNDECODABLE_TOLERANCE ? "partial" : "text";
}

function pagesPhrase(count: number, total: number): string {
  return count === total ? "every page" : `${count} of ${total} pages`;
}

/** The text of a PDF and how much of it that is. `full`: every page that
 * shows text was decoded. `none`: no page yielded any — the reason names
 * the ceiling met (no character map, a scan, an unreadable file). Else
 * `partial`, saying which pages were lost and why. Pages past
 * MAX_PDF_TEXT_PAGES are not read (partial, over the size ceiling). */
export async function extractPdfText(
  input: Uint8Array | PdfDocument,
  options: { maxPages?: number } = {},
): Promise<ExtractedText & { pageCount: number }> {
  let doc: PdfDocument;
  try {
    doc = input instanceof Uint8Array ? await parsePdf(input) : input;
  } catch (error) {
    const reason = error instanceof PdfEncryptedError ? COVERAGE_REASONS.encrypted : COVERAGE_REASONS.unreadable;
    return { text: "", coverage: "none", reason, pageCount: 0 };
  }
  const maxPages = options.maxPages ?? MAX_PDF_TEXT_PAGES;
  const pages = doc.pages.slice(0, maxPages);
  const texts: string[] = [];
  const counts: Record<PageKind, number> = { text: 0, partial: 0, scanned: 0, unmapped: 0, blank: 0, failed: 0 };
  for (const page of pages) {
    const result = await readPage(doc, page);
    const kind = pageKind(result);
    counts[kind]++;
    if ((kind === "text" || kind === "partial") && result.text !== "") {
      texts.push(result.text);
    }
  }
  const total = doc.pages.length;
  const text = texts.join("\n");
  const withText = counts.text + counts.partial;
  const reasons: string[] = [];
  if (counts.unmapped > 0 || counts.partial > 0) {
    reasons.push(`${COVERAGE_REASONS.noCharMap} on ${pagesPhrase(counts.unmapped + counts.partial, total)}`);
  }
  if (counts.scanned > 0) {
    reasons.push(`${COVERAGE_REASONS.scanned} on ${pagesPhrase(counts.scanned, total)}`);
  }
  if (counts.failed > 0) {
    reasons.push(`${COVERAGE_REASONS.unreadable}: ${pagesPhrase(counts.failed, total)}`);
  }
  if (total > pages.length) {
    reasons.push(`${COVERAGE_REASONS.sizeCeiling} — only the first ${maxPages.toLocaleString("en-US")} pages are searched`);
  }
  let coverage: TextCoverage;
  let reason: string | undefined;
  if (withText === 0) {
    coverage = "none";
    reason =
      counts.unmapped > 0
        ? COVERAGE_REASONS.noCharMap
        : counts.scanned > 0
          ? COVERAGE_REASONS.scanned
          : counts.failed > 0
            ? COVERAGE_REASONS.unreadable
            : COVERAGE_REASONS.empty;
  } else if (reasons.length === 0) {
    coverage = "full";
  } else {
    coverage = "partial";
    reason = reasons.join("; ");
  }
  return { text: coverage === "none" ? "" : text, coverage, ...(reason ? { reason } : {}), pageCount: total };
}
