// Excel's number-format code grammar (SPEC-118 "The grammar") — the lingua
// franca every number format in this repo is stored and shown in: ODS number
// styles map onto it both ways (odsNumberStyles.ts), the built-in ids expand
// to it (builtin.ts) and render.ts renders it. Codes are always in the
// invariant syntax: `.` is the decimal point and `,` the grouping/scaling
// comma whatever the locale; `/` and `:` are literal characters.
//
//   code     := section ( ";" section ){0,3}
//   section  := bracket* body              (brackets are read anywhere in a section)
//   bracket  := "[" color "]" | "[" condition "]" | "[$" symbol? ( "-" lcid )? "]" | "[" elapsed "]"
//   body     := ( placeholder | "." | "," | "%" | exponent | "/" | "@" | "General"
//               | datetoken | ampm | literal )*
//   literal  := '"' … '"' | "\" char | "_" char | "*" char | any other char
//
// - A section with any date/time token is a date section: `0#?` there are
//   literals except the fractional-second zeros after a seconds token.
// - The minute rule: `m`/`mm` is minutes when the nearest preceding date/time
//   token is an hour or the nearest following one is a second; else the month.
// - Errors (FormatCodeError, with the offset of the offending character): an
//   unterminated quote or bracket, a fifth section, an unknown bracket, a
//   section mixing digit placeholders and `@`, an `E` with no exponent
//   digits. Unknown bare letters are lenient literals.
// - Calendars and numeral systems out of scope (`B1`/`B2`, `bb`, the era
//   tokens `e`/`g`, `[DBNum*]`/`[NatNum*]`, LibreOffice's `[~calendar]`, an
//   LCID above 0xFFFF) parse as `{ unsupported: true }` — rendered as General.
// - Inside quotes `\"` is a literal quote: the escape odsNumberStyleToCode
//   writes for a `number:text` holding one.
// Results are cached by code (FORMAT_CACHE_LIMIT entries, cleared whole on
// overflow) — a sheet renders thousands of cells through a handful of codes.
// formatCodeSections is the same scan with each token's source offset, for
// an edit that must keep every other character as written (SPEC-123's
// decimal stepping).

export const FORMAT_CACHE_LIMIT = 1024;

/** A structural error in a format code, at `offset` (UTF-16 index). */
export class FormatCodeError extends Error {
  readonly offset: number;

  constructor(message: string, offset: number) {
    super(`${message} (at offset ${offset})`);
    this.name = "FormatCodeError";
    this.offset = offset;
  }
}

export type ConditionOp = "<" | "<=" | ">" | ">=" | "=" | "<>";

/** The calendar part a date/time token shows. */
export type DatePart = "year" | "month" | "minute" | "day" | "hour" | "second" | "quarter" | "week";

/** One element of a section's body. The shape is internal to parse/render. */
export type FormatToken =
  | { type: "literal"; text: string }
  | { type: "space"; char: string } // _x: one space (as wide as x)
  | { type: "fill"; char: string } // *x: x repeated to fill the cell
  | { type: "digit"; placeholder: "0" | "#" | "?" }
  | { type: "point" }
  | { type: "comma" }
  | { type: "percent" }
  | { type: "exponent"; text: string } // "E+", "E-", "e+", "e-" as written
  | { type: "slash" }
  | { type: "denominator"; value: number; text: string } // the "8" of ?/8
  | { type: "text" } // @
  | { type: "general" }
  | { type: "currency"; symbol: string } // the symbol of a [$sym-lcid]
  | { type: "date"; part: DatePart; length: number }
  | { type: "elapsed"; unit: "h" | "m" | "s"; length: number }
  | { type: "ampm"; am: string; pm: string }
  | { type: "subsecond"; digits: number };

export interface FormatSection {
  condition?: { op: ConditionOp; value: number };
  /** A colour name as written canonically: "Red", "Blue", …, or "Color10". */
  color?: string;
  lcid?: number;
  kind: "number" | "fraction" | "scientific" | "date" | "text" | "general" | "literal";
  tokens: FormatToken[];
}

export interface ParsedFormat {
  sections: FormatSection[]; // 1..4
  /** A calendar or numeral system out of scope: render as General. */
  unsupported?: true;
}

/** The eight colour names, lower-cased → canonical. */
const COLOR_NAMES: ReadonlyMap<string, string> = new Map(
  ["Black", "Blue", "Cyan", "Green", "Magenta", "Red", "White", "Yellow"].map((name) => [name.toLowerCase(), name]),
);

/** The LCID words `[$-x-sysdate]` and `[$-x-systime]` stand for. */
export const LCID_SYSTEM_LONG_DATE = 0xf800;
export const LCID_SYSTEM_LONG_TIME = 0xf400;

const CONDITION = /^(<=|>=|<>|<|>|=)\s*([-+]?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?)$/i;
const CURRENCY = /^\$(.*?)(?:-(x-sysdate|x-systime|[0-9a-f]{1,8}))?$/i;

/** A token with the offset of its first character, before classification. */
interface RawToken {
  token: FormatToken | EraToken;
  offset: number;
}

/** A calendar token out of scope (`e`, `g`, `B1`, `bb`): makes the code
 * unsupported. */
interface EraToken {
  type: "era";
}

interface RawSection {
  start: number;
  tokens: RawToken[];
  condition?: FormatSection["condition"];
  color?: string;
  lcid?: number;
}

/** What a scan of one code collects before the sections are classified. */
interface ScanState {
  unsupported: boolean;
}

function isLetter(ch: string): boolean {
  return /^[a-z]$/i.test(ch);
}

/** The length of the run of `ch` (case-insensitive) starting at `i`. */
function runLength(code: string, i: number, ch: string): number {
  let j = i;
  while (j < code.length && code[j]!.toLowerCase() === ch) {
    j += 1;
  }
  return j - i;
}

/** Reads a bracket's content into the section (or a body token); returns
 * the index after the `]`. */
function readBracket(code: string, open: number, section: RawSection, state: ScanState): number {
  const close = code.indexOf("]", open + 1);
  if (close < 0) {
    throw new FormatCodeError("Unterminated bracket", open);
  }
  const content = code.slice(open + 1, close);
  const lower = content.toLowerCase();
  const colorName = COLOR_NAMES.get(lower);
  const colorIndex = /^color(\d{1,2})$/i.exec(content);
  const condition = CONDITION.exec(content);
  if (colorName !== undefined) {
    section.color = colorName;
  } else if (colorIndex && Number(colorIndex[1]) >= 1 && Number(colorIndex[1]) <= 56) {
    section.color = `Color${Number(colorIndex[1])}`;
  } else if (condition) {
    section.condition = { op: condition[1] as ConditionOp, value: Number(condition[2]) };
  } else if (content.startsWith("$")) {
    const match = CURRENCY.exec(content)!;
    const symbol = match[1] ?? "";
    const lcidText = match[2]?.toLowerCase();
    if (lcidText !== undefined) {
      const lcid =
        lcidText === "x-sysdate" ? LCID_SYSTEM_LONG_DATE : lcidText === "x-systime" ? LCID_SYSTEM_LONG_TIME : Number.parseInt(lcidText, 16);
      if (lcid > 0xffff) {
        state.unsupported = true; // a calendar or numeral-system byte
      }
      section.lcid = lcid & 0xffff;
    }
    if (symbol !== "") {
      section.tokens.push({ token: { type: "currency", symbol }, offset: open });
    }
  } else if (/^(h+|m+|s+)$/i.test(content)) {
    section.tokens.push({ token: { type: "elapsed", unit: lower[0] as "h" | "m" | "s", length: content.length }, offset: open });
  } else if (/^(dbnum\d*|natnum\d*.*)$/i.test(content) || content.startsWith("~")) {
    state.unsupported = true; // a numeral system, or LibreOffice's [~calendar]
  } else {
    throw new FormatCodeError(`Unknown bracket [${content}]`, open);
  }
  return close + 1;
}

/** Reads a quoted literal starting at `open`; returns [text, index after]. */
function readQuoted(code: string, open: number): [string, number] {
  let text = "";
  let i = open + 1;
  while (i < code.length) {
    const ch = code[i]!;
    if (ch === "\\" && code[i + 1] === '"') {
      text += '"';
      i += 2;
      continue;
    }
    if (ch === '"') {
      return [text, i + 1];
    }
    text += ch;
    i += 1;
  }
  throw new FormatCodeError("Unterminated quote", open);
}

/** Reads the letters at `i` — a multi-letter token or a run of one letter;
 * returns the index after it. */
function readLetters(code: string, i: number, section: RawSection): number {
  const rest = code.slice(i);
  const lower = rest.toLowerCase();
  const push = (token: FormatToken | EraToken): void => {
    section.tokens.push({ token, offset: i });
  };
  if (lower.startsWith("general")) {
    push({ type: "general" });
    return i + 7;
  }
  if (lower.startsWith("am/pm")) {
    push({ type: "ampm", am: rest[0] === "A" ? "AM" : "am", pm: rest[3] === "P" ? "PM" : "pm" });
    return i + 5;
  }
  if (lower.startsWith("a/p")) {
    push({ type: "ampm", am: rest[0] === "A" ? "A" : "a", pm: rest[2] === "P" ? "P" : "p" });
    return i + 3;
  }
  const ch = lower[0]!;
  if (ch === "e" && (rest[1] === "+" || rest[1] === "-")) {
    push({ type: "exponent", text: rest.slice(0, 2) });
    return i + 2;
  }
  if (ch === "b" && (rest[1] === "1" || rest[1] === "2")) {
    push({ type: "era" });
    return i + 2;
  }
  const length = runLength(code, i, ch);
  switch (ch) {
    case "y":
      push({ type: "date", part: "year", length });
      return i + length;
    case "m":
      push({ type: "date", part: "month", length }); // or minutes: the minute rule decides
      return i + length;
    case "d":
      push({ type: "date", part: "day", length });
      return i + length;
    case "h":
      push({ type: "date", part: "hour", length });
      return i + length;
    case "s":
      push({ type: "date", part: "second", length });
      return i + length;
    case "q":
      push({ type: "date", part: "quarter", length });
      return i + length;
    case "e":
    case "g":
      push({ type: "era" });
      return i + length;
    case "b":
      if (length === 2 || length === 4) {
        push({ type: "era" }); // Buddhist years
        return i + length;
      }
      break;
    case "w":
      if (length >= 2) {
        push({ type: "date", part: "week", length });
        return i + length;
      }
      break;
    default:
      break;
  }
  // An unknown bare letter is a lenient literal.
  push({ type: "literal", text: code.slice(i, i + length) });
  return i + length;
}

/** Scans one code into raw sections (brackets resolved, body tokenised). */
function scan(code: string, state: ScanState): RawSection[] {
  const sections: RawSection[] = [];
  let section: RawSection = { start: 0, tokens: [] };
  let i = 0;
  while (i < code.length) {
    const ch = code[i]!;
    const push = (token: FormatToken): void => {
      section.tokens.push({ token, offset: i });
    };
    switch (ch) {
      case ";":
        sections.push(section);
        if (sections.length === 4) {
          throw new FormatCodeError("A format code has at most four sections", i);
        }
        section = { start: i + 1, tokens: [] };
        i += 1;
        continue;
      case '"': {
        const [text, next] = readQuoted(code, i);
        push({ type: "literal", text });
        i = next;
        continue;
      }
      case "\\":
        if (i + 1 < code.length) {
          push({ type: "literal", text: code[i + 1]! });
        }
        i += 2;
        continue;
      case "_":
        if (i + 1 < code.length) {
          push({ type: "space", char: code[i + 1]! });
        }
        i += 2;
        continue;
      case "*":
        if (i + 1 < code.length) {
          push({ type: "fill", char: code[i + 1]! });
        }
        i += 2;
        continue;
      case "[":
        i = readBracket(code, i, section, state);
        continue;
      case "0":
      case "#":
      case "?":
        push({ type: "digit", placeholder: ch });
        i += 1;
        continue;
      case ".":
        push({ type: "point" });
        i += 1;
        continue;
      case ",":
        push({ type: "comma" });
        i += 1;
        continue;
      case "%":
        push({ type: "percent" });
        i += 1;
        continue;
      case "@":
        push({ type: "text" });
        i += 1;
        continue;
      case "/": {
        push({ type: "slash" });
        const fixed = /^[1-9][0-9]*/.exec(code.slice(i + 1));
        if (fixed) {
          section.tokens.push({ token: { type: "denominator", value: Number(fixed[0]), text: fixed[0] }, offset: i + 1 });
          i += 1 + fixed[0].length;
        } else {
          i += 1;
        }
        continue;
      }
      default:
        if (isLetter(ch)) {
          i = readLetters(code, i, section);
        } else {
          push({ type: "literal", text: ch });
          i += 1;
        }
        continue;
    }
  }
  sections.push(section);
  return sections;
}

function isDateToken(token: FormatToken): boolean {
  return token.type === "date" || token.type === "elapsed" || token.type === "ampm";
}

function isHour(token: FormatToken): boolean {
  return (token.type === "date" && token.part === "hour") || (token.type === "elapsed" && token.unit === "h");
}

function isSecond(token: FormatToken): boolean {
  return (token.type === "date" && token.part === "second") || (token.type === "elapsed" && token.unit === "s");
}

function isTimeOrDate(token: FormatToken): boolean {
  return token.type === "date" || token.type === "elapsed";
}

/** The minute rule over a date section's tokens: each `m`/`mm` whose nearest
 * preceding date/time token is an hour, or whose nearest following one is a
 * second, becomes a minute token. */
function applyMinuteRule(tokens: FormatToken[]): void {
  tokens.forEach((token, index) => {
    if (token.type !== "date" || token.part !== "month" || token.length > 2) {
      return;
    }
    let previous: FormatToken | undefined;
    for (let j = index - 1; j >= 0; j -= 1) {
      if (isTimeOrDate(tokens[j]!)) {
        previous = tokens[j];
        break;
      }
    }
    let next: FormatToken | undefined;
    for (let j = index + 1; j < tokens.length; j += 1) {
      if (isTimeOrDate(tokens[j]!)) {
        next = tokens[j];
        break;
      }
    }
    if ((previous && isHour(previous)) || (next && isSecond(next))) {
      tokens[index] = { type: "date", part: "minute", length: token.length };
    }
  });
}

/** A literal standing for a token that has no meaning in a date section. */
function asLiteral(token: FormatToken): FormatToken {
  switch (token.type) {
    case "digit":
      return { type: "literal", text: token.placeholder };
    case "point":
      return { type: "literal", text: "." };
    case "comma":
      return { type: "literal", text: "," };
    case "percent":
      return { type: "literal", text: "%" };
    case "slash":
      return { type: "literal", text: "/" };
    case "denominator":
      return { type: "literal", text: token.text };
    case "text":
      return { type: "literal", text: "@" };
    case "general":
      return { type: "literal", text: "General" };
    case "exponent":
      return { type: "literal", text: token.text };
    default:
      return token;
  }
}

/** A date section's body: placeholders and punctuation as literals, the
 * zeros after `<seconds>.` as fractional seconds (at most three). */
function dateTokens(raw: readonly FormatToken[]): FormatToken[] {
  const out: FormatToken[] = [];
  for (let i = 0; i < raw.length; i += 1) {
    const token = raw[i]!;
    const previous = out.at(-1);
    const afterSeconds =
      previous !== undefined &&
      ((previous.type === "date" && previous.part === "second") || (previous.type === "elapsed" && previous.unit === "s"));
    if (token.type === "point" && afterSeconds) {
      let zeros = 0;
      while (i + 1 + zeros < raw.length) {
        const next = raw[i + 1 + zeros]!;
        if (next.type !== "digit" || next.placeholder !== "0") {
          break;
        }
        zeros += 1;
      }
      if (zeros > 0) {
        out.push({ type: "subsecond", digits: Math.min(3, zeros) });
        for (let extra = 3; extra < zeros; extra += 1) {
          out.push({ type: "literal", text: "0" });
        }
        i += zeros;
        continue;
      }
    }
    out.push(asLiteral(token));
  }
  return out;
}

/** Whether a number section's slash makes it a fraction: a placeholder
 * before it and a placeholder or a fixed denominator right after it. */
function isFraction(tokens: readonly FormatToken[]): boolean {
  const slash = tokens.findIndex((token) => token.type === "slash");
  if (slash <= 0) {
    return false;
  }
  const before = tokens.slice(0, slash).some((token) => token.type === "digit");
  const after = tokens[slash + 1];
  return before && after !== undefined && (after.type === "digit" || after.type === "denominator");
}

/** Classifies one raw section (see the module header). */
function classify(raw: RawSection, state: ScanState): FormatSection {
  if (raw.tokens.some(({ token }) => token.type === "era")) {
    state.unsupported = true;
  }
  const entries = raw.tokens.filter((entry): entry is { token: FormatToken; offset: number } => entry.token.type !== "era");
  let tokens = entries.map((entry) => entry.token);
  let kind: FormatSection["kind"];
  if (tokens.some(isDateToken)) {
    if (tokens.some((token) => token.type === "exponent")) {
      state.unsupported = true; // `e` next to date tokens is the era year
    }
    tokens = dateTokens(tokens);
    applyMinuteRule(tokens);
    kind = "date";
  } else {
    const firstText = entries.find(({ token }) => token.type === "text");
    const firstDigit = entries.find(({ token }) => token.type === "digit");
    if (firstText && firstDigit) {
      const later = firstText.offset > firstDigit.offset ? firstText : entries.find(({ token, offset }) => token.type === "digit" && offset > firstText.offset)!;
      throw new FormatCodeError("A section cannot mix digit placeholders and @", later.offset);
    }
    entries.forEach(({ token, offset }, index) => {
      if (token.type === "exponent") {
        const next = entries[index + 1]?.token;
        if (next?.type !== "digit") {
          throw new FormatCodeError("An exponent needs digit placeholders", offset);
        }
      }
    });
    if (firstText) {
      kind = "text";
    } else if (tokens.some((token) => token.type === "general")) {
      kind = "general";
    } else if (tokens.some((token) => token.type === "exponent")) {
      kind = "scientific";
    } else if (isFraction(tokens)) {
      kind = "fraction";
    } else if (firstDigit) {
      kind = "number";
    } else {
      kind = "literal";
    }
    if (kind !== "fraction") {
      tokens = tokens.map((token) => (token.type === "slash" || token.type === "denominator" ? asLiteral(token) : token));
    }
    if (kind === "text" || kind === "literal" || kind === "general") {
      tokens = tokens.map((token) => (token.type === "point" || token.type === "comma" ? asLiteral(token) : token));
    }
  }
  // *x: the last fill in a section wins; the others are dropped.
  const lastFill = tokens.findLastIndex((token) => token.type === "fill");
  if (lastFill >= 0) {
    tokens = tokens.filter((token, index) => token.type !== "fill" || index === lastFill);
  }
  const section: FormatSection = { kind, tokens };
  if (raw.condition !== undefined) {
    section.condition = raw.condition;
  }
  if (raw.color !== undefined) {
    section.color = raw.color;
  }
  if (raw.lcid !== undefined) {
    section.lcid = raw.lcid;
  }
  return section;
}

function parseUncached(code: string): ParsedFormat {
  const state: ScanState = { unsupported: false };
  const sections = scan(code, state).map((raw) => classify(raw, state));
  return state.unsupported ? { sections, unsupported: true } : { sections };
}

const cache = new Map<string, ParsedFormat | FormatCodeError>();

/** Parses a format code (see the module header); cached. Throws
 * FormatCodeError for a structural error. */
export function parseFormatCode(code: string): ParsedFormat {
  let entry = cache.get(code);
  if (entry === undefined) {
    try {
      entry = parseUncached(code);
    } catch (error) {
      if (!(error instanceof FormatCodeError)) {
        throw error;
      }
      entry = error;
    }
    if (cache.size >= FORMAT_CACHE_LIMIT) {
      cache.clear();
    }
    cache.set(code, entry);
  }
  if (entry instanceof FormatCodeError) {
    throw entry;
  }
  return entry;
}

/** A body token as written, at the offset (UTF-16 index) of its first
 * character. */
export interface ScannedToken {
  token: FormatToken;
  offset: number;
}

/** One section of a code as written: its kind as parseFormatCode classifies
 * it and its body tokens as scanned — before classification reads some as
 * literals (a date section's placeholders, a text section's points). */
export interface FormatCodeSection {
  kind: FormatSection["kind"];
  tokens: ScannedToken[];
}

/** A code's sections with their tokens' source offsets (calendar tokens
 * left out). Throws FormatCodeError for an invalid code, as parseFormatCode
 * does. */
export function formatCodeSections(code: string): FormatCodeSection[] {
  const { sections } = parseFormatCode(code);
  return scan(code, { unsupported: false }).map((raw, index) => ({
    kind: sections[index]!.kind,
    tokens: raw.tokens.filter((entry): entry is ScannedToken => entry.token.type !== "era"),
  }));
}

/** How many codes the parse cache holds (for the cache-limit test). */
export function formatCacheSize(): number {
  return cache.size;
}
