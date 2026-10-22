// ODS data styles ↔ format codes (SPEC-118 "ODS number styles"). An .ods
// cell style names a `number:*-style` (in content.xml's automatic styles or
// styles.xml's common styles) by `style:data-style-name`; the reader turns it
// into the invariant format code every other part of the engine speaks, and
// SPEC-122's writer turns a code back into data-style elements.
//
// odsNumberStyleToCode walks one data style:
// - `number:number`: integer `0`×min-integer-digits (`#` for 0), padded to
//   `#,##0` form with grouping; `.` + `0`×min-decimal-places +
//   `#`×(decimal-places − min); without decimal-places it is `General`;
//   display-factor 1000^k → k scaling commas; embedded text → the quoted text
//   p digits left of the point.
// - `number:scientific-number` (exponent-interval n → `#`×(n − i) + `0`×i),
//   `number:fraction` (an integer part only when min-integer-digits is
//   present), percentage (`%`), currency (`[$sym-LCID]`, or `[$sym]` for an
//   unknown language/country pair), boolean (`"TRUE";"TRUE";"FALSE"`), text
//   (`@`), fill-character (`*x`), the date and time parts (long forms, textual
//   months), `number:time-style number:truncate-on-overflow="false"` → the
//   first time token bracketed.
// - `number:text` is bare when it holds only `- / : . , ( )` and spaces in a
//   date/time section or only `- ( ) $ +` and spaces in a number section;
//   else it is quoted, `\"` escaping a quote.
// - `style:text-properties@fo:color` → the colour name of that exact hex,
//   else `[ColorN]` for a match in the default palette, else none.
// - `number:language`/`number:country` on a style with textual names →
//   a `[$-LCID]` prefix when the pair is one of the six locales.
// - `style:map style:condition="value()OPn"` → a section `[OPn]` + the
//   applied style's code, the main style last; LibreOffice's sign patterns
//   (`value()>=0` for two sections; `value()>0`, `value()<0` for three; the
//   text style as main with `>0`, `<0`, `=0`) become plain sections, since
//   its negative style writes its own "-".
// - `number:era` or any unknown child → `General` (the source style is
//   reused unchanged by sourceStyleName).
// odsNumberStyleXml is the inverse, for SPEC-122's writer: the sub-styles
// `${name}P0`, `${name}P1`… and the main `${name}` with its style:maps.
import { NUMBER_NAMESPACE, STYLE_NAMESPACE, FO_NAMESPACE } from "../../core/odfXml";
import { escapeXmlAttribute, escapeXmlText } from "../../core/xml";
import { INDEXED_COLORS } from "../styles";
import { LCID_SYSTEM_LONG_DATE, LCID_SYSTEM_LONG_TIME, parseFormatCode, type FormatSection, type FormatToken } from "./parse";

/** LibreOffice's extension namespace (`loext:min-decimal-places`, …). */
const LOEXT_NAMESPACE = "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0";

/** The six locales' (language, country) pairs ↔ their LCIDs. */
const LOCALE_PAIRS: ReadonlyArray<readonly [string, string, number]> = [
  ["en", "US", 0x409],
  ["en", "GB", 0x809],
  ["de", "DE", 0x407],
  ["fr", "FR", 0x40c],
  ["es", "ES", 0xc0a],
  ["cs", "CZ", 0x405],
];

const NAMED_COLOR_HEX: ReadonlyArray<readonly [string, string]> = [
  ["Black", "#000000"],
  ["Blue", "#0000ff"],
  ["Cyan", "#00ffff"],
  ["Green", "#00ff00"],
  ["Magenta", "#ff00ff"],
  ["Red", "#ff0000"],
  ["White", "#ffffff"],
  ["Yellow", "#ffff00"],
];

const DATA_STYLES: ReadonlySet<string> = new Set([
  "number-style",
  "currency-style",
  "percentage-style",
  "date-style",
  "time-style",
  "boolean-style",
  "text-style",
]);

function lcidOf(language: string | null, country: string | null): number | undefined {
  if (!language || !country) {
    return undefined;
  }
  return LOCALE_PAIRS.find(([lang, region]) => lang === language.toLowerCase() && region === country.toUpperCase())?.[2];
}

function lcidHex(lcid: number): string {
  return lcid.toString(16).toUpperCase();
}

// --- ODS → code -----------------------------------------------------------------

function numberAttr(el: Element, local: string): string | null {
  return el.getAttributeNS(NUMBER_NAMESPACE, local) ?? null;
}

/** An attribute in the number namespace, or LibreOffice's loext twin. */
function numberOrLoext(el: Element, local: string): string | null {
  const value = el.getAttributeNS(NUMBER_NAMESPACE, local);
  return value !== null && value !== "" ? value : el.getAttributeNS(LOEXT_NAMESPACE, local) || null;
}

function intAttr(value: string | null): number | undefined {
  if (value === null || value.trim() === "") {
    return undefined;
  }
  const n = Number(value);
  return Number.isInteger(n) && n >= 0 ? n : undefined;
}

/** A `number:text` as code: bare when only the section's safe characters,
 * else quoted with `\"` escapes. */
function textCode(text: string, dateSection: boolean): string {
  if (text === "") {
    return "";
  }
  const safe = dateSection ? /^[-/:.,() ]+$/ : /^[-()$+ ]+$/;
  return safe.test(text) ? text : `"${text.replaceAll('"', '\\"')}"`;
}

/** The colour name (or [ColorN]) of an exact hex. */
function colorCode(hex: string | null): string {
  const value = hex?.trim().toLowerCase();
  if (!value) {
    return "";
  }
  const named = NAMED_COLOR_HEX.find(([, h]) => h === value);
  if (named) {
    return `[${named[0]}]`;
  }
  const index = INDEXED_COLORS.findIndex((color, i) => i >= 8 && color === value);
  return index >= 8 ? `[Color${index - 7}]` : "";
}

/** The integer placeholders of `number:number`: `0`×i (`#` for 0), padded
 * with `#` to `#,##0` form when grouping. */
function integerPlaceholders(minInteger: number, grouping: boolean): string[] {
  if (!grouping) {
    return minInteger === 0 ? ["#"] : Array.from({ length: minInteger }, () => "0");
  }
  const digits = [...Array.from({ length: Math.max(0, 4 - minInteger) }, () => "#"), ...Array.from({ length: minInteger }, () => "0")];
  digits.splice(digits.length - 3, 0, ",");
  return digits;
}

function decimalsCode(el: Element): string | null {
  const decimals = intAttr(numberAttr(el, "decimal-places"));
  if (decimals === undefined) {
    return null;
  }
  const min = Math.min(decimals, intAttr(numberOrLoext(el, "min-decimal-places")) ?? decimals);
  return decimals > 0 ? `.${"0".repeat(min)}${"#".repeat(decimals - min)}` : "";
}

/** `number:number` → its code (General without decimal-places). */
function numberCode(el: Element): string {
  const fraction = decimalsCode(el);
  if (fraction === null) {
    return "General";
  }
  const integer = integerPlaceholders(intAttr(numberAttr(el, "min-integer-digits")) ?? 0, numberAttr(el, "grouping") === "true");
  // Embedded texts, each p placeholders left of the point.
  const embedded = Array.from(el.children)
    .filter((child) => child.namespaceURI === NUMBER_NAMESPACE && child.localName === "embedded-text")
    .map((child) => ({ position: intAttr(numberAttr(child, "position")) ?? 0, text: child.textContent ?? "" }))
    .toSorted((a, b) => a.position - b.position);
  const parts = integer.map((ch) => ch);
  for (const { position, text } of embedded.toReversed()) {
    let seen = 0;
    let at = parts.length;
    while (at > 0 && seen < position) {
      at -= 1;
      if (parts[at] === "0" || parts[at] === "#") {
        seen += 1;
      }
    }
    parts.splice(at, 0, `"${text.replaceAll('"', '\\"')}"`);
  }
  const factor = Number(numberAttr(el, "display-factor") ?? 1);
  const scale = factor > 1 ? Math.round(Math.log10(factor) / 3) : 0;
  return `${parts.join("")}${",".repeat(scale)}${fraction}`;
}

function scientificCode(el: Element): string {
  const fraction = decimalsCode(el) ?? "";
  const minInteger = intAttr(numberAttr(el, "min-integer-digits")) ?? 1;
  const interval = intAttr(numberOrLoext(el, "exponent-interval")) ?? minInteger;
  const exponentDigits = intAttr(numberAttr(el, "min-exponent-digits")) ?? 2;
  const sign = numberOrLoext(el, "forced-exponent-sign") === "false" ? "-" : "+";
  return `${"#".repeat(Math.max(0, interval - minInteger))}${"0".repeat(minInteger)}${fraction}E${sign}${"0".repeat(Math.max(1, exponentDigits))}`;
}

function fractionCode(el: Element): string {
  const minInteger = intAttr(numberAttr(el, "min-integer-digits"));
  const integer = minInteger === undefined ? "" : `${minInteger === 0 ? "#" : "0".repeat(minInteger)} `;
  const numerator = "?".repeat(Math.max(1, intAttr(numberAttr(el, "min-numerator-digits")) ?? 1));
  const fixed = intAttr(numberAttr(el, "denominator-value"));
  const denominator = fixed !== undefined && fixed > 0 ? String(fixed) : "?".repeat(Math.max(1, intAttr(numberAttr(el, "min-denominator-digits")) ?? 1));
  return `${integer}${numerator}/${denominator}`;
}

function isLong(el: Element): boolean {
  return numberAttr(el, "style") === "long";
}

/** One data style's own section (its maps aside), or null for General (an
 * era, an unknown child, a non-Gregorian calendar). */
function sectionCode(style: Element): string | null {
  const kind = style.localName;
  if (kind === "boolean-style") {
    return '"TRUE";"TRUE";"FALSE"';
  }
  const dateSection = kind === "date-style" || kind === "time-style";
  const elapsed = kind === "time-style" && numberAttr(style, "truncate-on-overflow") === "false";
  let color = "";
  let textual = false;
  let bracketed = false;
  const body: string[] = [];
  for (const child of Array.from(style.children)) {
    if (child.namespaceURI === STYLE_NAMESPACE) {
      if (child.localName === "text-properties") {
        color = colorCode(child.getAttributeNS(FO_NAMESPACE, "color"));
      }
      continue; // style:map is the caller's
    }
    if (child.namespaceURI !== NUMBER_NAMESPACE) {
      continue;
    }
    const calendar = numberAttr(child, "calendar");
    if (calendar && calendar !== "gregorian") {
      return null;
    }
    /** A time token, bracketed when it is the first of an elapsed style. */
    const time = (token: string): string => {
      if (elapsed && !bracketed) {
        bracketed = true;
        return `[${token}]`;
      }
      return token;
    };
    switch (child.localName) {
      case "number":
        body.push(numberCode(child));
        break;
      case "scientific-number":
        body.push(scientificCode(child));
        break;
      case "fraction":
        body.push(fractionCode(child));
        break;
      case "text": {
        const text = child.textContent ?? "";
        if (kind === "percentage-style" && text.includes("%")) {
          body.push(text.split("%").map((piece) => textCode(piece, false)).join("%"));
        } else {
          body.push(textCode(text, dateSection));
        }
        break;
      }
      case "currency-symbol": {
        const symbol = child.textContent ?? "";
        const lcid = lcidOf(numberAttr(child, "language"), numberAttr(child, "country"));
        if (symbol !== "") {
          body.push(`[$${symbol}${lcid !== undefined ? `-${lcidHex(lcid)}` : ""}]`);
        }
        break;
      }
      case "text-content":
        body.push("@");
        break;
      case "fill-character":
        body.push(`*${(child.textContent ?? " ")[0] ?? " "}`);
        break;
      case "day":
        body.push(isLong(child) ? "dd" : "d");
        break;
      case "month": {
        const isTextual = numberAttr(child, "textual") === "true";
        textual ||= isTextual;
        body.push(isTextual ? (isLong(child) ? "mmmm" : "mmm") : isLong(child) ? "mm" : "m");
        break;
      }
      case "year":
        body.push(isLong(child) ? "yyyy" : "yy");
        break;
      case "day-of-week":
        textual = true;
        body.push(isLong(child) ? "dddd" : "ddd");
        break;
      case "quarter":
        body.push(isLong(child) ? "QQ" : "Q");
        break;
      case "week-of-year":
        body.push("WW");
        break;
      case "hours":
        body.push(time(isLong(child) ? "hh" : "h"));
        break;
      case "minutes":
        body.push(time(isLong(child) ? "mm" : "m"));
        break;
      case "seconds": {
        const decimals = intAttr(numberAttr(child, "decimal-places")) ?? 0;
        body.push(time(isLong(child) ? "ss" : "s") + (decimals > 0 ? `.${"0".repeat(Math.min(3, decimals))}` : ""));
        break;
      }
      case "am-pm":
        body.push("AM/PM");
        break;
      default:
        return null; // number:era, or an element this mapping does not know
    }
  }
  let code = body.join("");
  if (kind === "percentage-style" && !code.includes("%")) {
    code += "%";
  }
  const lcid = textual ? lcidOf(numberAttr(style, "language"), numberAttr(style, "country")) : undefined;
  return `${color}${lcid !== undefined ? `[$-${lcidHex(lcid)}]` : ""}${code}`;
}

/** A style:map's condition as `OPn` (">=0"), or null when it is not a
 * `value()` comparison. */
function mapCondition(condition: string | null): string | null {
  const match = /^value\(\)\s*(<=|>=|<>|!=|<|>|=)\s*(.+)$/.exec(condition?.trim() ?? "");
  if (!match) {
    return null;
  }
  return `${match[1] === "!=" ? "<>" : match[1]}${match[2]!.trim()}`;
}

/** The format code a data style stands for (see the module header);
 * undefined when the element is not a data style. `lookup` finds the style
 * a style:map applies, by name. */
export function odsNumberStyleToCode(style: Element, lookup: (name: string) => Element | undefined): string | undefined {
  if (style.namespaceURI !== NUMBER_NAMESPACE || !DATA_STYLES.has(style.localName)) {
    return undefined;
  }
  const main = sectionCode(style);
  if (main === null) {
    return "General";
  }
  const maps = Array.from(style.children)
    .filter((child) => child.namespaceURI === STYLE_NAMESPACE && child.localName === "map")
    .map((map) => ({
      condition: mapCondition(map.getAttributeNS(STYLE_NAMESPACE, "condition")),
      applied: lookup(map.getAttributeNS(STYLE_NAMESPACE, "apply-style-name") ?? ""),
    }))
    .filter((map): map is { condition: string; applied: Element } => map.condition !== null && map.applied !== undefined)
    .slice(0, 3);
  if (maps.length === 0) {
    return main;
  }
  const codes = maps.map((map) => sectionCode(map.applied) ?? "General");
  const conditions = maps.map((map) => map.condition).join(" ");
  const text = style.localName === "text-style";
  const plain =
    (!text && (conditions === ">=0" || conditions === ">0 <0")) ||
    (text && (conditions === ">0 <0 =0" || conditions === ">=0 <0" || conditions === ">=0"));
  const sections = plain ? codes : codes.map((code, i) => `[${maps[i]!.condition}]${code}`);
  return [...sections, main].join(";");
}

// --- code → ODS -----------------------------------------------------------------

/** The XML of one number:text run (empty text writes nothing). */
function textElement(text: string): string {
  return text === "" ? "" : `<number:text>${escapeXmlText(text)}</number:text>`;
}

function colorHex(color: string | undefined): string | undefined {
  if (color === undefined) {
    return undefined;
  }
  const named = NAMED_COLOR_HEX.find(([name]) => name === color);
  if (named) {
    return named[1];
  }
  const index = /^Color(\d+)$/.exec(color);
  return index ? INDEXED_COLORS[Number(index[1]) + 7] : undefined;
}

/** Collects body XML, merging consecutive literal text into one number:text. */
class BodyWriter {
  private parts: string[] = [];
  private pending = "";

  text(text: string): void {
    this.pending += text;
  }

  element(xml: string): void {
    this.flush();
    this.parts.push(xml);
  }

  flush(): void {
    if (this.pending !== "") {
      this.parts.push(textElement(this.pending));
      this.pending = "";
    }
  }

  xml(): string {
    this.flush();
    return this.parts.join("");
  }
}

function attrs(pairs: ReadonlyArray<readonly [string, string | number | undefined]>): string {
  return pairs
    .filter((pair): pair is readonly [string, string | number] => pair[1] !== undefined)
    .map(([name, value]) => ` ${name}="${escapeXmlAttribute(String(value))}"`)
    .join("");
}

/** A date or time section's body. */
function dateBody(section: FormatSection): string {
  const body = new BodyWriter();
  const tokens = section.tokens;
  tokens.forEach((token, index) => {
    const long = (min: number): readonly [string, string | undefined] => ["number:style", "length" in token && token.length >= min ? "long" : undefined];
    switch (token.type) {
      case "date":
        switch (token.part) {
          case "year":
            body.element(`<number:year${attrs([long(3)])}/>`);
            break;
          case "month":
            body.element(
              token.length >= 3
                ? `<number:month${attrs([["number:textual", "true"], token.length === 4 || token.length >= 6 ? ["number:style", "long"] : ["number:style", undefined]])}/>`
                : `<number:month${attrs([long(2)])}/>`,
            );
            break;
          case "day":
            body.element(token.length >= 3 ? `<number:day-of-week${attrs([long(4)])}/>` : `<number:day${attrs([long(2)])}/>`);
            break;
          case "quarter":
            body.element(`<number:quarter${attrs([long(2)])}/>`);
            break;
          case "week":
            body.element("<number:week-of-year/>");
            break;
          case "hour":
            body.element(`<number:hours${attrs([long(2)])}/>`);
            break;
          case "minute":
            body.element(`<number:minutes${attrs([long(2)])}/>`);
            break;
          case "second": {
            const next = tokens[index + 1];
            body.element(`<number:seconds${attrs([long(2), ["number:decimal-places", next?.type === "subsecond" ? next.digits : undefined]])}/>`);
            break;
          }
        }
        break;
      case "elapsed": {
        const element = token.unit === "h" ? "hours" : token.unit === "m" ? "minutes" : "seconds";
        const next = tokens[index + 1];
        const decimals = token.unit === "s" && next?.type === "subsecond" ? next.digits : undefined;
        body.element(`<number:${element}${attrs([long(2), ["number:decimal-places", decimals]])}/>`);
        break;
      }
      case "ampm":
        body.element("<number:am-pm/>");
        break;
      case "subsecond":
        break; // written as the seconds' decimal-places
      case "literal":
        body.text(token.text);
        break;
      case "space":
        body.text(" ");
        break;
      case "currency":
        body.text(token.symbol);
        break;
      case "fill":
        body.element(`<number:fill-character>${escapeXmlText(token.char)}</number:fill-character>`);
        break;
      default:
        break;
    }
  });
  return body.xml();
}

type Placeholder = "0" | "#" | "?";

function placeholders(tokens: readonly FormatToken[], from: number, to: number): Placeholder[] {
  const out: Placeholder[] = [];
  for (let i = from; i < to; i += 1) {
    const token = tokens[i]!;
    if (token.type === "digit") {
      out.push(token.placeholder);
    }
  }
  return out;
}

function countZeros(list: readonly Placeholder[]): number {
  return list.filter((p) => p === "0").length;
}

/** A number-kind section's body: literals around one number element. */
function numberBody(section: FormatSection, lcid: number | undefined): string {
  const tokens = section.tokens;
  const body = new BodyWriter();
  const firstDigit = tokens.findIndex((token) => token.type === "digit");
  const lastDigit = tokens.findLastIndex((token) => token.type === "digit");
  const coreStart = section.kind === "general" ? tokens.findIndex((token) => token.type === "general") : firstDigit;
  const coreEnd = section.kind === "general" ? coreStart : lastDigit;
  const pair = lcid === undefined ? undefined : LOCALE_PAIRS.find(([, , id]) => id === lcid);
  const emitOutside = (token: FormatToken): void => {
    switch (token.type) {
      case "literal":
        body.text(token.text);
        break;
      case "space":
        body.text(" ");
        break;
      case "percent":
        body.text("%");
        break;
      case "comma":
        body.text(",");
        break;
      case "point":
        body.text(".");
        break;
      case "text":
        body.element("<number:text-content/>");
        break;
      case "fill":
        body.element(`<number:fill-character>${escapeXmlText(token.char)}</number:fill-character>`);
        break;
      case "currency":
        body.element(
          `<number:currency-symbol${attrs([["number:language", pair?.[0]], ["number:country", pair?.[1]]])}>${escapeXmlText(token.symbol)}</number:currency-symbol>`,
        );
        break;
      default:
        break;
    }
  };
  if (coreStart < 0) {
    tokens.forEach(emitOutside);
    return body.xml();
  }
  tokens.slice(0, coreStart).forEach(emitOutside);
  body.element(coreElement(section, coreStart, coreEnd));
  tokens.slice(coreEnd + 1).forEach((token, offset) => {
    const index = coreEnd + 1 + offset;
    // Scaling commas right after the core are the core's display factor.
    if (token.type === "comma" && tokens.slice(coreEnd + 1, index + 1).every((t) => t.type === "comma")) {
      return;
    }
    emitOutside(token);
  });
  return body.xml();
}

/** The number:number / scientific-number / fraction element of tokens
 * [start, end]. */
function coreElement(section: FormatSection, start: number, end: number): string {
  const tokens = section.tokens;
  if (section.kind === "general") {
    return `<number:number${attrs([["number:min-integer-digits", 1]])}/>`;
  }
  if (section.kind === "fraction") {
    const slash = tokens.findIndex((token) => token.type === "slash");
    let numeratorStart = slash;
    while (numeratorStart > 0 && tokens[numeratorStart - 1]!.type === "digit") {
      numeratorStart -= 1;
    }
    const integer = placeholders(tokens, 0, numeratorStart);
    const numerator = placeholders(tokens, numeratorStart, slash);
    const after = tokens[slash + 1];
    const denominator = after?.type === "denominator" ? undefined : placeholders(tokens, slash + 1, end + 1);
    return `<number:fraction${attrs([
      ["number:min-integer-digits", integer.length > 0 ? countZeros(integer) : undefined],
      ["number:min-numerator-digits", numerator.length],
      ["number:min-denominator-digits", denominator?.length],
      ["number:denominator-value", after?.type === "denominator" ? after.value : undefined],
    ])}/>`;
  }
  const exponent = tokens.findIndex((token) => token.type === "exponent");
  const mantissaEnd = exponent >= 0 ? exponent : end + 1;
  const point = tokens.findIndex((token, index) => token.type === "point" && index < mantissaEnd);
  const intEnd = point >= 0 ? point : mantissaEnd;
  const integer = placeholders(tokens, 0, intEnd);
  const fraction = point >= 0 ? placeholders(tokens, point + 1, mantissaEnd) : [];
  const decimals: Array<readonly [string, number | undefined]> = [
    ["number:decimal-places", fraction.length],
    ["number:min-decimal-places", countZeros(fraction)],
  ];
  if (exponent >= 0) {
    const expText = tokens[exponent]!.type === "exponent" ? (tokens[exponent] as { text: string }).text : "E+";
    return `<number:scientific-number${attrs([
      ...decimals,
      ["number:min-integer-digits", countZeros(integer)],
      ["number:min-exponent-digits", countZeros(placeholders(tokens, exponent + 1, end + 1))],
      ["number:exponent-interval", integer.includes("#") && integer.length > 1 ? integer.length : countZeros(integer)],
      ["number:forced-exponent-sign", expText[1] === "-" ? "false" : undefined],
    ])}/>`;
  }
  // Commas between integer placeholders group; commas after the last one
  // (in or right after the core) scale by 1000 each.
  let grouping = false;
  let scale = 0;
  const intPlaceholderIndices = tokens.map((token, index) => (token.type === "digit" && index < intEnd ? index : -1)).filter((i) => i >= 0);
  const lastInt = intPlaceholderIndices.at(-1) ?? -1;
  const embedded: string[] = [];
  for (let i = start; i < intEnd; i += 1) {
    const token = tokens[i]!;
    if (token.type === "comma") {
      if (i < lastInt) {
        grouping = true;
      } else {
        scale += 1;
      }
    } else if (token.type === "literal" && i < lastInt) {
      const position = intPlaceholderIndices.filter((index) => index > i).length;
      embedded.push(`<number:embedded-text${attrs([["number:position", position]])}>${escapeXmlText(token.text)}</number:embedded-text>`);
    }
  }
  // Commas right after the core (`#,##0,"K"`, `0.0,`): the core's region
  // ends at its last placeholder, so the loop above never saw them.
  for (let i = end + 1; i < tokens.length && tokens[i]!.type === "comma"; i += 1) {
    scale += 1;
  }
  const numberAttrs = attrs([
    ...decimals,
    ["number:min-integer-digits", countZeros(integer)],
    ["number:grouping", grouping ? "true" : undefined],
    ["number:display-factor", scale > 0 ? 1000 ** scale : undefined],
  ]);
  return embedded.length > 0 ? `<number:number${numberAttrs}>${embedded.join("")}</number:number>` : `<number:number${numberAttrs}/>`;
}

/** The data-style element kind a section writes as. */
function elementKind(section: FormatSection): string {
  if (section.kind === "date" || section.lcid === LCID_SYSTEM_LONG_DATE || section.lcid === LCID_SYSTEM_LONG_TIME) {
    const dated = section.tokens.some((token) => token.type === "date" && token.part !== "hour" && token.part !== "minute" && token.part !== "second");
    return dated || section.lcid === LCID_SYSTEM_LONG_DATE ? "date-style" : "time-style";
  }
  if (section.kind === "text") {
    return "text-style";
  }
  if (section.tokens.some((token) => token.type === "percent")) {
    return "percentage-style";
  }
  if (section.tokens.some((token) => token.type === "currency")) {
    return "currency-style";
  }
  return "number-style";
}

/** One data-style element for a section. */
function sectionXml(section: FormatSection, name: string, extra: { volatile?: boolean; maps?: string } = {}): string {
  const kind = elementKind(section);
  const pair = section.lcid === undefined ? undefined : LOCALE_PAIRS.find(([, , id]) => id === section.lcid);
  const elapsed = section.tokens.some((token) => token.type === "elapsed");
  const styleAttrs = attrs([
    ["style:name", name],
    ["style:volatile", extra.volatile ? "true" : undefined],
    ["number:language", kind === "date-style" || kind === "time-style" ? pair?.[0] : undefined],
    ["number:country", kind === "date-style" || kind === "time-style" ? pair?.[1] : undefined],
    ["number:truncate-on-overflow", kind === "time-style" && elapsed ? "false" : undefined],
  ]);
  const hex = colorHex(section.color);
  const color = hex === undefined ? "" : `<style:text-properties fo:color="${hex}"/>`;
  const body = kind === "date-style" || kind === "time-style" ? dateBody(section) : numberBody(section, section.lcid);
  return `<number:${kind}${styleAttrs}>${color}${body}${extra.maps ?? ""}</number:${kind}>`;
}

function mapXml(condition: string, applied: string): string {
  return `<style:map style:condition="${escapeXmlAttribute(`value()${condition}`)}" style:apply-style-name="${escapeXmlAttribute(applied)}"/>`;
}

function conditionText(section: FormatSection): string {
  return section.condition ? `${section.condition.op}${section.condition.value}` : ">=0";
}

/** The data-style XML for a code (see the module header): the sub-styles
 * `${name}P0`, `${name}P1`… first, then the main `${name}` with one
 * style:map per sub-style. An invalid or unsupported code writes General. */
export function odsNumberStyleXml(code: string, name: string): string {
  let sections: FormatSection[];
  try {
    const parsed = /^general$/i.test(code.trim()) || code.trim() === "" ? null : parseFormatCode(code);
    sections = parsed === null || parsed.unsupported ? [] : parsed.sections;
  } catch {
    sections = [];
  }
  if (sections.length === 0) {
    return `<number:number-style${attrs([["style:name", name]])}><number:number${attrs([["number:min-integer-digits", 1]])}/></number:number-style>`;
  }
  const textIndex = sections.length === 4 ? 3 : sections.at(-1)!.kind === "text" ? sections.length - 1 : -1;
  const numeric = sections.filter((_, index) => index !== textIndex);
  const conditioned = numeric.slice(0, 2).some((section) => section.condition !== undefined);
  const subs: Array<{ section: FormatSection; condition: string }> = [];
  let main: FormatSection;
  if (textIndex >= 0) {
    main = sections[textIndex]!;
    if (conditioned) {
      numeric.forEach((section) => subs.push({ section, condition: conditionText(section) }));
    } else if (numeric.length === 3) {
      subs.push({ section: numeric[0]!, condition: ">0" }, { section: numeric[1]!, condition: "<0" }, { section: numeric[2]!, condition: "=0" });
    } else if (numeric.length === 2) {
      subs.push({ section: numeric[0]!, condition: ">=0" }, { section: numeric[1]!, condition: "<0" });
    } else if (numeric.length === 1) {
      // One numeric section serves every number: its negative twin writes
      // the "-" the automatic sign would have.
      const negative: FormatSection = { ...numeric[0]!, tokens: [{ type: "literal", text: "-" }, ...numeric[0]!.tokens] };
      subs.push({ section: numeric[0]!, condition: ">=0" }, { section: negative, condition: "<0" });
    }
  } else if (numeric.length === 1) {
    main = numeric[0]!;
  } else if (conditioned) {
    const fallback = numeric.find((section) => section.condition === undefined) ?? numeric.at(-1)!;
    main = fallback;
    numeric.filter((section) => section !== fallback).forEach((section) => subs.push({ section, condition: conditionText(section) }));
  } else if (numeric.length === 2) {
    main = numeric[1]!;
    subs.push({ section: numeric[0]!, condition: ">=0" });
  } else {
    main = numeric[2]!;
    subs.push({ section: numeric[0]!, condition: ">0" }, { section: numeric[1]!, condition: "<0" });
  }
  const subXml = subs.map((sub, index) => sectionXml(sub.section, `${name}P${index}`, { volatile: true })).join("");
  const maps = subs.map((sub, index) => mapXml(sub.condition, `${name}P${index}`)).join("");
  return subXml + sectionXml(main, name, { maps });
}
