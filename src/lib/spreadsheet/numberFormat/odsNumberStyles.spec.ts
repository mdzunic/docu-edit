// SPEC-118 "ODS number styles": odsNumberStyleToCode over LibreOffice-shaped
// data styles, each row of the Reference's table (number, scientific,
// fraction, percentage, currency, boolean and text styles, number:text bare
// or quoted, fills, every date and time part, elapsed time styles, era and
// unknown children as General, fo:color, a textual date's language, and
// style:map conditions — LibreOffice's sign patterns as plain sections);
// then odsNumberStyleXml, the inverse SPEC-122's writer calls (sub-styles
// ${name}P0, ${name}P1… with style:volatile, the main style last with its
// style:maps); and the corpus round trip —
// odsNumberStyleToCode(parse(odsNumberStyleXml(c, "N1"))) === c.
import { describe, expect, it } from "vitest";
import { INDEXED_COLORS } from "../styles";
import { BUILTIN_NUMBER_FORMATS, LOCALE_DEPENDENT_BUILTINS } from "./builtin";
import { odsNumberStyleToCode, odsNumberStyleXml } from "./odsNumberStyles";

const NUMBER_NS = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0";
const STYLE_NS = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const FO_NS = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const LOEXT_NS = "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0";

/** The elements of an XML fragment, parsed in a root declaring the ODF
 * namespaces. */
function parse(xml: string): Element[] {
  const doc = new DOMParser().parseFromString(
    `<root xmlns:number="${NUMBER_NS}" xmlns:style="${STYLE_NS}" xmlns:fo="${FO_NS}" xmlns:loext="${LOEXT_NS}">${xml}</root>`,
    "application/xml",
  );
  expect(doc.getElementsByTagName("parsererror"), xml).toHaveLength(0);
  return Array.from(doc.documentElement.children);
}

/** odsNumberStyleToCode of the style named `name` (default: the last
 * element), the lookup finding styles by style:name. */
function codeOf(xml: string, name?: string): string | undefined {
  const styles = parse(xml);
  const lookup = (wanted: string): Element | undefined => styles.find((el) => el.getAttributeNS(STYLE_NS, "name") === wanted);
  const main = name === undefined ? styles.at(-1) : lookup(name);
  if (main === undefined) {
    throw new Error(`no style ${String(name)} in ${xml}`);
  }
  return odsNumberStyleToCode(main, lookup);
}

/** A data style N1 of `kind` around `body`. */
function dataStyle(kind: string, body: string, attrs = ""): string {
  return `<number:${kind} style:name="N1"${attrs}>${body}</number:${kind}>`;
}

function numberStyle(body: string, attrs = ""): string {
  return dataStyle("number-style", body, attrs);
}

function dateStyle(body: string, attrs = ""): string {
  return dataStyle("date-style", body, attrs);
}

function timeStyle(body: string, attrs = ""): string {
  return dataStyle("time-style", body, attrs);
}

const text = (content: string): string => `<number:text>${content}</number:text>`;
const map = (condition: string, applied: string): string => `<style:map style:condition="${condition}" style:apply-style-name="${applied}"/>`;
const INTEGER = '<number:number number:decimal-places="0" number:min-integer-digits="1"/>';
const TWO_DECIMALS = '<number:number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1"/>';

describe("odsNumberStyleToCode — the table, row by row", () => {
  it.each<[label: string, xml: string, code: string]>([
    // number:number
    ["number:number without decimal-places", numberStyle('<number:number number:min-integer-digits="1"/>'), "General"],
    ["decimal-places 2, min-decimal-places 2", numberStyle(TWO_DECIMALS), "0.00"],
    ["decimal-places 2, no min-decimal-places", numberStyle('<number:number number:decimal-places="2" number:min-integer-digits="1"/>'), "0.00"],
    ["decimal-places 3, min-decimal-places 1", numberStyle('<number:number number:decimal-places="3" number:min-decimal-places="1" number:min-integer-digits="1"/>'), "0.0##"],
    ["loext:min-decimal-places", numberStyle('<number:number number:decimal-places="3" loext:min-decimal-places="1" number:min-integer-digits="1"/>'), "0.0##"],
    ["min-integer-digits 0 → #", numberStyle('<number:number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="0"/>'), "#.00"],
    ["min-integer-digits 3", numberStyle('<number:number number:decimal-places="0" number:min-integer-digits="3"/>'), "000"],
    ["decimal-places 0", numberStyle(INTEGER), "0"],
    ["grouping", numberStyle('<number:number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1" number:grouping="true"/>'), "#,##0.00"],
    ["grouping, min-integer-digits 0", numberStyle('<number:number number:decimal-places="0" number:min-integer-digits="0" number:grouping="true"/>'), "#,###"],
    ["display-factor 1000", numberStyle('<number:number number:decimal-places="0" number:min-integer-digits="1" number:grouping="true" number:display-factor="1000"/>'), "#,##0,"],
    ["display-factor 1000000", numberStyle('<number:number number:decimal-places="0" number:min-integer-digits="1" number:display-factor="1000000"/>'), "0,,"],
    [
      "embedded-text p digits left of the point",
      numberStyle(
        '<number:number number:decimal-places="0" number:min-integer-digits="9">' +
          '<number:embedded-text number:position="4">-</number:embedded-text>' +
          '<number:embedded-text number:position="6">-</number:embedded-text></number:number>',
      ),
      '000"-"00"-"0000',
    ],
    // number:scientific-number
    ["scientific-number", numberStyle('<number:scientific-number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1" number:min-exponent-digits="2"/>'), "0.00E+00"],
    [
      "scientific-number, exponent-interval 3",
      numberStyle('<number:scientific-number number:decimal-places="1" number:min-decimal-places="1" number:min-integer-digits="1" number:min-exponent-digits="1" number:exponent-interval="3"/>'),
      "##0.0E+0",
    ],
    [
      "scientific-number, loext:exponent-interval 3",
      numberStyle('<number:scientific-number number:decimal-places="1" number:min-decimal-places="1" number:min-integer-digits="1" number:min-exponent-digits="1" loext:exponent-interval="3"/>'),
      "##0.0E+0",
    ],
    // number:fraction
    ["fraction, min-integer-digits 0", numberStyle('<number:fraction number:min-integer-digits="0" number:min-numerator-digits="1" number:min-denominator-digits="1"/>'), "# ?/?"],
    ["fraction, min-integer-digits 1", numberStyle('<number:fraction number:min-integer-digits="1" number:min-numerator-digits="1" number:min-denominator-digits="1"/>'), "0 ?/?"],
    ["fraction without min-integer-digits", numberStyle('<number:fraction number:min-numerator-digits="1" number:min-denominator-digits="1"/>'), "?/?"],
    ["fraction, two-digit terms", numberStyle('<number:fraction number:min-integer-digits="0" number:min-numerator-digits="2" number:min-denominator-digits="2"/>'), "# ??/??"],
    ["fraction, denominator-value", numberStyle('<number:fraction number:min-integer-digits="0" number:min-numerator-digits="1" number:denominator-value="8"/>'), "# ?/8"],
    ["fraction, denominator-value, no integer part", numberStyle('<number:fraction number:min-numerator-digits="1" number:denominator-value="16"/>'), "?/16"],
    // percentage, currency, boolean, text
    ["percentage-style", dataStyle("percentage-style", TWO_DECIMALS + text("%")), "0.00%"],
    ["percentage-style without its % text", dataStyle("percentage-style", INTEGER), "0%"],
    [
      "currency-style, de/DE after the number",
      dataStyle("currency-style", `${'<number:number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1" number:grouping="true"/>'}${text(" ")}<number:currency-symbol number:language="de" number:country="DE">€</number:currency-symbol>`),
      "#,##0.00 [$€-407]",
    ],
    [
      "currency-style, en/US before the number",
      dataStyle("currency-style", '<number:currency-symbol number:language="en" number:country="US">$</number:currency-symbol><number:number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1" number:grouping="true"/>'),
      "[$$-409]#,##0.00",
    ],
    [
      "currency-style, a pair that maps to no known LCID",
      dataStyle("currency-style", `<number:currency-symbol number:language="nl" number:country="NL">€</number:currency-symbol>${text(" ")}${TWO_DECIMALS}`),
      "[$€] 0.00",
    ],
    ["currency-style without language/country", dataStyle("currency-style", `<number:currency-symbol>$</number:currency-symbol>${INTEGER}`), "[$$]0"],
    ["boolean-style", dataStyle("boolean-style", "<number:boolean/>"), '"TRUE";"TRUE";"FALSE"'],
    ["text-style", dataStyle("text-style", "<number:text-content/>"), "@"],
    ["text-style with text", dataStyle("text-style", `${text("Name: ")}<number:text-content/>`), '"Name: "@'],
    // number:text, bare or quoted
    ["number:text ( ) in a number section", numberStyle(`${text("(")}${INTEGER}${text(")")}`), "(0)"],
    ["number:text - $ + and spaces in a number section", numberStyle(`${text("- $+ ")}${INTEGER}`), "- $+ 0"],
    ["number:text with letters in a number section", numberStyle(`${INTEGER}${text(" units")}`), '0" units"'],
    ["number:text / in a number section", numberStyle(`${INTEGER}${text("/")}`), '0"/"'],
    ['number:text with a quote → \\"', numberStyle(`${INTEGER}${text('say "hi"')}`), '0"say \\"hi\\""'],
    ["number:text - / : . , ( ) and spaces in a date section", dateStyle(`<number:day/>${text("-/:.,() ")}<number:month/>`), "d-/:.,() m"],
    ["number:text with letters in a date section", dateStyle(`<number:day/>${text(" at ")}<number:month/>`), 'd" at "m'],
    // number:fill-character
    ["fill-character before the number", numberStyle(`<number:fill-character> </number:fill-character>${INTEGER}`), "* 0"],
    ["fill-character after the number", numberStyle(`${INTEGER}<number:fill-character>-</number:fill-character>`), "0*-"],
    // date and time parts
    [
      "year, month, day — long",
      dateStyle(`<number:year number:style="long"/>${text("-")}<number:month number:style="long"/>${text("-")}<number:day number:style="long"/>`),
      "yyyy-mm-dd",
    ],
    ["year, month, day — short", dateStyle(`<number:day/>${text("/")}<number:month/>${text("/")}<number:year/>`), "d/m/yy"],
    ["textual months", dateStyle(`<number:month number:textual="true"/>${text(" ")}<number:month number:textual="true" number:style="long"/>`), "mmm mmmm"],
    ["day-of-week", dateStyle(`<number:day-of-week/>${text(" ")}<number:day-of-week number:style="long"/>`), "ddd dddd"],
    ["quarter", dateStyle(`<number:quarter/>${text(" ")}<number:quarter number:style="long"/>${text(" ")}<number:year number:style="long"/>`), "Q QQ yyyy"],
    ["week-of-year", dateStyle("<number:week-of-year/>"), "WW"],
    ["hours, minutes, seconds — short", timeStyle(`<number:hours/>${text(":")}<number:minutes/>${text(":")}<number:seconds/>`), "h:m:s"],
    [
      "hours, minutes, seconds — long, decimal-places, am-pm",
      timeStyle(`<number:hours number:style="long"/>${text(":")}<number:minutes number:style="long"/>${text(":")}<number:seconds number:style="long" number:decimal-places="2"/>${text(" ")}<number:am-pm/>`),
      "hh:mm:ss.00 AM/PM",
    ],
    ["a date and a time", dateStyle(`<number:day/>${text(".")}<number:month/>${text(" ")}<number:hours/>${text(":")}<number:minutes number:style="long"/>`), "d.m h:mm"],
    // time-style truncate-on-overflow="false" → the first time token bracketed
    [
      "an elapsed [h]",
      timeStyle(`<number:hours/>${text(":")}<number:minutes number:style="long"/>${text(":")}<number:seconds number:style="long"/>`, ' number:truncate-on-overflow="false"'),
      "[h]:mm:ss",
    ],
    ["an elapsed [hh]", timeStyle(`<number:hours number:style="long"/>${text(":")}<number:minutes number:style="long"/>`, ' number:truncate-on-overflow="false"'), "[hh]:mm"],
    ["an elapsed [mm]", timeStyle(`<number:minutes number:style="long"/>${text(":")}<number:seconds number:style="long"/>`, ' number:truncate-on-overflow="false"'), "[mm]:ss"],
    ["an elapsed [ss]", timeStyle(`<number:seconds number:style="long"/>`, ' number:truncate-on-overflow="false"'), "[ss]"],
    // era and unknown children → General
    ["number:era", dateStyle(`<number:era/>${text(" ")}<number:year/>`), "General"],
    ["an unknown child", numberStyle("<number:frobnicate/>"), "General"],
    ["an unknown child after a known one", numberStyle(`${INTEGER}<number:frobnicate/>`), "General"],
    // fo:color
    ["fo:color of a named colour", numberStyle(`<style:text-properties fo:color="#ff0000"/>${INTEGER}`), "[Red]0"],
    ["fo:color in upper case", numberStyle(`<style:text-properties fo:color="#FF0000"/>${INTEGER}`), "[Red]0"],
    ["fo:color #0000ff", numberStyle(`<style:text-properties fo:color="#0000ff"/>${INTEGER}`), "[Blue]0"],
    ["fo:color of a default-palette colour", numberStyle(`<style:text-properties fo:color="#800000"/>${INTEGER}`), "[Color9]0"],
    ["fo:color of no palette colour", numberStyle(`<style:text-properties fo:color="#123456"/>${INTEGER}`), "0"],
    // number:language/number:country on a style with textual names
    [
      "a textual date's language",
      dateStyle(`<number:day/>${text(". ")}<number:month number:style="long" number:textual="true"/>${text(" ")}<number:year number:style="long"/>`, ' number:language="de" number:country="DE"'),
      "[$-407]d. mmmm yyyy",
    ],
    ["a day-of-week's language", dateStyle('<number:day-of-week number:style="long"/>', ' number:language="fr" number:country="FR"'), "[$-40C]dddd"],
    ["a numeric date's language (no prefix)", dateStyle(`<number:day/>${text(".")}<number:month/>${text(".")}<number:year/>`, ' number:language="de" number:country="DE"'), "d.m.yy"],
    ["a language that maps to no known LCID", dateStyle('<number:day-of-week number:style="long"/>', ' number:language="nl" number:country="NL"'), "dddd"],
  ])("%s", (_label, xml, code) => {
    expect(codeOf(xml)).toBe(code);
  });

  it("maps every default-palette colour to a named colour or a [ColorN] naming that colour", () => {
    const named: Record<string, string> = {
      "#000000": "Black",
      "#0000ff": "Blue",
      "#00ffff": "Cyan",
      "#00ff00": "Green",
      "#ff00ff": "Magenta",
      "#ff0000": "Red",
      "#ffffff": "White",
      "#ffff00": "Yellow",
    };
    for (const hex of INDEXED_COLORS.slice(8)) {
      const code = codeOf(numberStyle(`<style:text-properties fo:color="${hex}"/>${INTEGER}`));
      if (named[hex] !== undefined) {
        expect(code, hex).toBe(`[${named[hex]}]0`);
      } else {
        const match = /^\[Color(\d+)\]0$/.exec(code ?? "");
        expect(match, `${hex} → ${String(code)}`).not.toBeNull();
        const n = Number(match![1]);
        expect(n >= 1 && n <= 56, `${hex} → ${String(code)}`).toBe(true);
        expect(INDEXED_COLORS[n + 7], `${hex} → ${String(code)}`).toBe(hex);
      }
    }
  });

  it("is undefined for an element that is not a data style", () => {
    expect(codeOf('<style:style style:name="ce1" style:family="table-cell"/>')).toBeUndefined();
    expect(codeOf(INTEGER)).toBeUndefined();
  });
});

describe("odsNumberStyleToCode — style:map conditions", () => {
  const P0 = '<number:number-style style:name="N1P0" style:volatile="true">';
  const P1 = '<number:number-style style:name="N1P1" style:volatile="true">';
  const P2 = '<number:number-style style:name="N1P2" style:volatile="true">';
  const END = "</number:number-style>";
  it("reads LibreOffice's two-section sign pattern (value()>=0) as plain sections", () => {
    const xml =
      `${P0}${TWO_DECIMALS}${END}` +
      `<number:number-style style:name="N1"><style:text-properties fo:color="#FF0000"/>${text("-")}${TWO_DECIMALS}${map("value()&gt;=0", "N1P0")}${END}`;
    expect(codeOf(xml, "N1")).toBe("0.00;[Red]-0.00");
  });

  it("reads LibreOffice's three-section sign pattern (value()>0, value()<0) as plain sections", () => {
    const xml =
      `${P0}${INTEGER}${END}${P1}${text("-")}${INTEGER}${END}` +
      `<number:number-style style:name="N1">${text("zero")}${map("value()&gt;0", "N1P0")}${map("value()&lt;0", "N1P1")}${END}`;
    expect(codeOf(xml, "N1")).toBe('0;-0;"zero"');
  });

  it("reads a text main style with value()>0, <0 and =0 maps as four plain sections", () => {
    const xml =
      `${P0}${INTEGER}${END}${P1}${text("-")}${INTEGER}${END}${P2}${text("zero")}${END}` +
      `<number:text-style style:name="N1">${text("text: ")}<number:text-content/>` +
      `${map("value()&gt;0", "N1P0")}${map("value()&lt;0", "N1P1")}${map("value()=0", "N1P2")}</number:text-style>`;
    expect(codeOf(xml, "N1")).toBe('0;-0;"zero";"text: "@');
  });

  it("reads any other condition as a [OPn] section, the main style last", () => {
    const conditioned =
      `${P0}<style:text-properties fo:color="#ff0000"/>${INTEGER}${END}` +
      `<number:number-style style:name="N1"><style:text-properties fo:color="#0000ff"/>${INTEGER}${map("value()&gt;100", "N1P0")}${END}`;
    expect(codeOf(conditioned, "N1")).toBe("[>100][Red]0;[Blue]0");
    const two =
      `${P0}${text("big")}${END}${P1}${text("small")}${END}` +
      `<number:number-style style:name="N1">${INTEGER}${map("value()&gt;100", "N1P0")}${map("value()&lt;-100", "N1P1")}${END}`;
    expect(codeOf(two, "N1")).toBe('[>100]"big";[<-100]"small";0');
    const others =
      `${P0}${text("low")}${END}<number:number-style style:name="N1">${INTEGER}${map("value()&lt;=5", "N1P0")}${END}`;
    expect(codeOf(others, "N1")).toBe('[<=5]"low";0');
  });

  it("finds the mapped styles through the lookup, wherever they are", () => {
    const xml =
      `<number:number-style style:name="N1">${text("-")}${TWO_DECIMALS}${map("value()&gt;=0", "N1P0")}${END}${P0}${TWO_DECIMALS}${END}`;
    expect(codeOf(xml, "N1")).toBe("0.00;-0.00");
  });
});

/** The number-namespace attributes of an element, by local name. */
function numberAttrs(el: Element | undefined): Record<string, string> {
  const out: Record<string, string> = {};
  for (const attr of Array.from(el?.attributes ?? [])) {
    if (attr.namespaceURI === NUMBER_NS) {
      out[attr.localName] = attr.value;
    }
  }
  return out;
}

/** The number-namespace children of a style: [local name, element]. */
function parts(el: Element): Array<[string, Element]> {
  return Array.from(el.children)
    .filter((child) => child.namespaceURI === NUMBER_NS)
    .map((child): [string, Element] => [child.localName, child]);
}

/** The one style odsNumberStyleXml writes for a one-section code. */
function writtenStyle(code: string): Element {
  const styles = parse(odsNumberStyleXml(code, "N1"));
  expect(styles, code).toHaveLength(1);
  expect(styles[0]!.getAttributeNS(STYLE_NS, "name")).toBe("N1");
  return styles[0]!;
}

/** The first number:<local> child of a style. */
function part(el: Element, local: string): Element {
  const found = parts(el).find(([name]) => name === local);
  if (!found) {
    throw new Error(`no number:${local} in ${el.outerHTML}`);
  }
  return found[1];
}

describe("odsNumberStyleXml — each row the other way", () => {
  it("writes General as a number:number without decimal-places", () => {
    const style = writtenStyle("General");
    expect(style.localName).toBe("number-style");
    expect(numberAttrs(part(style, "number"))["decimal-places"]).toBeUndefined();
  });

  it("writes decimals, minimum decimals and integer digits", () => {
    expect(numberAttrs(part(writtenStyle("0.00"), "number"))).toMatchObject({ "decimal-places": "2", "min-decimal-places": "2", "min-integer-digits": "1" });
    expect(numberAttrs(part(writtenStyle("#.0##"), "number"))).toMatchObject({ "decimal-places": "3", "min-decimal-places": "1", "min-integer-digits": "0" });
    expect(numberAttrs(part(writtenStyle("000"), "number"))).toMatchObject({ "decimal-places": "0", "min-integer-digits": "3" });
    expect(numberAttrs(part(writtenStyle("0.00"), "number")).grouping).toBeUndefined();
  });

  it("writes grouping and the display factor", () => {
    expect(numberAttrs(part(writtenStyle("#,##0"), "number"))).toMatchObject({ grouping: "true", "min-integer-digits": "1" });
    expect(numberAttrs(part(writtenStyle("#,##0,"), "number"))).toMatchObject({ grouping: "true", "display-factor": "1000" });
    expect(numberAttrs(part(writtenStyle("0,,"), "number"))["display-factor"]).toBe("1000000");
  });

  it("writes literals between integer placeholders as embedded text", () => {
    const number = part(writtenStyle("000-00-0000"), "number");
    expect(numberAttrs(number)["min-integer-digits"]).toBe("9");
    const embedded = parts(number)
      .filter(([name]) => name === "embedded-text")
      .map(([, el]) => [Number(numberAttrs(el).position), el.textContent])
      .toSorted((a, b) => Number(a[0]) - Number(b[0]));
    expect(embedded).toEqual([
      [4, "-"],
      [6, "-"],
    ]);
  });

  it("writes scientific-number with its exponent digits and interval", () => {
    expect(numberAttrs(part(writtenStyle("0.00E+00"), "scientific-number"))).toMatchObject({
      "decimal-places": "2",
      "min-integer-digits": "1",
      "min-exponent-digits": "2",
    });
    expect(numberAttrs(part(writtenStyle("##0.0E+0"), "scientific-number"))).toMatchObject({
      "decimal-places": "1",
      "min-integer-digits": "1",
      "min-exponent-digits": "1",
      "exponent-interval": "3",
    });
  });

  it("writes fractions with their integer part, numerator and denominator", () => {
    expect(numberAttrs(part(writtenStyle("# ?/?"), "fraction"))).toMatchObject({
      "min-integer-digits": "0",
      "min-numerator-digits": "1",
      "min-denominator-digits": "1",
    });
    const fixed = numberAttrs(part(writtenStyle("?/8"), "fraction"));
    expect(fixed["denominator-value"]).toBe("8");
    expect(fixed["min-integer-digits"]).toBeUndefined();
  });

  it("writes percentage-style, currency-style and text-style", () => {
    const percent = writtenStyle("0.00%");
    expect(percent.localName).toBe("percentage-style");
    expect(parts(percent).map(([name, el]) => (name === "text" ? `text:${el.textContent}` : name))).toEqual(["number", "text:%"]);

    const euro = writtenStyle("[$€-407]#,##0.00");
    expect(euro.localName).toBe("currency-style");
    const symbol = part(euro, "currency-symbol");
    expect(symbol.textContent).toBe("€");
    expect(numberAttrs(symbol)).toMatchObject({ language: "de", country: "DE" });
    expect(numberAttrs(part(writtenStyle("[$€]0"), "currency-symbol"))).toEqual({});

    const textStyle = writtenStyle('"Name: "@');
    expect(textStyle.localName).toBe("text-style");
    expect(parts(textStyle).map(([name, el]) => (name === "text" ? `text:${el.textContent}` : name))).toEqual(["text:Name: ", "text-content"]);
    expect(parts(writtenStyle("@")).map(([name]) => name)).toEqual(["text-content"]);
  });

  it("writes literals as number:text and *x as number:fill-character", () => {
    expect(parts(writtenStyle('0" units"')).map(([name, el]) => (name === "text" ? `text:${el.textContent}` : name))).toEqual(["number", "text: units"]);
    const filled = parts(writtenStyle("* 0"));
    expect(filled.map(([name]) => name)).toEqual(["fill-character", "number"]);
    expect(filled[0]![1].textContent).toBe(" ");
    expect(odsNumberStyleXml('"a<b&c"0', "N1")).toContain("<number:text>a&lt;b&amp;c</number:text>");
  });

  it.each<[code: string, expected: string[]]>([
    ["yyyy-mm-dd", ["year long", "text:-", "month long", "text:-", "day long"]],
    ["d/m/yy", ["day", "text:/", "month", "text:/", "year"]],
    ["mmm mmmm", ["month textual", "text: ", "month long textual"]],
    ["ddd dddd", ["day-of-week", "text: ", "day-of-week long"]],
    ["Q QQ yyyy", ["quarter", "text: ", "quarter long", "text: ", "year long"]],
    ["WW", ["week-of-year"]],
  ])("writes the date parts of %j", (code, expected) => {
    const style = writtenStyle(code);
    expect(style.localName).toBe("date-style");
    expect(describeParts(style)).toEqual(expected);
  });

  it("writes the time parts, fractional seconds and AM/PM", () => {
    const style = writtenStyle("h:mm:ss.00 AM/PM");
    expect(style.localName).toBe("time-style");
    expect(describeParts(style)).toEqual(["hours", "text::", "minutes long", "text::", "seconds long", "text: ", "am-pm"]);
    expect(numberAttrs(part(style, "seconds"))["decimal-places"]).toBe("2");
  });

  it("writes an elapsed time as a time-style with truncate-on-overflow false", () => {
    const elapsed = writtenStyle("[h]:mm:ss");
    expect(elapsed.localName).toBe("time-style");
    expect(numberAttrs(elapsed)["truncate-on-overflow"]).toBe("false");
    expect(numberAttrs(writtenStyle("h:mm"))["truncate-on-overflow"]).toBeUndefined();
  });

  it("writes a colour as style:text-properties fo:color", () => {
    const colorOf = (code: string): string | null =>
      Array.from(writtenStyle(code).children)
        .find((child) => child.namespaceURI === STYLE_NS && child.localName === "text-properties")
        ?.getAttributeNS(FO_NS, "color") ?? null;
    expect(colorOf("[Red]0")).toBe("#ff0000");
    expect(colorOf("[Color10]0")).toBe(INDEXED_COLORS[17]);
    expect(colorOf("0")).toBeNull();
  });

  it("writes a section LCID as the date style's language and country", () => {
    expect(numberAttrs(writtenStyle("[$-407]d. mmmm yyyy"))).toMatchObject({ language: "de", country: "DE" });
    expect(numberAttrs(writtenStyle("[$-40C]dddd"))).toMatchObject({ language: "fr", country: "FR" });
  });

  it.each<[code: string, kind: string]>([
    ["yyyy-mm-dd", "date-style"],
    ["d-mmm-yy h:mm", "date-style"],
    ["[$-F800]dddd, mmmm dd, yyyy", "date-style"],
    ["h:mm:ss", "time-style"],
    ["[h]:mm:ss", "time-style"],
    ["0.00%", "percentage-style"],
    ["[$€-407]#,##0.00", "currency-style"],
    ["@", "text-style"],
    ["#,##0.00", "number-style"],
    ["0.00E+00", "number-style"],
    ["# ?/?", "number-style"],
    ['"text"', "number-style"],
  ])("writes %j as a number:%s", (code, kind) => {
    expect(writtenStyle(code).localName).toBe(kind);
  });
});

/** A style's number children as "name [long] [textual]" / "text:…" strings. */
function describeParts(style: Element): string[] {
  return parts(style).map(([name, el]) => {
    if (name === "text") {
      return `text:${el.textContent}`;
    }
    const attrs = numberAttrs(el);
    return [name, attrs.style === "long" ? "long" : "", attrs.textual === "true" ? "textual" : ""].filter(Boolean).join(" ");
  });
}

describe("odsNumberStyleXml — sections and maps", () => {
  /** [name, element kind, volatile, maps as "condition→style"]. */
  function shape(xml: string): Array<[string | null, string, string | null, string[]]> {
    return parse(xml).map((el) => [
      el.getAttributeNS(STYLE_NS, "name"),
      el.localName,
      el.getAttributeNS(STYLE_NS, "volatile"),
      Array.from(el.children)
        .filter((child) => child.namespaceURI === STYLE_NS && child.localName === "map")
        .map((styleMap) => `${styleMap.getAttributeNS(STYLE_NS, "condition")}→${styleMap.getAttributeNS(STYLE_NS, "apply-style-name")}`),
    ]);
  }

  it("writes one section as the main style alone", () => {
    expect(shape(odsNumberStyleXml("0.00", "N1"))).toEqual([["N1", "number-style", null, []]]);
  });

  it("writes an unconditioned two-section code as ${name}P0 mapped by value()>=0", () => {
    expect(shape(odsNumberStyleXml("0.00;[Red]-0.00", "N1"))).toEqual([
      ["N1P0", "number-style", "true", []],
      ["N1", "number-style", null, ["value()>=0→N1P0"]],
    ]);
  });

  it("writes an unconditioned three-section code as P0 (value()>0) and P1 (value()<0), the zero section main", () => {
    expect(shape(odsNumberStyleXml('0;-0;"zero"', "N1"))).toEqual([
      ["N1P0", "number-style", "true", []],
      ["N1P1", "number-style", "true", []],
      ["N1", "number-style", null, ["value()>0→N1P0", "value()<0→N1P1"]],
    ]);
  });

  it("writes a text section as the main style, every numeric section mapped from it", () => {
    expect(shape(odsNumberStyleXml('0;-0;"zero";"text: "@', "N1"))).toEqual([
      ["N1P0", "number-style", "true", []],
      ["N1P1", "number-style", "true", []],
      ["N1P2", "number-style", "true", []],
      ["N1", "text-style", null, ["value()>0→N1P0", "value()<0→N1P1", "value()=0→N1P2"]],
    ]);
  });

  it("writes a conditioned section with its own condition", () => {
    expect(shape(odsNumberStyleXml("[>100][Red]0;[Blue]0", "N1"))).toEqual([
      ["N1P0", "number-style", "true", []],
      ["N1", "number-style", null, ["value()>100→N1P0"]],
    ]);
  });

  it("names the sub-styles after the main style", () => {
    expect(shape(odsNumberStyleXml('0;-0;"zero"', "N42")).map(([name]) => name)).toEqual(["N42P0", "N42P1", "N42"]);
  });

  it("escapes the condition text", () => {
    const xml = odsNumberStyleXml('0;-0;"zero"', "N1");
    expect(xml).toContain('style:condition="value()&gt;0"');
    expect(xml).toContain('style:condition="value()&lt;0"');
    expect(xml).not.toContain("value()<0");
  });
});

/** The round-trip corpus: the locale-independent built-ins plus the
 * Reference's extra codes. */
const CORPUS: string[] = [
  ...Object.entries(BUILTIN_NUMBER_FORMATS)
    .filter(([id]) => !LOCALE_DEPENDENT_BUILTINS.includes(Number(id)))
    .map(([, code]) => code),
  "0.00;[Red]-0.00",
  '0;-0;"zero";"text: "@',
  '#,##0,"K"',
  "[h]:mm:ss",
  "dddd d mmmm yyyy",
  "# ?/8",
  "0.00E+00",
  "[$€-407]#,##0.00",
];

describe("the corpus round trip", () => {
  it("covers the 26 locale-independent built-ins and the eight extra codes", () => {
    expect(CORPUS).toHaveLength(34);
    expect(CORPUS).toContain("General");
    expect(CORPUS).toContain("@");
    expect(CORPUS).not.toContain("m/d/yyyy");
  });

  it.each(CORPUS)("odsNumberStyleToCode(parse(odsNumberStyleXml(%j, \"N1\"))) is the code", (code) => {
    expect(codeOf(odsNumberStyleXml(code, "N1"), "N1")).toBe(code);
  });
});
