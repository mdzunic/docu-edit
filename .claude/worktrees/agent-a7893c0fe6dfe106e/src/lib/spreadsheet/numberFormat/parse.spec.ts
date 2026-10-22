// SPEC-118 "The grammar": parseFormatCode reads Excel's format-code grammar —
// up to four `;`-separated sections; the bracket forms (colours, conditions
// with each operator, `[$sym-lcid]` currency/locale tags incl. the system
// long date/time words, elapsed units); the body tokens (the placeholders
// `0 # ?`, `.` `,` `%`, the four exponents, fractions incl. fixed
// denominators, `@`, literals in their four spellings, `General`) and every
// date/time token; the minute rule in both directions; each FormatCodeError
// kind with its offset; every unsupported calendar or numeral system
// (`{ unsupported: true }`, never a throw); unknown bare letters as lenient
// literals; and the parse cache's limit (FORMAT_CACHE_LIMIT, cleared whole
// on overflow).
import { describe, expect, it } from "vitest";
import {
  FORMAT_CACHE_LIMIT,
  FormatCodeError,
  formatCacheSize,
  LCID_SYSTEM_LONG_DATE,
  LCID_SYSTEM_LONG_TIME,
  parseFormatCode,
  type FormatSection,
  type FormatToken,
} from "./parse";

/** The one section of a one-section code. */
function only(code: string): FormatSection {
  const { sections } = parseFormatCode(code);
  expect(sections, code).toHaveLength(1);
  return sections[0]!;
}

/** The FormatCodeError an invalid code throws. */
function errorOf(code: string): FormatCodeError {
  try {
    parseFormatCode(code);
  } catch (error) {
    if (error instanceof FormatCodeError) {
      return error;
    }
    throw error;
  }
  throw new Error(`${JSON.stringify(code)} parsed without an error`);
}

/** The calendar parts a one-section code's date/time tokens show, in order. */
function dateParts(code: string): string[] {
  return only(code)
    .tokens.filter((token): token is Extract<FormatToken, { type: "date" }> => token.type === "date")
    .map((token) => token.part);
}

/** The literal text of a section, its literal tokens concatenated. */
function literalText(section: FormatSection): string {
  return section.tokens
    .filter((token): token is Extract<FormatToken, { type: "literal" }> => token.type === "literal")
    .map((token) => token.text)
    .join("");
}

describe("sections", () => {
  it.each<[code: string, count: number]>([
    ["0", 1],
    ["0;-0", 2],
    ['0;-0;"zero"', 3],
    ['0;-0;"zero";@', 4],
  ])("%j has %i section(s)", (code, count) => {
    expect(parseFormatCode(code).sections).toHaveLength(count);
  });

  it("reads an empty section as a section with no tokens (`;;;` is four of them)", () => {
    const { sections } = parseFormatCode(";;;");
    expect(sections).toHaveLength(4);
    for (const section of sections) {
      expect(section.tokens).toEqual([]);
      expect(section.kind).toBe("literal");
    }
  });

  it.each<[code: string, kind: FormatSection["kind"]]>([
    ["0.00", "number"],
    ["#,##0", "number"],
    ["0%", "number"],
    ["# ?/?", "fraction"],
    ["?/8", "fraction"],
    ["0.00E+00", "scientific"],
    ["yyyy-mm-dd", "date"],
    ["h:mm", "date"],
    ["[h]:mm:ss", "date"],
    ["@", "text"],
    ['"Name: "@', "text"],
    ["General", "general"],
    ['"zero"', "literal"],
  ])("classifies %j as a %s section", (code, kind) => {
    expect(only(code).kind).toBe(kind);
  });

  it("classifies each section of a code on its own", () => {
    expect(parseFormatCode('#,##0.00;[Red]-# ?/?;"zero";"text: "@').sections.map((section) => section.kind)).toEqual([
      "number",
      "fraction",
      "literal",
      "text",
    ]);
  });
});

describe("brackets", () => {
  it.each<[code: string, op: string, value: number]>([
    ["[<10]0", "<", 10],
    ["[<=10]0", "<=", 10],
    ["[>10]0", ">", 10],
    ["[>=10]0", ">=", 10],
    ["[=10]0", "=", 10],
    ["[<>10]0", "<>", 10],
    ["[<-1.5]0", "<", -1.5],
    ["[>=0.5]0", ">=", 0.5],
    ["[>100][Red]0", ">", 100],
  ])("reads the condition of %j as %s %d", (code, op, value) => {
    expect(only(code).condition).toEqual({ op, value });
  });

  it("keeps each section's condition and colour with that section", () => {
    const [first, second] = parseFormatCode("[>100][Red]0;[Blue]0").sections;
    expect(first).toMatchObject({ condition: { op: ">", value: 100 }, color: "Red" });
    expect(second!.condition).toBeUndefined();
    expect(second!.color).toBe("Blue");
  });

  it("has no condition or colour on a plain section", () => {
    const section = only("0.00");
    expect(section.condition).toBeUndefined();
    expect(section.color).toBeUndefined();
    expect(section.lcid).toBeUndefined();
  });

  it.each(["Black", "Blue", "Cyan", "Green", "Magenta", "Red", "White", "Yellow"])(
    "reads [%s] case-insensitively as its canonical name",
    (name) => {
      for (const written of [name, name.toLowerCase(), name.toUpperCase()]) {
        expect(only(`[${written}]0`).color, written).toBe(name);
      }
    },
  );

  it("reads [Color1] … [Color56] (case-insensitive) as ColorN", () => {
    for (let n = 1; n <= 56; n += 1) {
      expect(only(`[Color${n}]0`).color).toBe(`Color${n}`);
    }
    expect(only("[color10]0").color).toBe("Color10");
    expect(only("[COLOR10]0").color).toBe("Color10");
  });

  it("reads [$sym-lcid] as a currency symbol in the body and the section's LCID", () => {
    const euro = only("[$€-407]#,##0.00");
    expect(euro.lcid).toBe(0x407);
    expect(euro.tokens[0]).toEqual({ type: "currency", symbol: "€" });
    expect(euro.kind).toBe("number");
    const dollar = only("[$$-409]0");
    expect(dollar.lcid).toBe(0x409);
    expect(dollar.tokens[0]).toEqual({ type: "currency", symbol: "$" });
    expect(only("[$USD-409] 0").tokens[0]).toEqual({ type: "currency", symbol: "USD" });
  });

  it("reads [$sym] without an LCID", () => {
    const section = only("[$€]0");
    expect(section.lcid).toBeUndefined();
    expect(section.tokens[0]).toEqual({ type: "currency", symbol: "€" });
  });

  it("reads [$-lcid] as the section's LCID with no symbol (1–8 hex digits, any case)", () => {
    const german = only("[$-407]mmmm");
    expect(german.lcid).toBe(0x407);
    expect(german.tokens.some((token) => token.type === "currency")).toBe(false);
    expect(only("[$-c0a]mmmm").lcid).toBe(0xc0a);
    expect(only("[$-C0A]mmmm").lcid).toBe(0xc0a);
    expect(only("[$-40A]mmmm").lcid).toBe(0x40a);
    expect(only("[$-00000409]mmmm").lcid).toBe(0x409);
  });

  it("reads [$-F800]/[$-x-sysdate] as the system long date and [$-F400]/[$-x-systime] as the system long time", () => {
    expect(LCID_SYSTEM_LONG_DATE).toBe(0xf800);
    expect(LCID_SYSTEM_LONG_TIME).toBe(0xf400);
    expect(only("[$-F800]dddd, mmmm dd, yyyy").lcid).toBe(LCID_SYSTEM_LONG_DATE);
    expect(only("[$-x-sysdate]dddd, mmmm dd, yyyy").lcid).toBe(LCID_SYSTEM_LONG_DATE);
    expect(only("[$-F400]h:mm:ss AM/PM").lcid).toBe(LCID_SYSTEM_LONG_TIME);
    expect(only("[$-x-systime]h:mm:ss AM/PM").lcid).toBe(LCID_SYSTEM_LONG_TIME);
    expect(parseFormatCode("[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy").unsupported).toBeUndefined();
  });

  it.each<[code: string, unit: "h" | "m" | "s", length: number]>([
    ["[h]:mm:ss", "h", 1],
    ["[hh]:mm", "h", 2],
    ["[H]:mm", "h", 1],
    ["[m]:ss", "m", 1],
    ["[mm]:ss", "m", 2],
    ["[MM]:ss", "m", 2],
    ["[s]", "s", 1],
    ["[ss]", "s", 2],
    ["[SS]", "s", 2],
  ])("reads %j's bracket as the elapsed %s (length %i) of a date section", (code, unit, length) => {
    const section = only(code);
    expect(section.tokens[0]).toEqual({ type: "elapsed", unit, length });
    expect(section.kind).toBe("date");
  });
});

describe("body tokens", () => {
  it("reads the digit placeholders 0 # ? in order", () => {
    expect(only("0#?").tokens).toEqual([
      { type: "digit", placeholder: "0" },
      { type: "digit", placeholder: "#" },
      { type: "digit", placeholder: "?" },
    ]);
  });

  it("reads `.`, `,` and `%` as the point, the comma and the percent sign", () => {
    expect(only("#,##0.0%").tokens).toEqual([
      { type: "digit", placeholder: "#" },
      { type: "comma" },
      { type: "digit", placeholder: "#" },
      { type: "digit", placeholder: "#" },
      { type: "digit", placeholder: "0" },
      { type: "point" },
      { type: "digit", placeholder: "0" },
      { type: "percent" },
    ]);
  });

  it.each(["E+", "E-", "e+", "e-"])("reads the exponent %s as written, making a scientific section", (exponent) => {
    const section = only(`0.00${exponent}00`);
    expect(section.kind).toBe("scientific");
    expect(section.tokens).toContainEqual({ type: "exponent", text: exponent });
  });

  it("reads a fraction: placeholders either side of a `/`", () => {
    expect(only("# ?/?").tokens).toEqual([
      { type: "digit", placeholder: "#" },
      { type: "literal", text: " " },
      { type: "digit", placeholder: "?" },
      { type: "slash" },
      { type: "digit", placeholder: "?" },
    ]);
  });

  it("reads a fixed denominator written as digits", () => {
    expect(only("?/8").tokens).toEqual([{ type: "digit", placeholder: "?" }, { type: "slash" }, { type: "denominator", value: 8, text: "8" }]);
    const sixteenths = only("# ??/16");
    expect(sixteenths.kind).toBe("fraction");
    expect(sixteenths.tokens.at(-1)).toEqual({ type: "denominator", value: 16, text: "16" });
  });

  it("reads @ as the text placeholder", () => {
    expect(only("@").tokens).toEqual([{ type: "text" }]);
    expect(only('"Name: "@').tokens).toEqual([{ type: "literal", text: "Name: " }, { type: "text" }]);
  });

  it('reads a quoted literal, with \\" as a quote inside the quotes', () => {
    expect(only('"abc"0').tokens).toEqual([{ type: "literal", text: "abc" }, { type: "digit", placeholder: "0" }]);
    expect(only('0"say \\"hi\\""').tokens).toEqual([{ type: "digit", placeholder: "0" }, { type: "literal", text: 'say "hi"' }]);
    // Quoted placeholders and separators are literal text.
    expect(only('0" 0.#,%"').tokens).toEqual([{ type: "digit", placeholder: "0" }, { type: "literal", text: " 0.#,%" }]);
  });

  it("reads \\x as the literal x, even a placeholder", () => {
    expect(only("0\\x").tokens).toEqual([{ type: "digit", placeholder: "0" }, { type: "literal", text: "x" }]);
    expect(only("\\00").tokens).toEqual([{ type: "literal", text: "0" }, { type: "digit", placeholder: "0" }]);
  });

  it("reads _x as a space as wide as x", () => {
    expect(only("_(0_)").tokens).toEqual([
      { type: "space", char: "(" },
      { type: "digit", placeholder: "0" },
      { type: "space", char: ")" },
    ]);
  });

  it("reads *x as a fill of x", () => {
    expect(only("*-0").tokens).toEqual([{ type: "fill", char: "-" }, { type: "digit", placeholder: "0" }]);
    expect(only('"$"* 0').tokens).toEqual([{ type: "literal", text: "$" }, { type: "fill", char: " " }, { type: "digit", placeholder: "0" }]);
  });

  it("keeps only the last *x of a section", () => {
    expect(only("*a0*b").tokens).toEqual([{ type: "digit", placeholder: "0" }, { type: "fill", char: "b" }]);
    expect(only("*a*b0*c").tokens.filter((token) => token.type === "fill")).toEqual([{ type: "fill", char: "c" }]);
    // Per section: each section keeps its own last fill.
    const [first, second] = parseFormatCode("*a0;*b0").sections;
    expect(first!.tokens[0]).toEqual({ type: "fill", char: "a" });
    expect(second!.tokens[0]).toEqual({ type: "fill", char: "b" });
  });

  it("reads General as the General token", () => {
    expect(only("General")).toEqual({ kind: "general", tokens: [{ type: "general" }] });
    expect(only("[Red]General")).toEqual({ kind: "general", color: "Red", tokens: [{ type: "general" }] });
    expect(only('"Total: "General').tokens).toEqual([{ type: "literal", text: "Total: " }, { type: "general" }]);
  });

  it("reads any other character as a literal", () => {
    const section = only("$-+():!^&'~{}<>=0");
    expect(section.kind).toBe("number");
    expect(literalText(section)).toBe("$-+():!^&'~{}<>=");
    expect(section.tokens.at(-1)).toEqual({ type: "digit", placeholder: "0" });
  });

  it("reads unknown bare letters as lenient literals, not errors", () => {
    for (const code of ["0 lux", "0 x", "#,##0 Zl", "W"]) {
      expect(() => parseFormatCode(code), code).not.toThrow();
      expect(parseFormatCode(code).unsupported, code).toBeUndefined();
    }
    const lux = only("0 lux");
    expect(lux.kind).toBe("number");
    expect(literalText(lux)).toBe(" lux");
    // A lone W is a letter; only WW is the ISO week.
    expect(only("W")).toMatchObject({ kind: "literal" });
    expect(literalText(only("W"))).toBe("W");
  });
});

describe("date and time tokens", () => {
  it.each<[code: string, part: string, length: number]>([
    ["y", "year", 1],
    ["yy", "year", 2],
    ["yyy", "year", 3],
    ["yyyy", "year", 4],
    ["yyyyy", "year", 5],
    ["m", "month", 1],
    ["mm", "month", 2],
    ["mmm", "month", 3],
    ["mmmm", "month", 4],
    ["mmmmm", "month", 5],
    ["mmmmmm", "month", 6],
    ["d", "day", 1],
    ["dd", "day", 2],
    ["ddd", "day", 3],
    ["dddd", "day", 4],
    ["ddddd", "day", 5],
    ["h", "hour", 1],
    ["hh", "hour", 2],
    ["hhh", "hour", 3],
    ["s", "second", 1],
    ["ss", "second", 2],
    ["sss", "second", 3],
    ["Q", "quarter", 1],
    ["QQ", "quarter", 2],
    ["WW", "week", 2],
    // Case-insensitive.
    ["YYYY", "year", 4],
    ["MMM", "month", 3],
    ["DDDD", "day", 4],
    ["HH", "hour", 2],
    ["SS", "second", 2],
    ["q", "quarter", 1],
    ["ww", "week", 2],
  ])("reads %j as the %s token of length %i in a date section", (code, part, length) => {
    const section = only(code);
    expect(section.kind).toBe("date");
    expect(section.tokens).toEqual([{ type: "date", part, length }]);
  });

  it.each<[code: string, digits: number]>([
    ["ss.0", 1],
    ["ss.00", 2],
    ["ss.000", 3],
    ["s.0", 1],
  ])("reads the zeros of %j after the seconds as %i fractional-second digit(s)", (code, digits) => {
    expect(only(code).tokens).toEqual([{ type: "date", part: "second", length: code.indexOf(".") }, { type: "subsecond", digits }]);
  });

  it("reads fractional seconds after a time's seconds and after an elapsed [ss]", () => {
    expect(only("h:mm:ss.00").tokens.at(-1)).toEqual({ type: "subsecond", digits: 2 });
    expect(only("[ss].000").tokens).toEqual([{ type: "elapsed", unit: "s", length: 2 }, { type: "subsecond", digits: 3 }]);
  });

  it("reads `.0` anywhere but after a seconds token as literal text", () => {
    const section = only("hh.00");
    expect(section.tokens.some((token) => token.type === "subsecond")).toBe(false);
    expect(literalText(section)).toBe(".00");
  });

  it.each<[code: string, am: string, pm: string]>([
    ["AM/PM", "AM", "PM"],
    ["am/pm", "am", "pm"],
    ["A/P", "A", "P"],
    ["a/p", "a", "p"],
  ])("reads %j as the AM/PM marker in the code's case", (marker, am, pm) => {
    const section = only(`h:mm ${marker}`);
    expect(section.kind).toBe("date");
    expect(section.tokens.at(-1)).toEqual({ type: "ampm", am, pm });
  });

  it("makes a section with any date/time token a date section, its placeholders literals", () => {
    const section = only("yyyy#0?");
    expect(section.kind).toBe("date");
    expect(section.tokens).toEqual([
      { type: "date", part: "year", length: 4 },
      { type: "literal", text: "#" },
      { type: "literal", text: "0" },
      { type: "literal", text: "?" },
    ]);
    expect(only("d 0.0%").tokens.filter((token) => token.type !== "date" && token.type !== "literal")).toEqual([]);
  });

  it("reads `/` and `:` in a date section as literal characters", () => {
    expect(only("m/d/yyyy h:mm").tokens.filter((token) => token.type === "literal")).toEqual([
      { type: "literal", text: "/" },
      { type: "literal", text: "/" },
      { type: "literal", text: " " },
      { type: "literal", text: ":" },
    ]);
  });
});

// "m/mm is minutes when the nearest preceding date/time token is an hour (h,
// hh, [h]) or the nearest following one is a second (s, ss, [s]); otherwise
// it is the month."
describe("the minute rule", () => {
  it.each<[code: string, parts: string[]]>([
    // After an hour.
    ["h:mm", ["hour", "minute"]],
    ["hh:m", ["hour", "minute"]],
    ['h" hours "mm', ["hour", "minute"]],
    ["m/d/yyyy h:mm", ["month", "day", "year", "hour", "minute"]],
    // Before a second.
    ["mm:ss", ["minute", "second"]],
    ["m:s", ["minute", "second"]],
    ["mm:ss.0", ["minute", "second"]],
    ["yyyy mm ss", ["year", "minute", "second"]],
    // Neither: a month.
    ["mm", ["month"]],
    ["m", ["month"]],
    ["dd/mm/yyyy", ["day", "month", "year"]],
    ["d-m-yyyy", ["day", "month", "year"]],
    ["d mm", ["day", "month"]],
    ["yyyy-mm", ["year", "month"]],
    // The nearest preceding token decides, not any earlier hour.
    ["h d mm", ["hour", "day", "month"]],
    // A minute before the seconds, a month after them.
    ["mm:ss mm", ["minute", "second", "month"]],
    // mmm and longer are always the month.
    ["hh mmm", ["hour", "month"]],
    ["mmmm ss", ["month", "second"]],
  ])("%j reads as %j", (code, parts) => {
    expect(dateParts(code)).toEqual(parts);
  });

  it("reads mm after [h] and before [ss] as minutes", () => {
    expect(only("[h]:mm").tokens.at(-1)).toEqual({ type: "date", part: "minute", length: 2 });
    expect(only("mm:[ss]").tokens[0]).toEqual({ type: "date", part: "minute", length: 2 });
    expect(only("[hh]:mm:ss").tokens[2]).toEqual({ type: "date", part: "minute", length: 2 });
  });
});

describe("structural errors", () => {
  it.each<[label: string, code: string, offset: number]>([
    ["an unterminated quote", '0"abc', 1],
    ["an unterminated quote at the start", '"abc', 0],
    ["an unterminated quote in a later section", '0;"x', 2],
    ["an unterminated bracket", "[Red0.00", 0],
    ["an unterminated bracket in a later section", "0.00;[Red", 5],
    ["a fifth section (the fourth `;`)", "0;0;0;@;0", 7],
    ["a fifth empty section", ";;;;", 3],
    ["an unknown bracket", "[Foo]0", 0],
    ["an unknown bracket in a later section", "0;[Bar]0", 2],
    ["[Color0] (colours run 1..56)", "[Color0]0", 0],
    ["[Color57]", "[Color57]0", 0],
    ["a digit placeholder then @", "0@", 1],
    ["@ then a digit placeholder", "@ 0", 2],
    ["digits and @ in a later section", "0;#@", 3],
    ["an E+ with no exponent digits", "0.00E+", 4],
    ["an E- followed by a letter", "0E-x", 1],
    ["an e+ followed by a quoted literal", '0.0e+"x"', 3],
  ])("%s (%j) throws FormatCodeError at offset %i", (_label, code, offset) => {
    const error = errorOf(code);
    expect(error).toBeInstanceOf(FormatCodeError);
    expect(error).toBeInstanceOf(Error);
    expect(error.offset, code).toBe(offset);
  });

  it("throws again for an invalid code it has seen before", () => {
    expect(() => parseFormatCode('0"never closed')).toThrow(FormatCodeError);
    expect(() => parseFormatCode('0"never closed')).toThrow(FormatCodeError);
  });

  it("does not count a quoted or escaped digit as a placeholder next to @", () => {
    expect(only('"0"@').kind).toBe("text");
    expect(only("\\0@").kind).toBe("text");
    expect(only('"E+"0').kind).toBe("number");
  });
});

describe("unsupported calendars and numeral systems", () => {
  it.each([
    // B1/B2 calendar prefixes.
    "B1dd/mm/yyyy",
    "B2yyyy-mm-dd",
    // Era tokens.
    "e",
    "ee",
    "g",
    "gg",
    "ggg",
    'ggge"年"m"月"d"日"',
    "[$-411]ge.m.d",
    // Numeral systems.
    "[DBNum1]General",
    "[DBNum2]0",
    "[DBNum3]#,##0",
    "[DBNum4]0.00",
    "[NatNum1]0",
    "[DBNum1][$-804]General",
    // A calendar or numeral-system byte above 0xFFFF in an LCID.
    "[$-1010409]d/m/yyyy",
    "[$-2010401]dddd",
    "[$-F0000409]d/m/yyyy",
  ])("parses %j as { unsupported: true } without throwing", (code) => {
    expect(() => parseFormatCode(code)).not.toThrow();
    expect(parseFormatCode(code).unsupported).toBe(true);
  });

  it("leaves a supported code without the flag — an LCID of at most 0xFFFF included", () => {
    for (const code of ["0.00", "[$-411]yyyy", "[$-F800]dddd", "[$€-407]0", "yyyy-mm-dd"]) {
      expect(parseFormatCode(code).unsupported, code).toBeUndefined();
    }
  });
});

describe("the parse cache", () => {
  it("returns the same object for a repeated code", () => {
    expect(parseFormatCode("#,##0.00")).toBe(parseFormatCode("#,##0.00"));
  });

  it("holds at most FORMAT_CACHE_LIMIT (1024) codes and clears itself whole when full", () => {
    expect(FORMAT_CACHE_LIMIT).toBe(1024);
    const before = parseFormatCode("0.0000");
    let previous = formatCacheSize();
    let cleared = false;
    for (let i = 0; i <= FORMAT_CACHE_LIMIT; i += 1) {
      parseFormatCode(`0" cache filler ${i}"`);
      const size = formatCacheSize();
      expect(size).toBeLessThanOrEqual(FORMAT_CACHE_LIMIT);
      if (size < previous) {
        // Cleared whole, then the new code cached.
        expect(size).toBe(1);
        cleared = true;
      }
      previous = size;
    }
    expect(cleared).toBe(true);
    expect(formatCacheSize()).toBeLessThanOrEqual(FORMAT_CACHE_LIMIT);
    // The clear dropped the earlier entry: the code parses afresh, equally.
    const after = parseFormatCode("0.0000");
    expect(after).not.toBe(before);
    expect(after).toEqual(before);
  });
});
