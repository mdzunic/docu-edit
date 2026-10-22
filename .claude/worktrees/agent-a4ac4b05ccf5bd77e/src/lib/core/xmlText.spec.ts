import { describe, expect, it } from "vitest";
import { hasIllegalXmlScalars, sanitizeXmlText } from "./xmlText";

// Built from char codes rather than literals: these characters are invisible
// in a source file, and a literal one in a test fixture is unreviewable.
const ch = (code: number) => String.fromCharCode(code);
const NUL = ch(0x00);
const SOH = ch(0x01);
const BACKSPACE = ch(0x08);
const TAB = ch(0x09);
const LF = ch(0x0a);
const VT = ch(0x0b);
const FF = ch(0x0c);
const CR = ch(0x0d);
const SO = ch(0x0e);
const UNIT_SEP = ch(0x1f);
const DEL = ch(0x7f);
const HIGH_SURROGATE = ch(0xd800);
const LOW_SURROGATE = ch(0xdc00);
const NOT_A_CHARACTER = ch(0xfffe);

describe("sanitizeXmlText", () => {
  it("leaves ordinary text alone, and returns the very same string", () => {
    const text = "Quarterly report — 3 × £4.50, naïve café 🎉";
    expect(sanitizeXmlText(text)).toBe(text);
    // Identity, not just equality: the clean path must not allocate.
    expect(sanitizeXmlText(text)).toBe(text);
    expect(hasIllegalXmlScalars(text)).toBe(false);
  });

  it("keeps the three C0 characters XML 1.0 allows", () => {
    const text = `a${TAB}b${LF}c${CR}d`;
    expect(sanitizeXmlText(text)).toBe(text);
    expect(hasIllegalXmlScalars(text)).toBe(false);
  });

  // U+000B is what Word puts in the clipboard's plain-text flavour for a
  // Shift+Enter break, so this is the common case, not an exotic one — and
  // this model spells a line break as a literal "\n" in run text.
  it("translates VT and FF to a line break rather than dropping them", () => {
    expect(sanitizeXmlText(`line one${VT}line two`)).toBe("line one\nline two");
    expect(sanitizeXmlText(`page one${FF}page two`)).toBe("page one\npage two");
  });

  it("drops the C0 controls XML 1.0 forbids", () => {
    expect(sanitizeXmlText(`a${NUL}b${SOH}c${BACKSPACE}d${SO}e${UNIT_SEP}f`)).toBe("abcdef");
  });

  it("drops U+FFFE and U+FFFF", () => {
    expect(sanitizeXmlText(`a${NOT_A_CHARACTER}b${ch(0xffff)}c`)).toBe("abc");
  });

  it("keeps DEL and the C1 range, which XML 1.0 permits", () => {
    const text = `a${DEL}b${ch(0x9f)}c`;
    expect(sanitizeXmlText(text)).toBe(text);
  });

  it("drops a lone surrogate but keeps a real pair", () => {
    expect(sanitizeXmlText(`a${HIGH_SURROGATE}b`)).toBe("ab");
    expect(sanitizeXmlText(`a${LOW_SURROGATE}b`)).toBe("ab");
    // U+1F600, a legitimate pair — must survive intact.
    expect(sanitizeXmlText("a😀b")).toBe("a😀b");
    expect(hasIllegalXmlScalars("a😀b")).toBe(false);
  });

  it("handles a run of several problems at once", () => {
    expect(sanitizeXmlText(`${NUL}start${VT}mid${SOH}${HIGH_SURROGATE}end${FF}`)).toBe(
      "start\nmidend\n",
    );
  });

  it("is stable when applied twice", () => {
    const once = sanitizeXmlText(`a${VT}b${NUL}c${HIGH_SURROGATE}`);
    expect(sanitizeXmlText(once)).toBe(once);
  });

  // A /g regex carries lastIndex between calls, which would make a shared
  // `test` regex alternate true/false on the same input.
  it("gives the same answer on repeated calls with the same input", () => {
    const dirty = `a${VT}b`;
    expect([1, 2, 3, 4].map(() => hasIllegalXmlScalars(dirty))).toEqual([true, true, true, true]);
    expect([1, 2, 3, 4].map(() => sanitizeXmlText(dirty))).toEqual(["a\nb", "a\nb", "a\nb", "a\nb"]);
  });

  it("handles the empty string", () => {
    expect(sanitizeXmlText("")).toBe("");
    expect(hasIllegalXmlScalars("")).toBe(false);
  });
});
