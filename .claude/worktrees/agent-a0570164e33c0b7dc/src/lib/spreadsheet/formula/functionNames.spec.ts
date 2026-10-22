// SPEC-119 "Function names": the transcribed table reproduces every test
// vector the spec pins, holds no duplicate canonical or wire name in either
// format, round-trips every entry both ways, and leaves a name absent from it
// identical in all three spellings.
import { describe, expect, it } from "vitest";
import { fromOdsFormula, fromXlsxFormula, toOdsFormula, toXlsxFormula } from "./dialect";
import { canonicalFromOds, canonicalFromXlsx, FUNCTION_NAMES, odsFromCanonical, xlsxFromCanonical } from "./functionNames";
import { getFunction, registeredFunctionNames } from "./registry";

describe("the pinned test vectors", () => {
  it.each([
    ["CONCAT", "_xlfn.CONCAT", "COM.MICROSOFT.CONCAT"],
    ["TEXTJOIN", "_xlfn.TEXTJOIN", "COM.MICROSOFT.TEXTJOIN"],
    ["IFS", "_xlfn.IFS", "COM.MICROSOFT.IFS"],
    ["SWITCH", "_xlfn.SWITCH", "COM.MICROSOFT.SWITCH"],
    ["MINIFS", "_xlfn.MINIFS", "COM.MICROSOFT.MINIFS"],
    ["MAXIFS", "_xlfn.MAXIFS", "COM.MICROSOFT.MAXIFS"],
    ["XOR", "_xlfn.XOR", "XOR"],
    ["IFNA", "_xlfn.IFNA", "IFNA"],
    ["DAYS", "_xlfn.DAYS", "DAYS"],
    // "per source": formulabase.cxx saFuncTable2021.
    ["SORT", "_xlfn._xlws.SORT", "COM.MICROSOFT.SORT"],
    ["FILTER", "_xlfn._xlws.FILTER", "COM.MICROSOFT.FILTER"],
    // "per source": LibreOffice has no name for Excel's @.
    ["SINGLE", "_xlfn.SINGLE", "SINGLE"],
    ["SUM", "SUM", "SUM"],
    ["VLOOKUP", "VLOOKUP", "VLOOKUP"],
  ])("%s ↔ xlsx %s ↔ ods %s", (canonical, xlsx, ods) => {
    expect(xlsxFromCanonical(canonical)).toBe(xlsx);
    expect(canonicalFromXlsx(xlsx)).toBe(canonical);
    expect(odsFromCanonical(canonical)).toBe(ods);
    expect(canonicalFromOds(ods)).toBe(canonical);
  });

  it("transcribes LibreOffice's renamed and LEGACY names, which never collide", () => {
    expect(odsFromCanonical("F.DIST")).toBe("FDIST");
    expect(odsFromCanonical("FDIST")).toBe("LEGACY.FDIST");
    expect(canonicalFromOds("FDIST")).toBe("F.DIST");
    expect(canonicalFromOds("LEGACY.FDIST")).toBe("FDIST");
    expect(odsFromCanonical("FORMULATEXT")).toBe("FORMULA");
    expect(odsFromCanonical("SKEW.P")).toBe("SKEWP");
    expect(odsFromCanonical("CEILING")).toBe("COM.MICROSOFT.CEILING");
    expect(odsFromCanonical("CEILING.MATH")).toBe("COM.MICROSOFT.CEILING.MATH");
    expect(odsFromCanonical("ISO.CEILING")).toBe("COM.MICROSOFT.ISO.CEILING");
    expect(xlsxFromCanonical("ISO.CEILING")).toBe("ISO.CEILING");
    expect(odsFromCanonical("DBCS")).toBe("JIS");
  });

  it("keeps LibreOffice-only names as written", () => {
    expect(canonicalFromOds("ORG.OPENOFFICE.ERRORTYPE")).toBe("ORG.OPENOFFICE.ERRORTYPE");
    expect(canonicalFromOds("ORG.LIBREOFFICE.RAWSUBTRACT")).toBe("ORG.LIBREOFFICE.RAWSUBTRACT");
  });
});

describe("the table", () => {
  it("holds no duplicate canonical name, xlsx wire name or ODS wire name", () => {
    const canonical = FUNCTION_NAMES.map((entry) => entry.name);
    const xlsx = FUNCTION_NAMES.map((entry) => (entry.xlsxPrefix ?? "") + entry.name);
    const ods = FUNCTION_NAMES.map((entry) => entry.odsName ?? entry.name);
    expect(new Set(canonical).size).toBe(canonical.length);
    expect(new Set(xlsx).size).toBe(xlsx.length);
    expect(new Set(ods).size).toBe(ods.length);
  });

  it("lists only names that differ somewhere, canonical upper-case and unprefixed", () => {
    for (const entry of FUNCTION_NAMES) {
      expect(entry.xlsxPrefix !== undefined || (entry.odsName !== undefined && entry.odsName !== entry.name), entry.name).toBe(true);
      expect(entry.name, entry.name).toBe(entry.name.toUpperCase());
      expect(entry.name.startsWith("_")).toBe(false);
    }
  });

  it("round-trips every entry both ways in both formats", () => {
    for (const entry of FUNCTION_NAMES) {
      const xlsx = xlsxFromCanonical(entry.name);
      const ods = odsFromCanonical(entry.name);
      expect(canonicalFromXlsx(xlsx), entry.name).toBe(entry.name);
      expect(canonicalFromOds(ods), entry.name).toBe(entry.name);
      expect(xlsxFromCanonical(canonicalFromXlsx(xlsx)), entry.name).toBe(xlsx);
      expect(odsFromCanonical(canonicalFromOds(ods)), entry.name).toBe(ods);
    }
  });

  it("leaves a name absent from the table identical in all three spellings", () => {
    for (const name of ["SUM", "VLOOKUP", "IF", "INDEX", "DAYS360", "MYUDF"]) {
      expect(xlsxFromCanonical(name)).toBe(name);
      expect(odsFromCanonical(name)).toBe(name);
      expect(canonicalFromOds(name)).toBe(name);
      expect(canonicalFromXlsx(name)).toBe(name);
    }
  });

  it("strips `_xlfn._xlws.` and `_xlfn.` in any case and upper-cases", () => {
    expect(canonicalFromXlsx("_xlfn._xlws.sort")).toBe("SORT");
    expect(canonicalFromXlsx("_XLFN.Concat")).toBe("CONCAT");
    expect(canonicalFromXlsx("_xlfn.VALUETOTEXT")).toBe("VALUETOTEXT");
  });
});

// SPEC-127: every registered name round-trips through the table on both
// wires — one call of each function, its required arguments as references.
describe("every registered function on the wire (SPEC-127)", () => {
  const calls = registeredFunctionNames().map((name) => `${name}(${Array.from({ length: getFunction(name)!.minArgs }, (_, i) => `A${i + 1}`).join(",")})`);

  it.each(calls)("%s round-trips through xlsx and ods", (text) => {
    expect(fromXlsxFormula(toXlsxFormula(text))).toBe(text);
    expect(fromOdsFormula(toOdsFormula(text))).toBe(text);
  });

  it("writes the future functions SPEC-127 registers with their prefixes", () => {
    expect(toXlsxFormula("XLOOKUP(A1,A2,A3)")).toBe("_xlfn.XLOOKUP(A1,A2,A3)");
    expect(toXlsxFormula("STDEV.S(A1)")).toBe("_xlfn.STDEV.S(A1)");
    expect(toXlsxFormula("AGGREGATE(A1,A2,A3)")).toBe("_xlfn.AGGREGATE(A1,A2,A3)");
    expect(toOdsFormula("XLOOKUP(A1,A2,A3)")).toContain("COM.MICROSOFT.XLOOKUP");
  });
});
