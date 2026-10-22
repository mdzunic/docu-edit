// Function names across the three spellings (SPEC-119 "Function names"):
// canonical (Excel's en-US UI name, upper-case, unprefixed — "CONCAT"), the
// xlsx wire (future functions prefixed `_xlfn.` — or `_xlfn._xlws.` for the
// worksheet-only ones — "_xlfn.CONCAT") and the ODS wire (OpenFormula's name
// — "COM.MICROSOFT.CONCAT"). One entry per function whose wire name differs
// from its canonical name in either format; a name absent from the table is
// identical in all three spellings (SUM, VLOOKUP, …).
//
// TRANSCRIBED, not recalled (fetched 2026-10-09):
// - the `_xlfn.` set — [MS-XLSX] "Functions" (the future-function table,
//   learn.microsoft.com/…/ms-xlsx/5d1b6d44-6fc1-4ecd-8fef-0b27406cc2bf, as
//   updated 2024-04-16) and the "Formulas" grammar's future-function-list /
//   worksheet-only-function-list (…/3d025add-118d-4413-9856-ab65712ec1b0,
//   which adds COPILOT and LONGTEXT). The "_xlfn. prefix" support article
//   (882f1ef7-68fb-4fcd-8d54-9fbb77fd5025) now redirects to "Detect formula
//   errors in Excel" and lists nothing, so where Microsoft's two lists are
//   silent the prefix comes from LibreOffice's FuncFlags::MACROCALL_NEW rows
//   (it writes `_xlfn.` for exactly those): ENCODEURL, XMATCH, ARRAYTOTEXT.
//   SINGLE and ANCHORARRAY are SPEC-119's own pins (Excel's wire forms of
//   `@` and `A1#`).
// - the ODS names — LibreOffice's sc/source/filter/oox/formulabase.cxx
//   (master), its Excel ↔ ODFF tables saFuncTableBiff3/Biff4/Biff8/Oox/
//   2010/2013/2016/2021/2024: { "<ODFF name>", "<OOXML name>", … }. Rows
//   flagged IMPORTONLY, EXPORTONLY or BIFFEXPORTONLY are one-way rewrites of
//   LibreOffice's own functions (DOLLAR ← USDOLLAR, CEILING → CEILING.MATH,
//   CSC → SIN, …), not name translations, and rows with a nullptr name
//   have no counterpart in the other format; neither is in this table.
//   ODFF's native CEILING/FLOOR (LibreOffice's own three-argument functions)
//   therefore read as Excel's CEILING/FLOOR, whose ODS spelling is
//   COM.MICROSOFT.CEILING/FLOOR.
// LibreOffice-only names (ORG.OPENOFFICE.*, ORG.LIBREOFFICE.*) have no
// canonical counterpart and stay as written; SPEC-120 reports them as
// unsupported.

export interface FunctionNameEntry {
  /** Canonical: "CONCAT". */
  name: string;
  /** How the xlsx wire spells it. */
  xlsxPrefix?: "_xlfn." | "_xlfn._xlws.";
  /** The OpenFormula name, when it differs. */
  odsName?: string;
}

const XLFN = "_xlfn.";
const XLWS = "_xlfn._xlws.";

/** A future function LibreOffice writes as `COM.MICROSOFT.<name>`. */
const microsoft = (name: string, xlsxPrefix: FunctionNameEntry["xlsxPrefix"] = XLFN): FunctionNameEntry => ({
  name,
  xlsxPrefix,
  odsName: `COM.MICROSOFT.${name}`,
});

/** A future function whose ODS name is its canonical one. */
const future = (name: string, xlsxPrefix: FunctionNameEntry["xlsxPrefix"] = XLFN): FunctionNameEntry => ({ name, xlsxPrefix });

export const FUNCTION_NAMES: readonly FunctionNameEntry[] = [
  // formulabase.cxx saFuncTable2010 (Excel 2010), all MACROCALL_NEW.
  microsoft("COVARIANCE.P"),
  microsoft("COVARIANCE.S"),
  microsoft("STDEV.P"),
  microsoft("STDEV.S"),
  microsoft("VAR.P"),
  microsoft("VAR.S"),
  microsoft("BETA.DIST"),
  microsoft("BETA.INV"),
  microsoft("BINOM.DIST"),
  microsoft("BINOM.INV"),
  microsoft("CHISQ.DIST"),
  microsoft("CHISQ.INV"),
  microsoft("CHISQ.DIST.RT"),
  microsoft("CHISQ.INV.RT"),
  microsoft("CHISQ.TEST"),
  microsoft("CONFIDENCE.NORM"),
  microsoft("CONFIDENCE.T"),
  { name: "F.DIST", xlsxPrefix: XLFN, odsName: "FDIST" },
  microsoft("F.DIST.RT"),
  { name: "F.INV", xlsxPrefix: XLFN, odsName: "FINV" },
  microsoft("F.INV.RT"),
  microsoft("F.TEST"),
  microsoft("EXPON.DIST"),
  microsoft("HYPGEOM.DIST"),
  microsoft("POISSON.DIST"),
  microsoft("WEIBULL.DIST"),
  microsoft("GAMMA.DIST"),
  microsoft("GAMMA.INV"),
  microsoft("GAMMALN.PRECISE"),
  microsoft("LOGNORM.DIST"),
  microsoft("LOGNORM.INV"),
  microsoft("NORM.DIST"),
  microsoft("NORM.INV"),
  microsoft("NORM.S.DIST"),
  microsoft("NORM.S.INV"),
  microsoft("T.DIST"),
  microsoft("T.DIST.2T"),
  microsoft("T.DIST.RT"),
  microsoft("T.INV"),
  microsoft("T.INV.2T"),
  microsoft("T.TEST"),
  microsoft("PERCENTILE.INC"),
  microsoft("PERCENTRANK.INC"),
  microsoft("QUARTILE.INC"),
  microsoft("RANK.EQ"),
  microsoft("PERCENTILE.EXC"),
  microsoft("PERCENTRANK.EXC"),
  microsoft("QUARTILE.EXC"),
  microsoft("RANK.AVG"),
  microsoft("MODE.SNGL"),
  microsoft("MODE.MULT"),
  microsoft("NEGBINOM.DIST"),
  microsoft("Z.TEST"),
  microsoft("CEILING.PRECISE"),
  microsoft("FLOOR.PRECISE"),
  microsoft("ERF.PRECISE"),
  microsoft("ERFC.PRECISE"),
  microsoft("AGGREGATE"),
  // formulabase.cxx saFuncTable2013 (Excel 2013), all MACROCALL_NEW.
  future("ACOT"),
  future("ACOTH"),
  future("ARABIC"),
  future("BASE"),
  future("BINOM.DIST.RANGE"),
  future("BITAND"),
  future("BITLSHIFT"),
  future("BITOR"),
  future("BITRSHIFT"),
  future("BITXOR"),
  microsoft("CEILING.MATH"),
  future("COMBINA"),
  future("COT"),
  future("COTH"),
  future("CSC"),
  future("CSCH"),
  future("DAYS"),
  future("DECIMAL"),
  microsoft("ENCODEURL"),
  microsoft("FILTERXML"),
  microsoft("FLOOR.MATH"),
  { name: "FORMULATEXT", xlsxPrefix: XLFN, odsName: "FORMULA" },
  future("GAMMA"),
  future("GAUSS"),
  future("IFNA"),
  future("IMCOSH"),
  future("IMCOT"),
  future("IMCSC"),
  future("IMCSCH"),
  future("IMSEC"),
  future("IMSECH"),
  future("IMSINH"),
  future("IMTAN"),
  future("ISFORMULA"),
  future("ISOWEEKNUM"),
  future("MUNIT"),
  future("NUMBERVALUE"),
  future("PDURATION"),
  future("PERMUTATIONA"),
  future("PHI"),
  future("RRI"),
  future("SEC"),
  future("SECH"),
  future("SHEET"),
  future("SHEETS"),
  { name: "SKEW.P", xlsxPrefix: XLFN, odsName: "SKEWP" },
  future("UNICHAR"),
  future("UNICODE"),
  microsoft("WEBSERVICE"),
  future("XOR"),
  // formulabase.cxx saFuncTable2016 (Excel 2016), all MACROCALL_NEW.
  microsoft("FORECAST.ETS"),
  microsoft("FORECAST.ETS.CONFINT"),
  microsoft("FORECAST.ETS.SEASONALITY"),
  microsoft("FORECAST.ETS.STAT"),
  microsoft("FORECAST.LINEAR"),
  microsoft("CONCAT"),
  microsoft("TEXTJOIN"),
  microsoft("IFS"),
  microsoft("SWITCH"),
  microsoft("MINIFS"),
  microsoft("MAXIFS"),
  // formulabase.cxx saFuncTable2021 (Excel 2021), all MACROCALL_NEW; FILTER
  // and SORT are [MS-XLSX]'s worksheet-only functions.
  microsoft("XLOOKUP"),
  microsoft("XMATCH"),
  microsoft("FILTER", XLWS),
  microsoft("SORT", XLWS),
  microsoft("SORTBY"),
  microsoft("SEQUENCE"),
  microsoft("RANDARRAY"),
  microsoft("UNIQUE"),
  microsoft("ARRAYTOTEXT"),
  microsoft("LET"),
  microsoft("LAMBDA"),
  microsoft("ISOMITTED"),
  microsoft("BYCOL"),
  microsoft("BYROW"),
  microsoft("MAKEARRAY"),
  microsoft("MAP"),
  microsoft("REDUCE"),
  microsoft("SCAN"),
  // formulabase.cxx saFuncTable2024, all MACROCALL_NEW.
  microsoft("CHOOSECOLS"),
  microsoft("CHOOSEROWS"),
  microsoft("DROP"),
  microsoft("EXPAND"),
  microsoft("HSTACK"),
  microsoft("VSTACK"),
  microsoft("TAKE"),
  microsoft("TEXTAFTER"),
  microsoft("TEXTBEFORE"),
  microsoft("TEXTSPLIT"),
  microsoft("TOCOL"),
  microsoft("TOROW"),
  microsoft("WRAPCOLS"),
  microsoft("WRAPROWS"),
  // [MS-XLSX] future functions LibreOffice has no ODFF name for.
  future("FIELDVALUE"),
  future("PQSOURCE"),
  future("PYTHON_STR"),
  future("PYTHON_TYPE"),
  future("PYTHON_TYPENAME"),
  future("QUERYSTRING"),
  future("COPILOT"),
  future("LONGTEXT"),
  future("PY", XLWS),
  // SPEC-119's pins: Excel's wire forms of `@x` and `A1#`.
  future("SINGLE"),
  future("ANCHORARRAY"),
  // Excel 2010 functions [MS-XLSX] lists without a prefix; formulabase.cxx
  // saFuncTableOox gives their ODFF names.
  { name: "ISO.CEILING", odsName: "COM.MICROSOFT.ISO.CEILING" },
  { name: "NETWORKDAYS.INTL", odsName: "COM.MICROSOFT.NETWORKDAYS.INTL" },
  { name: "WORKDAY.INTL", odsName: "COM.MICROSOFT.WORKDAY.INTL" },
  // Pre-2007 functions with an ODFF name of their own (saFuncTableBiff3,
  // Biff4, Biff8).
  { name: "DBCS", odsName: "JIS" },
  { name: "CHIDIST", odsName: "LEGACY.CHIDIST" },
  { name: "CHIINV", odsName: "LEGACY.CHIINV" },
  { name: "CHITEST", odsName: "LEGACY.CHITEST" },
  { name: "FDIST", odsName: "LEGACY.FDIST" },
  { name: "FINV", odsName: "LEGACY.FINV" },
  { name: "NORMSDIST", odsName: "LEGACY.NORMSDIST" },
  { name: "NORMSINV", odsName: "LEGACY.NORMSINV" },
  { name: "TDIST", odsName: "LEGACY.TDIST" },
  { name: "FLOOR", odsName: "COM.MICROSOFT.FLOOR" },
  { name: "CEILING", odsName: "COM.MICROSOFT.CEILING" },
  { name: "BAHTTEXT", odsName: "COM.MICROSOFT.BAHTTEXT" },
];

const BY_CANONICAL: ReadonlyMap<string, FunctionNameEntry> = new Map(FUNCTION_NAMES.map((entry) => [entry.name, entry]));
const BY_ODS_NAME: ReadonlyMap<string, FunctionNameEntry> = new Map(
  FUNCTION_NAMES.flatMap((entry) => (entry.odsName === undefined ? [] : [[entry.odsName, entry] as const])),
);

/** An xlsx wire name → canonical: `_xlfn._xlws.` or `_xlfn.` stripped,
 * upper-cased ("_xlfn.concat" → "CONCAT"). */
export function canonicalFromXlsx(wire: string): string {
  const upper = wire.toUpperCase();
  if (upper.startsWith(XLWS.toUpperCase())) {
    return upper.slice(XLWS.length);
  }
  return upper.startsWith(XLFN.toUpperCase()) ? upper.slice(XLFN.length) : upper;
}

/** A canonical name → the xlsx wire: its prefix from the table restored;
 * a name absent from the table unchanged. */
export function xlsxFromCanonical(name: string): string {
  const entry = BY_CANONICAL.get(name.toUpperCase());
  return entry?.xlsxPrefix === undefined ? name : entry.xlsxPrefix + entry.name;
}

/** An OpenFormula name → canonical ("COM.MICROSOFT.CONCAT" → "CONCAT"); a
 * name absent from the table unchanged. */
export function canonicalFromOds(wire: string): string {
  return BY_ODS_NAME.get(wire.toUpperCase())?.name ?? wire;
}

/** A canonical name → its OpenFormula name; a name absent from the table
 * unchanged. */
export function odsFromCanonical(name: string): string {
  return BY_CANONICAL.get(name.toUpperCase())?.odsName ?? name;
}
