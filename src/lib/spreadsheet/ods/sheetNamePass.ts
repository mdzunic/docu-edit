// The retained-XML sheet-reference pass for .ods text (SPEC-121 "Sheet
// references in formulas", retained side): renamed and deleted sheet names
// rewritten in exactly the table's locations, every other byte verbatim.
//
// | content.xml | `table:formula`; `table:named-range@table:cell-range-address`,
// |             | `@table:base-cell-address`; `table:named-expression@table:expression`,
// |             | `@table:base-cell-address`; `table:database-range@table:target-range-address`;
// |             | `table:content-validation@table:condition`, `@table:base-cell-address`;
// |             | `calcext:condition@calcext:value`, `@calcext:base-cell-address`;
// |             | `calcext:conditional-format@calcext:target-range-address`;
// |             | `@table:end-cell-address`; `draw:object@draw:notify-on-update-of-ranges`;
// |             | `text:a@xlink:href` starting `#` |
// | Object N/content.xml | `@chart:values-cell-range-address`, `@table:cell-range-address`,
// |                      | `@chart:label-cell-address` |
//
// - Formulas (`table:formula`, a named expression's `table:expression`) go
//   through fromOdsFormula → remapSheetReferences → toOdsFormula, emitted in
//   `of:` form; one whose remapped canonical text equals its original, or
//   that the dialect cannot read, is left as it is.
// - Address lists go through remapOdsAddresses (a deleted sheet becomes
//   LibreOffice's `$#REF!.`). A condition (`of:cell-content-is-in-list(…)`,
//   `formula-is(…)`) is not a formula: each bracketed address in it is
//   remapped the same way.
// - Text that mentions no remapped name is returned unscanned.
// removeRangesOnDeletedSheets drops the named ranges and database ranges
// whose range lies on a deleted sheet (SPEC-121 "Deleting a sheet").
import { fromOdsFormula, toOdsFormula } from "../formula/dialect";
import { FormulaSyntaxError } from "../formula/parse";
import { isEmptyRemap, remapOdsAddresses, remapSheetReferences, type SheetRemap } from "../formula/sheetRefs";
import { decodeXmlText, escapeAttributeValue } from "../xmlSplice";
import { remappedNames, textMentionsAnyName } from "../xlsx/sheetNamePass";
import { parseOdsRangeAddress, splitOdsRangeAddressList } from "./rangeAddress";

type Transform = (value: string, remap: SheetRemap, namespaces: ReadonlyMap<string, string>) => string;

/** A formula attribute through the ODS pair. */
const formula: Transform = (value, remap, namespaces) => {
  let canonical: string;
  try {
    canonical = fromOdsFormula(value, namespaces);
  } catch (error) {
    if (error instanceof FormulaSyntaxError) {
      return value;
    }
    throw error;
  }
  const next = remapSheetReferences(canonical, remap);
  return next === canonical ? value : toOdsFormula(next);
};

const addresses: Transform = (value, remap) => remapOdsAddresses(value, remap);

/** Each bracketed address of a condition, remapped. */
const condition: Transform = (value, remap) =>
  value.replace(/\[([^[\]]*)\]/g, (whole, inner: string) => {
    const next = remapOdsAddresses(inner, remap);
    return next === inner ? whole : `[${next}]`;
  });

/** An internal link (`#Sheet.A1`); any other href is not a location. */
const href: Transform = (value, remap) => {
  if (!value.startsWith("#")) {
    return value;
  }
  const next = remapOdsAddresses(value.slice(1), remap);
  return next === value.slice(1) ? value : `#${next}`;
};

const LOCATIONS: Readonly<Record<string, Transform>> = {
  "table:formula": formula,
  "table:expression": formula,
  "table:cell-range-address": addresses,
  "table:base-cell-address": addresses,
  "table:target-range-address": addresses,
  "table:end-cell-address": addresses,
  "table:condition": condition,
  "calcext:value": condition,
  "calcext:base-cell-address": addresses,
  "calcext:target-range-address": addresses,
  "draw:notify-on-update-of-ranges": addresses,
  "xlink:href": href,
  "chart:values-cell-range-address": addresses,
  "chart:label-cell-address": addresses,
};

const LOCATION_PATTERN = new RegExp(`(\\s(${Object.keys(LOCATIONS).join("|")})\\s*=\\s*")([^"]*)"`, "g");

/** The prefix → namespace map an ODS part's root declares (the dialect
 * prefixes of `table:formula` resolve against it). */
export function rootNamespaces(xml: string): Map<string, string> {
  const root = /<[\w.-]+:document[\w-]*\b[^>]*>/.exec(xml)?.[0] ?? "";
  return new Map([...root.matchAll(/\sxmlns:([\w.-]+)\s*=\s*"([^"]*)"/g)].map((match) => [match[1]!, match[2]!] as const));
}

/** Retained ODS text with the pass applied (see the module header). */
export function passOdsText(xml: string, remap: SheetRemap, namespaces: ReadonlyMap<string, string>): string {
  if (isEmptyRemap(remap) || !textMentionsAnyName(xml, remappedNames(remap))) {
    return xml;
  }
  return xml.replace(LOCATION_PATTERN, (whole, head: string, name: string, raw: string) => {
    const value = decodeXmlText(raw);
    const next = LOCATIONS[name]!(value, remap, namespaces);
    return next === value ? whole : `${head}${escapeAttributeValue(next)}"`;
  });
}

/** Whether an address list names a deleted sheet. */
function onDeletedSheet(list: string, remap: SheetRemap): boolean {
  const deleted = new Set(remap.deleted.map((name) => name.toUpperCase()));
  return splitOdsRangeAddressList(list).some((entry, index) => {
    if (index % 2 === 1) {
      return false;
    }
    const ref = parseOdsRangeAddress(entry);
    return ref !== null && [ref.sheet, ref.sheetEnd].some((sheet) => sheet !== undefined && deleted.has(sheet.toUpperCase()));
  });
}

const RANGE_ELEMENTS: ReadonlyArray<[RegExp, string]> = [
  [/<table:named-range\b[^>]*?(?:\/>|>[\s\S]*?<\/table:named-range>)/g, "table:cell-range-address"],
  [/<table:database-range\b[^>]*?(?:\/>|>[\s\S]*?<\/table:database-range>)/g, "table:target-range-address"],
];

/** `xml` without the named ranges and database ranges whose range lies on
 * a deleted sheet. */
export function removeRangesOnDeletedSheets(xml: string, remap: SheetRemap): string {
  if (remap.deleted.length === 0) {
    return xml;
  }
  let out = xml;
  for (const [pattern, attribute] of RANGE_ELEMENTS) {
    out = out.replace(pattern, (element) => {
      const match = new RegExp(`\\s${attribute}\\s*=\\s*"([^"]*)"`).exec(element);
      return match && onDeletedSheet(decodeXmlText(match[1]!), remap) ? "" : element;
    });
  }
  return out;
}
