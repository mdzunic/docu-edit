// Sheet references in formulas (SPEC-121 "Sheet references in formulas") —
// THE rewrite sheet names move with. The model side: renameSheet and
// deleteSheet remap every Cell.formula through remapSheetReferences. The
// retained side: each writer remaps source-time retained text — formula text
// through remapSheetReferences, ODS address lists through remapOdsAddresses —
// with one remap derived from the source workbook; that side's helpers
// (sourceSheetRemap, resolveRetainedSheetName) read the source package, so
// they live in ../sourceNames.ts and are re-exported at the end of this
// module. Callers skip `formulaUnparsed` text (wire text, never transformed).
//
// - A SheetRemap names sheets by their OLD names, the names the text was
//   written against, matched case-insensitively (toUpperCase on both
//   sides): `renamed` maps one to its new name, `deleted` lists the gone,
//   `orderBefore` is the sheet order the text was written against (3-D
//   endpoints walk it). A name both renamed and deleted is deleted. An empty
//   remap (isEmptyRemap) returns the text at once.
// - remapSheetReferences is token-level over tokens.ts — the sibling of
//   rewriteReferences, not a caller, because that cannot express a
//   prefix-less `#REF!`: only the tokens it rewrites change, whitespace and
//   every other token survive byte-for-byte, and a formula nothing touched
//   comes back as the identical string. Every token is looked up once
//   against the old names, so a remap is one simultaneous pass — a swap of
//   two names round-trips — and for a single rename the result equals
//   renameSheetInFormula(text, from, to) (property-tested).
// - A `ref` naming a deleted sheet becomes the error literal `#REF!` as a
//   whole, no prefix (SPEC-119's out-of-grid convention; Excel would show
//   `#REF!A1`). A renamed one is formatReference'd with the new name,
//   re-quoted as SPEC-119 prints; one whose new name is spelled exactly as
//   written stays verbatim.
// - A 3-D `ref` (`Jan:Mar!B2`) moves a deleted endpoint INWARD along
//   `orderBefore` — the start toward the end, the end toward the start (a
//   reversed span walks the other way) — to the nearest sheet not deleted.
//   A span left with one sheet becomes a plain reference (`Feb!B2`); one
//   left with none, or with an endpoint `orderBefore` does not hold while
//   one of them is deleted, becomes `#REF!`. Renames then apply to the
//   surviving endpoints. Deleting a sheet strictly inside the span changes
//   nothing; a rename alone never collapses a span (renameSheetInFormula's
//   `Facts:Facts!A1`).
// - A sheet-prefixed `name` or `error` token (`Sheet1!Rate`, `Sheet2!#REF!`,
//   `Jan:Mar!#REF!`) becomes `#REF!` when any sheet it names is deleted, and
//   is renamed as renameSheetInFormula renames it (formatSheetPrefix + the
//   rest as written); a prefixed `function` token is renamed only, never
//   `#REF!`. External references (`[1]Sheet1!A1`) name another workbook's
//   sheets and are left alone.
// - remapOdsAddresses rewrites an ODS cell-range-address list
//   (ods/rangeAddress.ts) entry by entry, keeping the space runs between
//   entries byte-for-byte: an entry the grammar cannot read, or whose sheets
//   the remap neither renames nor deletes, keeps its exact bytes; a remapped
//   one goes through odsRangeAddressToCanonical's reference, this remap and
//   canonicalToOdsRangeAddress's form; one whose every sheet is deleted
//   becomes LibreOffice's `$#REF!.` form keeping its cell parts
//   (`$Lookup.$A$1:.$B$5` → `$#REF!.$A$1:.$B$5`).
import { canonicalToOdsRangeAddress, formatOdsRangeAddress, parseOdsRangeAddress, splitOdsRangeAddressList } from "../ods/rangeAddress";
import { formatReference, formatSheetPrefix } from "./print";
import type { ReferenceNode } from "./references";
import { joinTokens, parseSheetPrefix, sheetPrefixOf, tokenize, type Token } from "./tokens";

export interface SheetRemap {
  /** OLD name (matched case-insensitively) → new name. */
  renamed: Record<string, string>;
  /** Names now gone (matched case-insensitively). */
  deleted: string[];
  /** The sheet order the text was written against (3-D endpoints). */
  orderBefore: string[];
}

/** What a reference into a deleted sheet becomes. */
const REF_ERROR = "#REF!";

/** No renames and no deletions: remapping changes nothing. */
export function isEmptyRemap(remap: SheetRemap): boolean {
  return remap.deleted.length === 0 && Object.keys(remap.renamed).length === 0;
}

/** A remap's case-insensitive lookups, keyed by upper-cased old name. */
interface RemapIndex {
  renamed: ReadonlyMap<string, string>;
  deleted: ReadonlySet<string>;
  orderBefore: readonly string[];
}

function indexRemap(remap: SheetRemap): RemapIndex {
  const deleted = new Set(remap.deleted.map((name) => name.toUpperCase()));
  const renamed = new Map<string, string>();
  for (const [from, to] of Object.entries(remap.renamed)) {
    const key = from.toUpperCase();
    if (!deleted.has(key)) {
      renamed.set(key, to);
    }
  }
  return { renamed, deleted, orderBefore: remap.orderBefore };
}

function isDeleted(index: RemapIndex, name: string): boolean {
  return index.deleted.has(name.toUpperCase());
}

/** A surviving sheet's name after the remap: its new name, or as written. */
function renamedName(index: RemapIndex, name: string): string {
  return index.renamed.get(name.toUpperCase()) ?? name;
}

/** The surviving endpoints of a 3-D span with a deleted endpoint (see the
 * module header) — as written, or as `orderBefore` names a sheet an endpoint
 * moved to; `sheetEnd` absent when one sheet survives; null for `#REF!`. */
function survivingSpan(index: RemapIndex, sheet: string, sheetEnd: string): { sheet: string; sheetEnd?: string } | null {
  const order = index.orderBefore.map((name) => name.toUpperCase());
  const first = order.indexOf(sheet.toUpperCase());
  const last = order.indexOf(sheetEnd.toUpperCase());
  if (first < 0 || last < 0) {
    return null;
  }
  /** The first sheet not deleted walking from `from` to `to`, inclusive. */
  const inward = (from: number, to: number): number => {
    const step = from <= to ? 1 : -1;
    for (let k = from; k !== to + step; k += step) {
      if (!index.deleted.has(order[k]!)) {
        return k;
      }
    }
    return -1;
  };
  const start = index.deleted.has(order[first]!) ? inward(first, last) : first;
  if (start < 0) {
    return null;
  }
  // Something survives, so the end's walk finds a sheet (at worst `start`).
  const end = index.deleted.has(order[last]!) ? inward(last, first) : last;
  const startName = start === first ? sheet : index.orderBefore[start]!;
  return start === end ? { sheet: startName } : { sheet: startName, sheetEnd: end === last ? sheetEnd : index.orderBefore[end]! };
}

/** A sheet-qualified reference's new text, or undefined to keep it. */
function remappedReference(ref: ReferenceNode & { sheet: string }, index: RemapIndex): string | undefined {
  let span: { sheet: string; sheetEnd?: string } = { sheet: ref.sheet, sheetEnd: ref.sheetEnd };
  if (ref.sheetEnd === undefined) {
    if (isDeleted(index, ref.sheet)) {
      return REF_ERROR;
    }
  } else if (isDeleted(index, ref.sheet) || isDeleted(index, ref.sheetEnd)) {
    const surviving = survivingSpan(index, ref.sheet, ref.sheetEnd);
    if (!surviving) {
      return REF_ERROR;
    }
    span = surviving;
  }
  const sheet = renamedName(index, span.sheet);
  const sheetEnd = span.sheetEnd === undefined ? undefined : renamedName(index, span.sheetEnd);
  if (sheet === ref.sheet && sheetEnd === ref.sheetEnd) {
    return undefined;
  }
  return formatReference({ ...ref, sheet, sheetEnd });
}

/** A token's remapped text, or undefined to keep it. */
function remappedToken(token: Token, index: RemapIndex): string | undefined {
  if (token.kind === "ref") {
    const sheet = token.ref?.sheet;
    return token.ref && sheet !== undefined ? remappedReference({ ...token.ref, sheet }, index) : undefined;
  }
  if (token.kind !== "name" && token.kind !== "error" && token.kind !== "function") {
    return undefined;
  }
  const prefix = sheetPrefixOf(token.text);
  if (prefix === "") {
    return undefined;
  }
  const { sheet, sheetEnd } = parseSheetPrefix(prefix);
  if (token.kind !== "function" && (isDeleted(index, sheet) || (sheetEnd !== undefined && isDeleted(index, sheetEnd)))) {
    return REF_ERROR;
  }
  const renamedSheet = renamedName(index, sheet);
  const renamedEnd = sheetEnd === undefined ? undefined : renamedName(index, sheetEnd);
  if (renamedSheet === sheet && renamedEnd === sheetEnd) {
    return undefined;
  }
  return formatSheetPrefix(renamedSheet, renamedEnd) + token.text.slice(prefix.length);
}

/** Canonical formula text with every sheet the remap renames or deletes
 * remapped (see the module header); the identical string when nothing is. */
export function remapSheetReferences(canonical: string, remap: SheetRemap): string {
  // Every sheet-qualified token holds a `!`.
  if (isEmptyRemap(remap) || !canonical.includes("!")) {
    return canonical;
  }
  const index = indexRemap(remap);
  const tokens = tokenize(canonical);
  let changed = false;
  for (const token of tokens) {
    const next = remappedToken(token, index);
    if (next !== undefined && next !== token.text) {
      token.text = next;
      changed = true;
    }
  }
  return changed ? joinTokens(tokens) : canonical;
}

/** One cell-range address (a single entry) remapped (see the module
 * header); the identical string when the remap does not touch it. */
export function remapOdsAddress(entry: string, remap: SheetRemap): string {
  const ref = parseOdsRangeAddress(entry);
  if (!ref || ref.sheet === undefined) {
    return entry;
  }
  const canonical = formatReference(ref);
  const remapped = remapSheetReferences(canonical, remap);
  if (remapped === canonical) {
    return entry;
  }
  if (remapped === REF_ERROR) {
    return `$${REF_ERROR}${formatOdsRangeAddress({ start: ref.start, end: ref.end })}`;
  }
  return canonicalToOdsRangeAddress(remapped) ?? entry;
}

/** An ODS cell-range-address list ("$Sheet1.A1:.B2 $'My S'.C3") remapped
 * entry by entry, separators kept (see the module header); the identical
 * string when no entry changes. */
export function remapOdsAddresses(list: string, remap: SheetRemap): string {
  if (isEmptyRemap(remap)) {
    return list;
  }
  const pieces = splitOdsRangeAddressList(list);
  let changed = false;
  for (let i = 0; i < pieces.length; i += 2) {
    const entry = pieces[i]!;
    const next = remapOdsAddress(entry, remap);
    if (next !== entry) {
      pieces[i] = next;
      changed = true;
    }
  }
  return changed ? pieces.join("") : list;
}

// The retained side's helpers (they read the source package's sheet list).
export { resolveRetainedSheetName, sourceSheetRemap } from "../sourceNames";
