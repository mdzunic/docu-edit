// The model-shaped equality relations SPEC-008's round-trip guarantees are
// stated in (see specs/008-round-trip-preservation.md "Equality").
// `equalBytes` lives in core/equality.ts and the XML canonicalization in
// core/xml.ts — both product-neutral; this module holds the relations typed
// against the document model:
//
//  - `equalStructural` — deep equality of two document models after
//    canonicalization, reporting the path of the first difference so a
//    failing round-trip test says *where* it broke. Opaque XML is compared
//    with XML canonicalization (C14N-style), so attribute order, namespace
//    prefixes and quoting can never cause false failures.
//
//  - `equalSemantic`    — `equalStructural(resolve(a), resolve(b))`, the
//    relation for cross-format conversion, always paired with an exclusion
//    list (SPEC-008's `dropped`/`withoutDropped`, see lib/opaque.ts).
//
// Framework-free on purpose: this must run under Vitest (jsdom) and in the
// browser with nothing but DOMParser.
import { canonicalizeXml, XmlParseError } from "./core/xml";
import type { Block, DocumentModel, HeaderFooterContent, Paragraph } from "./documentModel";
import { resolveNamedStyle } from "./styles";
import { mergeParagraphFormatting, mergeRunFormatting } from "./documentModel";


function xmlEquivalent(a: string, b: string): boolean {
  try {
    return canonicalizeXml(a) === canonicalizeXml(b);
  } catch (error) {
    if (error instanceof XmlParseError) {
      // Unparseable opaque payloads fall back to string identity.
      return a === b;
    }
    throw error;
  }
}

// --- structural equality ---------------------------------------------------

/** Model fields holding XML strings, compared canonically. */
const XML_STRING_KEYS = new Set([
  "xml",
  "docxSectPr",
  "odtFontFaceDecls",
  // Verbatim table properties preserved from a .docx source (SPEC-022).
  "docxTblPr",
  "docxTrPr",
  "docxTcPr",
]);

/** Model fields holding arrays of XML strings. */
const XML_ARRAY_KEYS = new Set(["odtAutomaticStyles"]);

/** Fields excluded from structural equality: `id`s are runtime-only and
 * regenerated freely; the rest are UI state, not content — none of them is
 * written to a `.docx`/`.odt` file, so a same-format round trip never
 * carries them either way. */
const IGNORED_KEYS = new Set([
  "id",
  "suppressConversionReport",
  "measurementUnit",
  "showRuler",
]);

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/** Absent, undefined, empty-object and empty-array values are all "nothing
 * to say" — a reader that writes `formatting: {}` and one that omits the
 * field describe the same document. */
function isEmpty(value: unknown): boolean {
  if (value === undefined) {
    return true;
  }
  if (Array.isArray(value)) {
    return value.length === 0;
  }
  if (isPlainObject(value)) {
    return Object.keys(value).every((key) => IGNORED_KEYS.has(key) || isEmpty(value[key]));
  }
  return false;
}

function difference(
  a: unknown,
  b: unknown,
  path: string,
  xmlValued: boolean,
): string | null {
  if (xmlValued && typeof a === "string" && typeof b === "string") {
    return xmlEquivalent(a, b) ? null : path;
  }
  if (Array.isArray(a) && Array.isArray(b)) {
    if (a.length !== b.length) {
      return `${path}.length`;
    }
    for (let i = 0; i < a.length; i++) {
      const diff = difference(a[i], b[i], `${path}[${i}]`, xmlValued);
      if (diff !== null) {
        return diff;
      }
    }
    return null;
  }
  if (isPlainObject(a) && isPlainObject(b)) {
    const keys = [...new Set([...Object.keys(a), ...Object.keys(b)])].sort();
    for (const key of keys) {
      if (IGNORED_KEYS.has(key)) {
        continue;
      }
      const left = a[key];
      const right = b[key];
      if (isEmpty(left) && isEmpty(right)) {
        continue;
      }
      const childPath = path.length > 0 ? `${path}.${key}` : key;
      if (isEmpty(left) !== isEmpty(right)) {
        return childPath;
      }
      const diff = difference(
        left,
        right,
        childPath,
        XML_STRING_KEYS.has(key) || XML_ARRAY_KEYS.has(key),
      );
      if (diff !== null) {
        return diff;
      }
    }
    return null;
  }
  if (typeof a === "number" && typeof b === "number" && Number.isNaN(a) && Number.isNaN(b)) {
    return null;
  }
  return a === b ? null : path;
}

/** The path of the first structural difference between two models (e.g.
 * "sections[0].blocks[3].runs[1].text"), or null when they are equal. */
export function firstStructuralDifference(a: unknown, b: unknown): string | null {
  return difference(a, b, "", false);
}

/** Structural equality between two document models: deep equality after
 * dropping ids and empty/undefined fields, with opaque XML compared via
 * `canonicalizeXml`. */
export function equalStructural(a: unknown, b: unknown): boolean {
  return firstStructuralDifference(a, b) === null;
}

// --- semantic equality (specs/002, "Implementation status") ----------------

function resolveParagraph(model: DocumentModel, paragraph: Paragraph): void {
  const styleResolved = resolveNamedStyle(model, paragraph.styleId);
  // mergeParagraphFormatting/mergeRunFormatting are a plain `{...base,
  // ...override}` (documentModel.ts) — the style's own contribution
  // underneath whatever the paragraph/run already bakes, so a fully-baked
  // run (every reader's own convention) is unaffected (its keys already
  // cover the style's) while a run that merely references a style for a
  // property it never baked directly picks that property up here instead.
  paragraph.formatting = mergeParagraphFormatting(styleResolved.paragraphFormatting, paragraph.formatting ?? {});
  for (const run of paragraph.runs) {
    const runStyleResolved = run.styleId ? resolveNamedStyle(model, run.styleId).runFormatting : {};
    run.formatting = mergeRunFormatting(
      mergeRunFormatting(styleResolved.runFormatting, runStyleResolved),
      run.formatting,
    );
    delete run.styleId;
  }
  delete paragraph.styleId;
}

function resolveBlocks(model: DocumentModel, blocks: Block[]): void {
  for (const block of blocks) {
    if (block.type === "paragraph") {
      resolveParagraph(model, block);
    } else if (block.type === "table") {
      delete block.styleId;
      for (const row of block.rows) {
        for (const cell of row.cells) {
          for (const paragraph of cell.blocks) {
            resolveParagraph(model, paragraph);
          }
        }
      }
    }
  }
}

function resolveHeaderFooter(model: DocumentModel, content: HeaderFooterContent | undefined): void {
  if (content) {
    resolveBlocks(model, content.blocks);
  }
}

/** Flattens a model's paragraph/run/table style *references* into the
 * effective formatting they resolve to, merged underneath whatever
 * formatting is already baked (specs/002 "Equality", relation 3: "so a
 * directly-bolded run and a run bolded by its style compare equal"). Every
 * reader in this codebase bakes resolved formatting onto a paragraph/run
 * unconditionally, which is what makes the merge a no-op for them — the
 * style contributes nothing baked formatting doesn't already carry — but a
 * writer that instead emits a style reference *in place of* baked
 * formatting (specs/002-create-blank-document.md's "Implementation status")
 * is resolved correctly here too, since the merge fills in exactly what
 * baked formatting left out. Resolution reads `model.styles` (and its
 * built-in fallback, lib/styles.ts) as they stand on the model passed in —
 * call this on each side of a comparison separately, each against its own
 * style catalogue, never on a model whose styles have already been merged
 * with another's. */
export function resolve(model: DocumentModel): DocumentModel {
  const clone = structuredClone(model);
  for (const section of clone.sections) {
    resolveBlocks(clone, section.blocks);
    for (const header of Object.values(section.headers ?? {})) {
      resolveHeaderFooter(clone, header);
    }
    for (const footer of Object.values(section.footers ?? {})) {
      resolveHeaderFooter(clone, footer);
    }
  }
  for (const note of clone.notes ?? []) {
    resolveBlocks(clone, note.blocks);
  }
  return clone;
}

/** Semantic equality between two document models (specs/002 "Equality",
 * relation 3): structural equality once style *references* are flattened
 * away (see `resolve`). This is the relation for a cross-format conversion,
 * and — per specs/008-round-trip-preservation.md's "Writer API" — it must
 * always be paired with an explicit exclusion list naming what the
 * conversion is permitted to lose, applied to the source side as a mask
 * (`lib/opaque.ts`'s `withoutDropped`) rather than a hand-rolled subset
 * comparison. */
export function equalSemantic(a: DocumentModel, b: DocumentModel): boolean {
  return equalStructural(resolve(a), resolve(b));
}
