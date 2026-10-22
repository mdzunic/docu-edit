// Formatted-run construction shared by every block builder (heading,
// paragraph, list item, table cell — SPEC-045 AC-1): a run is either a
// plain string (unformatted text) or an object naming which of the
// canonical model's RunFormatting fields to set. This is the API's only
// path to a `Run`, so paragraph/heading/list/table content is guaranteed
// the same shape and validation regardless of which block type it lands in.
import type { Run, RunFormatting } from "../documentModel";
import { DocumentGenerationError } from "./errors";
import {
  expectBoolean,
  expectHexColor,
  expectPlainObject,
  expectPositiveNumber,
  expectString,
  rejectUnknownKeys,
} from "./validate";

const RUN_KEYS = [
  "text",
  "bold",
  "italic",
  "underline",
  "strikethrough",
  "fontFamily",
  "fontSizePt",
  "color",
  "highlight",
] as const;

export interface RunSpecObject {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  fontFamily?: string;
  fontSizePt?: number;
  color?: string;
  highlight?: string;
}

/** A run of text: a plain string, or an object naming its text and
 * formatting. */
export type RunSpec = string | RunSpecObject;

export function buildRun(spec: unknown, path: string): Run {
  if (typeof spec === "string") {
    return { text: spec, formatting: {} };
  }
  const obj = expectPlainObject(spec, path);
  rejectUnknownKeys(obj, RUN_KEYS, path);
  const text = expectString(obj.text, `${path}.text`);
  const formatting: RunFormatting = {};
  if (obj.bold !== undefined) {
    formatting.bold = expectBoolean(obj.bold, `${path}.bold`);
  }
  if (obj.italic !== undefined) {
    formatting.italic = expectBoolean(obj.italic, `${path}.italic`);
  }
  if (obj.underline !== undefined) {
    formatting.underline = expectBoolean(obj.underline, `${path}.underline`);
  }
  if (obj.strikethrough !== undefined) {
    formatting.strikethrough = expectBoolean(obj.strikethrough, `${path}.strikethrough`);
  }
  if (obj.fontFamily !== undefined) {
    formatting.fontFamily = expectString(obj.fontFamily, `${path}.fontFamily`);
  }
  if (obj.fontSizePt !== undefined) {
    formatting.fontSizePt = expectPositiveNumber(obj.fontSizePt, `${path}.fontSizePt`);
  }
  if (obj.color !== undefined) {
    formatting.color = expectHexColor(obj.color, `${path}.color`);
  }
  if (obj.highlight !== undefined) {
    formatting.highlight = expectHexColor(obj.highlight, `${path}.highlight`);
  }
  return { text, formatting };
}

export function buildRuns(spec: unknown, path: string): Run[] {
  const arr = spec;
  if (!Array.isArray(arr)) {
    throw new DocumentGenerationError(`${path}: expected an array of runs`);
  }
  return arr.map((item, i) => buildRun(item, `${path}[${i}]`));
}

/** Paragraph-shaped content: a plain string (one unformatted run), an array
 * of runs, or an object naming exactly one of `text`/`runs` — the shorthand
 * every block type that carries text (heading, paragraph, list item, table
 * cell) accepts identically (SPEC-045 AC-4's 1:1 JSON/builder mapping). */
export function buildRunsFromContent(spec: unknown, path: string): Run[] {
  if (typeof spec === "string") {
    return [{ text: spec, formatting: {} }];
  }
  if (Array.isArray(spec)) {
    return buildRuns(spec, path);
  }
  const obj = expectPlainObject(spec, path);
  rejectUnknownKeys(obj, ["text", "runs"], path);
  return buildRunsFromContent(contentFromTextOrRuns(obj, path), path);
}

/** The `text`/`runs` pair every block-spec object carries alongside its own
 * other fields, resolved to the SAME raw (not-yet-built) content shape
 * `buildRunsFromContent` accepts — a string, an array of run specs, or `[]`
 * for neither given (an empty paragraph). Returning raw content rather than
 * built `Run[]` matters: the caller (blocks.ts's dispatcher) still has to
 * feed this into `buildHeadingParagraph`/`buildParagraph`, which build runs
 * themselves — passing already-built `Run[]` back through that would parse
 * a `Run`'s own `{ text, formatting }` shape as a `RunSpec` a second time
 * and reject `formatting` as an unknown property. The whole object's other
 * keys are validated by the caller, not here. */
export function contentFromTextOrRuns(obj: Record<string, unknown>, path: string): unknown {
  if (obj.text !== undefined && obj.runs !== undefined) {
    throw new DocumentGenerationError(`${path}: specify either "text" or "runs", not both`);
  }
  if (obj.runs !== undefined) {
    return obj.runs;
  }
  if (obj.text !== undefined) {
    return expectString(obj.text, `${path}.text`);
  }
  return [];
}
