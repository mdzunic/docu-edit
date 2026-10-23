// RunSpec → SlideRun (SPEC-106): every text field the deck builder accepts
// takes SPEC-045's `RunSpec` exactly as runs.ts declares it, validated by
// runs.ts's own `buildRun` (its closed `RUN_KEYS` guard is what makes the
// model-side name `sizePt` a `DocumentGenerationError` rather than a
// synonym). This module only maps the validated run onto SPEC-081's
// `SlideRun`: `fontSizePt` is renamed `sizePt`, the boolean marks keep only
// `true` (the slide model's convention), and `strikethrough`/`highlight` —
// which `SlideRun` has no field for — are dropped and counted, the same
// report slide Markdown gives for what it cannot carry.
import type { SlideParagraph, SlideRun } from "../presentation/model";
import { DocumentGenerationError } from "./errors";
import { buildRun, type RunSpec } from "./runs";

/** Run fields the slide model cannot carry, counted per field. */
export type SlideRunDrops = Map<string, number>;

function countDrop(drops: SlideRunDrops, field: string): void {
  drops.set(field, (drops.get(field) ?? 0) + 1);
}

/** One validated `RunSpec` as a `SlideRun`, counting what it drops. */
export function buildSlideRun(spec: unknown, path: string, drops: SlideRunDrops): SlideRun {
  const { text, formatting } = buildRun(spec, path);
  const run: SlideRun = { text };
  if (formatting.bold) {
    run.bold = true;
  }
  if (formatting.italic) {
    run.italic = true;
  }
  if (formatting.underline) {
    run.underline = true;
  }
  if (formatting.fontFamily !== undefined) {
    run.fontFamily = formatting.fontFamily;
  }
  if (formatting.fontSizePt !== undefined) {
    run.sizePt = formatting.fontSizePt;
  }
  if (formatting.color !== undefined) {
    run.color = formatting.color;
  }
  if (formatting.strikethrough) {
    countDrop(drops, "strikethrough");
  }
  if (formatting.highlight !== undefined) {
    countDrop(drops, "highlight");
  }
  return run;
}

/** Text-field content — a plain string or an array of `RunSpec`s — as slide
 * runs. Empty runs are left out: a slide paragraph with no text has no runs. */
export function buildSlideRuns(content: unknown, path: string, drops: SlideRunDrops): SlideRun[] {
  let runs: SlideRun[];
  if (typeof content === "string") {
    runs = [{ text: content }];
  } else if (Array.isArray(content)) {
    runs = (content as RunSpec[]).map((spec, i) => buildSlideRun(spec, `${path}[${i}]`, drops));
  } else {
    throw new DocumentGenerationError(`${path}: expected string or run array`);
  }
  return runs.filter((run) => run.text.length > 0);
}

/** One paragraph holding `runs`, or none when they carry no text — an empty
 * placeholder is born with no paragraphs, and stays that way. */
export function paragraphsOf(runs: SlideRun[]): SlideParagraph[] {
  return runs.length > 0 ? [{ runs }] : [];
}
