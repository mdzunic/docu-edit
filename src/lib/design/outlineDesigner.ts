// The model-free deck designer (SPEC-115) — a first-class designer, not a
// stub: deterministic, pure, and needing neither a model nor a network. It is
// what "Outline this document" runs when the provider is off, unreachable or
// declined, and the designer the end-to-end no-provider path exercises.
//
// Its input is `OutlineSource`, plain product-neutral data: a title, an
// optional theme name (a deck source's), and the source's blocks in document
// order — headings and paragraphs/list items. The app layer
// (src/views/design/outlineSource.ts) builds it from a text document's model
// (SPEC-037's heading outline and the paragraphs between the headings) or
// from pasted text, so this module never imports the text product.
//
// The rule table, in order (outlineDesigner.spec.ts tests each):
//   1. A title slide from the source's title, its subtitle the paragraph that
//      directly follows the first Heading 1, when there is one (that paragraph
//      is then not repeated as a bullet).
//   2. One section slide per Heading 1 that has Heading 2 children, its
//      bullets the paragraphs between it and its first Heading 2; a Heading 1
//      with none becomes a content slide.
//   3. One content slide per Heading 2, its bullets taken from the following
//      paragraphs and list items — at most 6, each cut at 12 words on a word
//      boundary with an ellipsis. Text that sits under no heading (all of a
//      document with no headings) becomes one content slide, titled with the
//      source's title, ahead of the rest.
//   4. Headings deeper than 2 fold into their parent's bullets, indented one
//      level per depth (SPEC-106's { text, level } rows); a paragraph under a
//      folded heading sits one level below it, plus its own list depth.
//   5. Layout per slide: title for 1, sectionHeader for 2, titleAndBody for a
//      slide with bullets, titleOnly for one without.
//   6. Theme: the source deck's (a built-in name), else the product default.
//      Size: 16:9.
// The plan is bounded at `maxSlides` (default and ceiling MAX_PLANNED_SLIDES).
import { DEFAULT_THEME } from "../core/theme";
import { DEFAULT_PRESENTATION_TITLE } from "../presentation/model";
import { MAX_INDENT_LEVEL } from "../presentation/text";
import {
  DEFAULT_PLAN_SIZE,
  MAX_BULLET_CHARS,
  MAX_PLANNED_SLIDES,
  makeBullet,
  normaliseThemeName,
  type DeckPlan,
  type PlannedBullet,
  type PlannedSlide,
} from "./plan";

/** One block of a source, in document order. `listLevel` is a list item's
 * 0-based nesting depth (absent for a body paragraph). */
export type OutlineBlock =
  | { kind: "heading"; level: number; text: string }
  | { kind: "paragraph"; text: string; listLevel?: number };

export interface OutlineSource {
  title: string;
  /** A deck source's theme name; absent for a text document or pasted text. */
  theme?: string;
  blocks: OutlineBlock[];
}

export const MAX_OUTLINE_BULLETS = 6;
export const MAX_OUTLINE_BULLET_WORDS = 12;

/** The GeneratedModel a deck from this designer records. */
export const OUTLINE_DESIGNER_MODEL = "outline-designer";

/** `text` with its whitespace collapsed, cut at `maxWords` words on a word
 * boundary with an ellipsis — and within a plan's bullet length. */
export function cutWords(text: string, maxWords = MAX_OUTLINE_BULLET_WORDS): string {
  const words = text.trim().split(/\s+/).filter((word) => word !== "");
  let cut = words.length > maxWords ? `${words.slice(0, maxWords).join(" ")}…` : words.join(" ");
  const chars = Array.from(cut);
  if (chars.length > MAX_BULLET_CHARS) {
    cut = `${chars.slice(0, MAX_BULLET_CHARS - 1).join("")}…`;
  }
  return cut;
}

interface SlideDraft {
  title: string;
  section: boolean;
  bullets: PlannedBullet[];
  /** Levels of the deeper headings folded so far, innermost last (rule 4). */
  folds: number[];
}

function addBullet(draft: SlideDraft, text: string, level: number): void {
  if (draft.bullets.length < MAX_OUTLINE_BULLETS && text.trim() !== "") {
    draft.bullets.push(makeBullet(cutWords(text), Math.min(level, MAX_INDENT_LEVEL)));
  }
}

function finish(draft: SlideDraft): PlannedSlide {
  const layout = draft.section ? "sectionHeader" : draft.bullets.length > 0 ? "titleAndBody" : "titleOnly";
  return { layout, title: draft.title, bullets: draft.bullets };
}

/** Whether the Heading 1 at `index` has a Heading 2 before the next Heading 1. */
function hasHeading2Child(blocks: OutlineBlock[], index: number): boolean {
  for (let i = index + 1; i < blocks.length; i++) {
    const block = blocks[i];
    if (block.kind === "heading" && block.level <= 2) {
      return block.level === 2;
    }
  }
  return false;
}

export function planFromOutline(source: OutlineSource, options: { maxSlides?: number } = {}): DeckPlan {
  const maxSlides = Math.min(Math.max(Math.trunc(options.maxSlides ?? MAX_PLANNED_SLIDES), 1), MAX_PLANNED_SLIDES);
  const blocks = source.blocks.filter((block) => block.text.trim() !== "");
  const title = source.title.trim() || DEFAULT_PRESENTATION_TITLE;

  // Rule 1: the title slide, and the subtitle it takes from the first H1.
  const firstH1 = blocks.findIndex((block) => block.kind === "heading" && block.level === 1);
  const afterFirstH1 = firstH1 >= 0 ? blocks[firstH1 + 1] : undefined;
  const subtitleIndex = afterFirstH1?.kind === "paragraph" ? firstH1 + 1 : -1;
  const slides: PlannedSlide[] = [
    { layout: "title", title, bullets: subtitleIndex >= 0 ? [cutWords(blocks[subtitleIndex].text)] : [] },
  ];

  let current: SlideDraft | null = null;
  const flush = (): void => {
    if (current !== null && slides.length < maxSlides) {
      slides.push(finish(current));
    }
    current = null;
  };

  for (let i = 0; i < blocks.length && slides.length < maxSlides; i++) {
    const block = blocks[i];
    if (block.kind === "heading" && block.level <= 2) {
      // Rules 2 and 3: a new slide per Heading 1 and Heading 2.
      flush();
      current = {
        title: block.text.trim(),
        section: block.level === 1 && hasHeading2Child(blocks, i),
        bullets: [],
        folds: [],
      };
      continue;
    }
    // Rule 3's text under no heading: one slide titled with the source title.
    current ??= { title, section: false, bullets: [], folds: [] };
    const draft: SlideDraft = current;
    if (block.kind === "heading") {
      // Rule 4: a deeper heading folds into its parent's bullets.
      while (draft.folds.length > 0 && draft.folds[draft.folds.length - 1] >= block.level) {
        draft.folds.pop();
      }
      draft.folds.push(block.level);
      addBullet(draft, block.text, draft.folds.length);
    } else if (i !== subtitleIndex) {
      const base = draft.folds.length === 0 ? 0 : draft.folds.length + 1;
      addBullet(draft, block.text, base + Math.max(0, Math.trunc(block.listLevel ?? 0)));
    }
  }
  flush();

  return {
    title,
    theme: normaliseThemeName(source.theme) ?? DEFAULT_THEME.name,
    size: DEFAULT_PLAN_SIZE,
    slides,
  };
}
