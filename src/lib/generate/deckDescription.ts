// The JSON deck-description format (SPEC-106): a plain-JSON object mapping
// 1:1 onto `PresentationBuilder` — each slide object names a `layout` and
// carries that layout's fields, interpreted through the same deckSlides.ts
// primitives the builder's methods call, so a deck can be generated from
// data with no imperative code. The published JSON Schema
// (schema/deck-description.json, beside SPEC-045's
// schema/document-description.json) describes this shape for external
// tooling; this file is its runtime validator. Every object is validated
// closed: an unknown top-level or per-slide key is a validation error.
import type { PresentationModel } from "../presentation/model";
import {
  addBlankSlide,
  addBulletSlide,
  addImage,
  addImageSlide,
  addTextBox,
  addTitleSlide,
  createDeckContext,
  setBackground,
  setNotes,
  type BulletRow,
  type DeckContext,
  type DeckSizeName,
  type PlacedImageSpec,
  type SlideImageSpec,
  type SlideTextContent,
} from "./deckSlides";
import type { LengthInput } from "./length";
import type { RunSpec } from "./runs";
import { contentFromTextOrRuns } from "./runs";
import { expectArray, expectEnum, expectPlainObject, rejectUnknownKeys } from "./validate";

const DESCRIPTION_KEYS = ["title", "format", "size", "slides"] as const;

export const DECK_LAYOUTS = ["title", "bullets", "image", "blank"] as const;

/** Each layout's allowed keys (beyond none): the Reference table, plus the
 * bullets layout's `style` — `addBulletSlide`'s options, so the mapping
 * stays 1:1. */
export const DECK_LAYOUT_KEYS: Record<(typeof DECK_LAYOUTS)[number], readonly string[]> = {
  title: ["layout", "title", "subtitle", "notes"],
  bullets: ["layout", "title", "bullets", "style", "notes"],
  image: ["layout", "title", "image", "notes"],
  blank: ["layout", "textBoxes", "images", "background", "notes"],
};

const TEXT_BOX_KEYS = ["text", "runs", "x", "y", "width", "height"] as const;

/** A positioned text box on a blank slide: `text` or `runs`, and its rect. */
export interface TextBoxDescription {
  text?: string;
  runs?: RunSpec[];
  x: LengthInput;
  y: LengthInput;
  width: LengthInput;
  height: LengthInput;
}

export type DeckSlideDescription =
  | { layout: "title"; title: SlideTextContent; subtitle?: SlideTextContent; notes?: string[] }
  | { layout: "bullets"; title: SlideTextContent; bullets: BulletRow[]; style?: "bullet" | "number"; notes?: string[] }
  | { layout: "image"; title: SlideTextContent; image: SlideImageSpec; notes?: string[] }
  | {
      layout: "blank";
      textBoxes?: TextBoxDescription[];
      images?: PlacedImageSpec[];
      background?: string;
      notes?: string[];
    };

/** A JSON deck description: every field optional — an object with no
 * `slides` describes an empty deck. */
export interface DeckDescription {
  title?: string;
  format?: "pptx" | "odp";
  size?: DeckSizeName | { width: LengthInput; height: LengthInput };
  slides?: DeckSlideDescription[];
}

/** Validates and interprets a description into the builder's context — the
 * shared body of `deckFromDescription` and `PresentationBuilder.fromDescription`.
 * Throws `DocumentGenerationError` on the first invalid field, named by its
 * path from the description's root ("slides[2].bullets[0]: …"). */
export function deckContextFromDescription(description: unknown): DeckContext {
  const obj = expectPlainObject(description, "");
  rejectUnknownKeys(obj, DESCRIPTION_KEYS, "");
  const ctx = createDeckContext(obj, "");
  if (obj.slides === undefined) {
    return ctx;
  }
  expectArray(obj.slides, "slides").forEach((raw, i) => {
    const path = `slides[${i}]`;
    const slide = expectPlainObject(raw, path);
    const layout = expectEnum(slide.layout, DECK_LAYOUTS, `${path}.layout`);
    rejectUnknownKeys(slide, DECK_LAYOUT_KEYS[layout], path);
    if (layout === "title") {
      addTitleSlide(ctx, slide.title, slide.subtitle, path);
    } else if (layout === "bullets") {
      addBulletSlide(ctx, slide.title, slide.bullets, slide.style === undefined ? undefined : { style: slide.style }, path);
    } else if (layout === "image") {
      addImageSlide(ctx, slide.title, slide.image, path);
    } else {
      addBlankSlide(ctx);
      if (slide.textBoxes !== undefined) {
        expectArray(slide.textBoxes, `${path}.textBoxes`).forEach((rawBox, k) => {
          const boxPath = `${path}.textBoxes[${k}]`;
          const box = expectPlainObject(rawBox, boxPath);
          rejectUnknownKeys(box, TEXT_BOX_KEYS, boxPath);
          const rect = { x: box.x, y: box.y, width: box.width, height: box.height };
          addTextBox(ctx, contentFromTextOrRuns(box, boxPath), rect, boxPath);
        });
      }
      if (slide.images !== undefined) {
        expectArray(slide.images, `${path}.images`).forEach((image, k) => addImage(ctx, image, `${path}.images[${k}]`));
      }
      if (slide.background !== undefined) {
        setBackground(ctx, slide.background, `${path}.background`);
      }
    }
    if (slide.notes !== undefined) {
      setNotes(ctx, slide.notes, `${path}.notes`);
    }
  });
  return ctx;
}

/** Validates and interprets a JSON deck description into the same
 * `PresentationModel` the fluent builder produces for the same content —
 * never writing a byte first. */
export function deckFromDescription(description: unknown): PresentationModel {
  return structuredClone(deckContextFromDescription(description).deck);
}
