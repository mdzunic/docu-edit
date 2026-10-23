// The deck-composition primitives (SPEC-106) — the presentation twin of
// blocks.ts: every slide and per-slide operation the fluent
// `PresentationBuilder` (presentation.ts) and the JSON deck description
// (deckDescription.ts) perform, validated on entry and named by path, so the
// two surfaces map 1:1 because they call the same functions. Slides are built
// from the editor's own pieces — `createBlankPresentation`'s model,
// `titleSlideFrames`/`titleBodyFrames` geometry, SPEC-082's
// `insertImageFrame`/`insertTextBoxFrame` — never from a second table, and no
// serialisation happens here. Imports lib/presentation and lib/core only
// (plus generate/ siblings): never documentModel.ts or a .docx/.odt layer.
import { SLIDE_SIZES, MAX_SLIDE_DIMENSION_CM, MIN_SLIDE_DIMENSION_CM } from "../core/pageSetup";
import { decodeImageDimensions } from "../core/imageDimensions";
import { base64ToBytes, bytesToBase64 } from "../core/packageParts";
import { pxToCm } from "../core/units";
import { insertImageFrame, insertTextBoxFrame, type FrameRect } from "../presentation/frames";
import type { PresentationFormat, PresentationModel, Slide, SlideFrame, SlideParagraph, TextFrame } from "../presentation/model";
import { createBlankPresentation, r3, titleBodyFrames, titleSlideFrames } from "../presentation/slides";
import { MAX_INDENT_LEVEL } from "../presentation/text";
import { slideLocation, type DeckDropRecord } from "./deckDrops";
import { DocumentGenerationError } from "./errors";
import { emuToCm, lengthToEmu, type LengthInput } from "./length";
import { contentFromTextOrRuns, type RunSpec } from "./runs";
import { buildSlideRuns, paragraphsOf, type SlideRunDrops } from "./slideRuns";
import {
  expectArray,
  expectEnum,
  expectHexColor,
  expectIntInRange,
  expectPlainObject,
  expectString,
  rejectUnknownKeys,
} from "./validate";

export const PRESENTATION_FORMATS: readonly PresentationFormat[] = ["pptx", "odp"];

/** The named slide sizes a deck may be generated at — SPEC-090's catalogue
 * entries, restated by name only (their dimensions come from `SLIDE_SIZES`). */
export const DECK_SIZE_NAMES = ["16:9", "4:3"] as const;
export type DeckSizeName = (typeof DECK_SIZE_NAMES)[number];

/** Text-field content: a plain string or an array of SPEC-045 `RunSpec`s. */
export type SlideTextContent = string | RunSpec[];

/** A bullet row: a string, a `{ text | runs, level? }` row (`level` 0–8,
 * absolute), or a nested array whose rows sit one level deeper than the array
 * holding it (clamped at `MAX_INDENT_LEVEL`). */
export type BulletRow = string | { text?: string; runs?: RunSpec[]; level?: number } | BulletRow[];

export interface BulletOptions {
  /** `"bullet"` (default) marks every row with a disc, `"number"` numbers them. */
  style?: "bullet" | "number";
}

/** A placed rectangle: every field a `LengthInput`. */
export interface RectSpec {
  x: LengthInput;
  y: LengthInput;
  width: LengthInput;
  height: LengthInput;
}

/** An image for `addImageSlide`: fitted into the body rect at its own aspect
 * (or the given size, shrunk to fit), centred. `data` is base64 (raw bytes
 * are accepted too, outside JSON). */
export interface SlideImageSpec {
  data: string | Uint8Array;
  contentType: "image/png" | "image/jpeg" | "image/gif";
  width?: LengthInput;
  height?: LengthInput;
  altText?: string;
}

/** An image for `addImage`: an explicit rect (height from the image's aspect
 * when omitted), clamped into the slide. */
export interface PlacedImageSpec {
  data: string | Uint8Array;
  contentType: "image/png" | "image/jpeg" | "image/gif";
  x: LengthInput;
  y: LengthInput;
  width: LengthInput;
  height?: LengthInput;
  altText?: string;
}

/** The deck under construction plus everything dropped along the way. */
export interface DeckContext {
  deck: PresentationModel;
  dropped: DeckDropRecord[];
}

const IMAGE_CONTENT_TYPES = ["image/png", "image/jpeg", "image/gif"] as const;
const SLIDE_IMAGE_KEYS = ["data", "contentType", "width", "height", "altText"] as const;
const PLACED_IMAGE_KEYS = ["data", "contentType", "x", "y", "width", "height", "altText"] as const;
const RECT_KEYS = ["x", "y", "width", "height"] as const;
const BULLET_OPTION_KEYS = ["style"] as const;
const BULLET_ROW_KEYS = ["text", "runs", "level"] as const;
const SIZE_KEYS = ["width", "height"] as const;

/** `parent.key`, or just `key` at a description's root (path ""). */
export function childPath(parent: string, key: string): string {
  return parent.length > 0 ? `${parent}.${key}` : key;
}

// --- the deck -----------------------------------------------------------------

function slideDimensionCm(value: unknown, path: string): number {
  const cm = r3(emuToCm(lengthToEmu(value, path)));
  if (cm < MIN_SLIDE_DIMENSION_CM || cm > MAX_SLIDE_DIMENSION_CM) {
    throw new DocumentGenerationError(
      `${path}: expected a slide dimension between ${MIN_SLIDE_DIMENSION_CM} and ${MAX_SLIDE_DIMENSION_CM} cm`,
    );
  }
  return cm;
}

/** A fresh deck from `createBlankPresentation()`'s model with its first slide
 * removed, configured by the validated `title`/`format`/`size` fields of
 * `fields` (the builder's options object, or a description's root). */
export function createDeckContext(fields: Record<string, unknown>, path: string): DeckContext {
  const deck = createBlankPresentation();
  deck.slides = [];
  if (fields.title !== undefined) {
    deck.title = expectString(fields.title, childPath(path, "title"));
  }
  if (fields.format !== undefined) {
    deck.format = expectEnum(fields.format, PRESENTATION_FORMATS, childPath(path, "format"));
  }
  if (fields.size !== undefined) {
    const sizePath = childPath(path, "size");
    if (typeof fields.size === "string") {
      const name = expectEnum(fields.size, DECK_SIZE_NAMES, sizePath);
      deck.slideWidthCm = SLIDE_SIZES[name].widthCm;
      deck.slideHeightCm = SLIDE_SIZES[name].heightCm;
    } else {
      const size = expectPlainObject(fields.size, sizePath);
      rejectUnknownKeys(size, SIZE_KEYS, sizePath);
      deck.slideWidthCm = slideDimensionCm(size.width, `${sizePath}.width`);
      deck.slideHeightCm = slideDimensionCm(size.height, `${sizePath}.height`);
    }
  }
  return { deck, dropped: [] };
}

/** The path a slide about to be appended is named by. */
export function nextSlidePath(ctx: DeckContext): string {
  return `slides[${ctx.deck.slides.length}]`;
}

function appendSlide(ctx: DeckContext, frames: SlideFrame[]): number {
  ctx.deck.slides.push({ id: crypto.randomUUID(), frames });
  return ctx.deck.slides.length - 1;
}

/** Reports a run mapping's dropped fields against slide `slideIndex`. */
function reportRunDrops(ctx: DeckContext, slideIndex: number, drops: SlideRunDrops): void {
  for (const [field, count] of drops) {
    ctx.dropped.push({
      kind: "runFormatting",
      location: slideLocation(slideIndex),
      reason: `${field} dropped from ${count} run${count === 1 ? "" : "s"}: a slide run has no ${field} field`,
    });
  }
}

function textContent(value: unknown, path: string, drops: SlideRunDrops): SlideParagraph[] {
  if (value === undefined) {
    throw new DocumentGenerationError(`${path}: expected string or run array`);
  }
  return paragraphsOf(buildSlideRuns(value, path, drops));
}

// --- adding slides ------------------------------------------------------------

/** A title slide: `createBlankPresentation`'s title + subtitle placeholders. */
export function addTitleSlide(ctx: DeckContext, title: unknown, subtitle: unknown, path: string): void {
  const drops: SlideRunDrops = new Map();
  const titleParagraphs = textContent(title, `${path}.title`, drops);
  const subtitleParagraphs = subtitle === undefined ? [] : textContent(subtitle, `${path}.subtitle`, drops);
  const [titleFrame, subtitleFrame] = titleSlideFrames(ctx.deck.slideWidthCm, ctx.deck.slideHeightCm);
  titleFrame.paragraphs = titleParagraphs;
  subtitleFrame.paragraphs = subtitleParagraphs;
  reportRunDrops(ctx, appendSlide(ctx, [titleFrame, subtitleFrame]), drops);
}

function bulletMarker(options: unknown, path: string): "disc" | "number" {
  if (options === undefined) {
    return "disc";
  }
  const obj = expectPlainObject(options, path);
  rejectUnknownKeys(obj, BULLET_OPTION_KEYS, path);
  if (obj.style === undefined) {
    return "disc";
  }
  return expectEnum(obj.style, ["bullet", "number"] as const, `${path}.style`) === "number" ? "number" : "disc";
}

/** One bullet row's paragraph — or none for a row with no text, since an
 * empty paragraph is something neither format's reader keeps (the slide
 * model's normal form has none). */
function pushBulletParagraph(
  out: SlideParagraph[],
  runsContent: unknown,
  level: number,
  marker: "disc" | "number",
  path: string,
  drops: SlideRunDrops,
): void {
  const runs = buildSlideRuns(runsContent, path, drops);
  if (runs.length === 0) {
    return;
  }
  const paragraph: SlideParagraph = { runs, bullet: marker };
  if (level > 0) {
    paragraph.level = level;
  }
  out.push(paragraph);
}

/** Flattens bullet rows into paragraphs: strings and `{ text | runs, level }`
 * rows at `depth` (an explicit `level` wins), nested arrays one deeper. */
function bulletParagraphs(
  rows: unknown,
  depth: number,
  marker: "disc" | "number",
  path: string,
  drops: SlideRunDrops,
  out: SlideParagraph[],
): SlideParagraph[] {
  expectArray(rows, path).forEach((row, i) => {
    const rowPath = `${path}[${i}]`;
    if (typeof row === "string") {
      pushBulletParagraph(out, row, depth, marker, rowPath, drops);
    } else if (Array.isArray(row)) {
      bulletParagraphs(row, Math.min(depth + 1, MAX_INDENT_LEVEL), marker, rowPath, drops, out);
    } else if (typeof row === "object" && row !== null) {
      const obj = row as Record<string, unknown>;
      rejectUnknownKeys(obj, BULLET_ROW_KEYS, rowPath);
      const level =
        obj.level === undefined ? depth : expectIntInRange(obj.level, `${rowPath}.level`, 0, MAX_INDENT_LEVEL);
      pushBulletParagraph(out, contentFromTextOrRuns(obj, rowPath), level, marker, rowPath, drops);
    } else {
      throw new DocumentGenerationError(
        `${rowPath}: expected string or run array row ({ text | runs, level? }), or a nested array of rows`,
      );
    }
  });
  return out;
}

/** A title + body slide: `addSlide`'s placeholders, bullets in the body. */
export function addBulletSlide(ctx: DeckContext, title: unknown, bullets: unknown, options: unknown, path: string): void {
  const drops: SlideRunDrops = new Map();
  const titleParagraphs = textContent(title, `${path}.title`, drops);
  const marker = bulletMarker(options, path);
  if (bullets === undefined) {
    throw new DocumentGenerationError(`${path}.bullets: expected an array`);
  }
  const body = bulletParagraphs(bullets, 0, marker, `${path}.bullets`, drops, []);
  const [titleFrame, bodyFrame] = titleBodyFrames(ctx.deck.slideWidthCm, ctx.deck.slideHeightCm);
  titleFrame.paragraphs = titleParagraphs;
  bodyFrame.paragraphs = body;
  reportRunDrops(ctx, appendSlide(ctx, [titleFrame, bodyFrame]), drops);
}

/** A title slide over an image: the title placeholder of `addSlide`'s pair,
 * the image fitted into its body rect. */
export function addImageSlide(ctx: DeckContext, title: unknown, image: unknown, path: string): void {
  const drops: SlideRunDrops = new Map();
  const titleParagraphs = textContent(title, `${path}.title`, drops);
  const imagePath = `${path}.image`;
  const spec = expectPlainObject(image, imagePath);
  rejectUnknownKeys(spec, SLIDE_IMAGE_KEYS, imagePath);
  const decoded = decodeImage(spec, imagePath);
  const explicit = {
    widthCm: spec.width === undefined ? undefined : positiveLengthCm(spec.width, `${imagePath}.width`),
    heightCm: spec.height === undefined ? undefined : positiveLengthCm(spec.height, `${imagePath}.height`),
  };
  const altText = altTextOf(spec, imagePath);
  const [titleFrame] = titleBodyFrames(ctx.deck.slideWidthCm, ctx.deck.slideHeightCm);
  titleFrame.paragraphs = titleParagraphs;
  const slideIndex = appendSlide(ctx, [titleFrame]);
  landImage(ctx, slideIndex, decoded, imageInBodyRect(ctx.deck, decoded.natural, explicit));
  reportAltText(ctx, slideIndex, altText);
  reportRunDrops(ctx, slideIndex, drops);
}

/** An empty slide (no placeholders) for the per-slide methods to fill. */
export function addBlankSlide(ctx: DeckContext): void {
  appendSlide(ctx, []);
}

// --- per-slide content --------------------------------------------------------

/** The most recently added slide, which every per-slide method acts on. */
function currentSlide(ctx: DeckContext, method: string): { slide: Slide; index: number; path: string } {
  const index = ctx.deck.slides.length - 1;
  if (index < 0) {
    throw new DocumentGenerationError(
      `${method}: no slide to act on — add one first (addTitleSlide, addBulletSlide, addImageSlide or addBlankSlide)`,
    );
  }
  return { slide: ctx.deck.slides[index], index, path: `slides[${index}]` };
}

function isPlaceholder(frame: SlideFrame, role: TextFrame["role"]): frame is TextFrame {
  return frame.kind === "text" && frame.role === role;
}

/** Appends bullet rows to the current slide's body placeholder, adding one at
 * `addSlide`'s body geometry when the slide has none. */
export function addBullets(ctx: DeckContext, items: unknown, options: unknown, path?: string): void {
  const { slide, index, path: slidePath } = currentSlide(ctx, "addBullets");
  const drops: SlideRunDrops = new Map();
  const marker = bulletMarker(options, path ?? slidePath);
  const rows = bulletParagraphs(items, 0, marker, `${path ?? slidePath}.bullets`, drops, []);
  let body = slide.frames.find((frame): frame is TextFrame => isPlaceholder(frame, "body"));
  if (!body) {
    body = titleBodyFrames(ctx.deck.slideWidthCm, ctx.deck.slideHeightCm)[1];
    slide.frames.push(body);
  }
  body.paragraphs.push(...rows);
  reportRunDrops(ctx, index, drops);
}

function rectCm(rect: unknown, path: string): FrameRect {
  const obj = expectPlainObject(rect, path);
  rejectUnknownKeys(obj, RECT_KEYS, path);
  return {
    xCm: r3(emuToCm(lengthToEmu(obj.x, `${path}.x`))),
    yCm: r3(emuToCm(lengthToEmu(obj.y, `${path}.y`))),
    widthCm: positiveLengthCm(obj.width, `${path}.width`),
    heightCm: positiveLengthCm(obj.height, `${path}.height`),
  };
}

function positiveLengthCm(value: unknown, path: string): number {
  const cm = r3(emuToCm(lengthToEmu(value, path)));
  if (cm <= 0) {
    throw new DocumentGenerationError(`${path}: expected a positive length`);
  }
  return cm;
}

/** Lands a free text box (no placeholder role) holding one paragraph at
 * `rect`, through SPEC-082's `insertTextBoxFrame`. */
export function addTextBox(ctx: DeckContext, content: unknown, rect: unknown, path?: string): void {
  const { slide, index, path: slidePath } = currentSlide(ctx, "addTextBox");
  const boxPath = path ?? `${slidePath}.textBoxes[${slide.frames.filter((f) => f.kind === "text" && !f.role).length}]`;
  const drops: SlideRunDrops = new Map();
  const paragraphs = textContent(content, boxPath, drops);
  const box = rectCm(rect, boxPath);
  const frameId = insertTextBoxFrame(ctx.deck, index, box)!;
  const frame = slide.frames.find((f): f is TextFrame => f.id === frameId)!;
  frame.paragraphs = paragraphs;
  settle(slide);
  reportRunDrops(ctx, index, drops);
}

/** Lands an image at an explicit rect, clamped into the slide (shrunk at its
 * aspect when larger than the slide, then moved fully on), through SPEC-082's
 * `insertImageFrame` and its unchanged media-naming rule. */
export function addImage(ctx: DeckContext, image: unknown, path?: string): void {
  const { slide, index, path: slidePath } = currentSlide(ctx, "addImage");
  const imagePath = path ?? `${slidePath}.images[${slide.frames.filter((f) => f.kind === "image").length}]`;
  const spec = expectPlainObject(image, imagePath);
  rejectUnknownKeys(spec, PLACED_IMAGE_KEYS, imagePath);
  const decoded = decodeImage(spec, imagePath);
  const xCm = r3(emuToCm(lengthToEmu(spec.x, `${imagePath}.x`)));
  const yCm = r3(emuToCm(lengthToEmu(spec.y, `${imagePath}.y`)));
  const widthCm = positiveLengthCm(spec.width, `${imagePath}.width`);
  const heightCm =
    spec.height === undefined
      ? (widthCm * decoded.natural.heightPx) / decoded.natural.widthPx
      : positiveLengthCm(spec.height, `${imagePath}.height`);
  const altText = altTextOf(spec, imagePath);
  landImage(ctx, index, decoded, clampIntoSlide(ctx.deck, { xCm, yCm, widthCm, heightCm }));
  reportAltText(ctx, index, altText);
}

/** Sets the current slide's speaker notes (SPEC-083's `Slide.notes`): one
 * entry per paragraph (an embedded newline starts another, as in the
 * editor's notes panel), trailing empty paragraphs stripped, and no field at
 * all — never `[]` — when nothing text-bearing remains. */
export function setNotes(ctx: DeckContext, paragraphs: unknown, path?: string): void {
  const { slide, path: slidePath } = currentSlide(ctx, "setNotes");
  const notesPath = path ?? `${slidePath}.notes`;
  const lines = expectArray(paragraphs, notesPath).flatMap((entry, i) =>
    expectString(entry, `${notesPath}[${i}]`).split(/\r\n|\r|\n/),
  );
  setSlideNotes(slide, lines);
}

/** `Slide.notes` from plain paragraphs, trailing empties stripped and `[]`
 * never stored. */
export function setSlideNotes(slide: Slide, lines: string[]): void {
  const notes = lines.slice();
  while (notes.length > 0 && notes[notes.length - 1].trim().length === 0) {
    notes.pop();
  }
  if (notes.length > 0) {
    slide.notes = notes;
  } else {
    delete slide.notes;
  }
}

/** Gives the current slide a solid `#rrggbb` background. */
export function setBackground(ctx: DeckContext, color: unknown, path?: string): void {
  const { slide, path: slidePath } = currentSlide(ctx, "setBackground");
  slide.background = { type: "solid", color: expectHexColor(color, path ?? `${slidePath}.background`) };
}

// --- images -------------------------------------------------------------------

export interface DecodedImage {
  contentType: string;
  /** Canonical base64 of the bytes. */
  data: string;
  natural: { widthPx: number; heightPx: number };
}

/** Validates `{ data, contentType }` and decodes the image's pixel size —
 * an image whose dimensions cannot be read is refused, never guessed. */
export function decodeImage(spec: Record<string, unknown>, path: string): DecodedImage {
  const contentType = expectEnum(spec.contentType, IMAGE_CONTENT_TYPES, `${path}.contentType`);
  let bytes: Uint8Array;
  if (spec.data instanceof Uint8Array) {
    bytes = spec.data;
  } else if (typeof spec.data === "string") {
    try {
      bytes = base64ToBytes(spec.data);
    } catch {
      throw new DocumentGenerationError(`${path}.data: not valid base64`);
    }
  } else {
    throw new DocumentGenerationError(`${path}.data: expected base64 image bytes`);
  }
  const natural = decodeImageDimensions(bytes, contentType);
  if (!natural || natural.widthPx <= 0 || natural.heightPx <= 0) {
    throw new DocumentGenerationError(`${path}.data: could not decode image dimensions for ${contentType}`);
  }
  return { contentType, data: bytesToBase64(bytes), natural };
}

/** The rect an image takes inside `addSlide`'s body rect: the given size (the
 * missing side from the image's aspect), else the largest size at the image's
 * aspect — shrunk at its aspect to fit, then centred. */
export function imageInBodyRect(
  deck: PresentationModel,
  natural: { widthPx: number; heightPx: number },
  explicit: { widthCm?: number; heightCm?: number } = {},
): FrameRect {
  const body = titleBodyFrames(deck.slideWidthCm, deck.slideHeightCm)[1];
  const aspect = natural.widthPx / natural.heightPx;
  let widthCm: number;
  let heightCm: number;
  if (explicit.widthCm !== undefined && explicit.heightCm !== undefined) {
    widthCm = explicit.widthCm;
    heightCm = explicit.heightCm;
  } else if (explicit.widthCm !== undefined) {
    widthCm = explicit.widthCm;
    heightCm = widthCm / aspect;
  } else if (explicit.heightCm !== undefined) {
    heightCm = explicit.heightCm;
    widthCm = heightCm * aspect;
  } else {
    widthCm = pxToCm(natural.widthPx);
    heightCm = pxToCm(natural.heightPx);
    const grow = Math.min(body.widthCm / widthCm, body.heightCm / heightCm);
    widthCm *= grow;
    heightCm *= grow;
  }
  const shrink = Math.min(1, body.widthCm / widthCm, body.heightCm / heightCm);
  widthCm = r3(widthCm * shrink);
  heightCm = r3(heightCm * shrink);
  return {
    xCm: r3(body.xCm + (body.widthCm - widthCm) / 2),
    yCm: r3(body.yCm + (body.heightCm - heightCm) / 2),
    widthCm,
    heightCm,
  };
}

/** `rect` shrunk at its aspect to fit the slide, then moved fully onto it. */
function clampIntoSlide(deck: PresentationModel, rect: FrameRect): FrameRect {
  const shrink = Math.min(1, deck.slideWidthCm / rect.widthCm, deck.slideHeightCm / rect.heightCm);
  const widthCm = r3(rect.widthCm * shrink);
  const heightCm = r3(rect.heightCm * shrink);
  return {
    xCm: r3(Math.min(Math.max(rect.xCm, 0), deck.slideWidthCm - widthCm)),
    yCm: r3(Math.min(Math.max(rect.yCm, 0), deck.slideHeightCm - heightCm)),
    widthCm,
    heightCm,
  };
}

/** Lands a decoded image on slide `slideIndex` through `insertImageFrame`. */
export function landImage(ctx: DeckContext, slideIndex: number, image: DecodedImage, rect: FrameRect): void {
  insertImageFrame(ctx.deck, slideIndex, { mediaId: "", contentType: image.contentType, data: image.data }, rect);
  settle(ctx.deck.slides[slideIndex]);
}

function altTextOf(spec: Record<string, unknown>, path: string): string | undefined {
  return spec.altText === undefined ? undefined : expectString(spec.altText, `${path}.altText`);
}

/** An image's alt text is accepted (SPEC-045's image shape) and validated, but
 * SPEC-079's `ImageFrame` has nowhere to keep it — so it is reported, never
 * silently swallowed. */
function reportAltText(ctx: DeckContext, slideIndex: number, altText: string | undefined): void {
  if (altText !== undefined && altText.length > 0) {
    ctx.dropped.push({
      kind: "media",
      location: slideLocation(slideIndex),
      reason: "image alt text not kept: a slide image frame has no alt-text field",
    });
  }
}

/** SPEC-082's insert commands mark the slide edited — a flag that only means
 * something for a deck with a retained package to rebuild against. A
 * generated deck has none, so the flag is cleared again. */
function settle(slide: Slide): void {
  delete slide.dirty;
}
