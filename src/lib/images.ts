// Image structure commands (SPEC-024), framework-free like every lib
// module. An image is a block-level object (documentModel.Image) addressed
// the same way a table is — see textEditing.ts's blockLength/clampCaret,
// which treat both as length-1 atoms. This module owns inserting an image at
// the caret, resizing/re-wrapping/re-labelling one, and deleting one with
// reference-counted media cleanup (AC-31, AC-32) — the character-level
// commands stay in commands.ts/textEditing.ts, exactly like tables.ts.
import {
  isImage,
  type DocumentModel,
  type Image,
  type ImageWrapMode,
  type MediaPart,
  type Paragraph,
  type Table,
} from "./documentModel";
import { blockLength, clampCaret, type CaretPosition } from "./textEditing";
import { caretSelection, type CommandResult, type ModelSelection } from "./commands";

/** Images larger than this are rejected at insertion time (AC-33). */
export const MAX_IMAGE_SIZE_BYTES = 10 * 1024 * 1024;

/** Shown when a rejected image is too large (AC-33, AC-34). */
export const IMAGE_TOO_LARGE_MESSAGE = "That image is too large (limit 10 MB)";

/** File types "Insert image" accepts, as both MIME types and an `<input
 * accept>` string (AC-1). */
export const ACCEPTED_IMAGE_MIME_TYPES = [
  "image/png",
  "image/jpeg",
  "image/gif",
  "image/svg+xml",
] as const;

export const ACCEPTED_IMAGE_ACCEPT_ATTR = ACCEPTED_IMAGE_MIME_TYPES.join(",");

/** Whether a file's declared MIME type is one "Insert image" accepts. */
export function isAcceptedImageType(mimeType: string): boolean {
  return (ACCEPTED_IMAGE_MIME_TYPES as readonly string[]).includes(mimeType);
}

/** Whether a candidate image's byte length exceeds the 10 MB limit (AC-33). */
export function isImageTooLarge(byteLength: number): boolean {
  return byteLength > MAX_IMAGE_SIZE_BYTES;
}

// commands.ts's container clone, restated for the reason clipboard.ts and
// tables.ts each restate it: the undo-registry guard test bars commands.ts
// from exporting helpers.
function cloneContainers(model: DocumentModel): DocumentModel {
  return {
    ...model,
    sections: model.sections.map((section) => ({
      ...section,
      blocks: [...section.blocks],
    })),
  };
}

function emptyParagraph(): Paragraph {
  return { type: "paragraph", runs: [] };
}

// A paragraph split into two around a character offset — restated from
// tables.ts's splitParagraphAt (that primitive writes into the model in
// place, this one doesn't touch the model yet).
function splitParagraphAt(paragraph: Paragraph, offset: number): [Paragraph, Paragraph] {
  let consumed = 0;
  const before: Paragraph = { ...paragraph, runs: [] };
  const after: Paragraph = {
    ...paragraph,
    runs: [],
    formatting: paragraph.formatting ? { ...paragraph.formatting } : undefined,
  };
  const beforeRuns = [];
  const afterRuns = [];
  for (const run of paragraph.runs) {
    const end = consumed + run.text.length;
    if (end <= offset) {
      beforeRuns.push(run);
    } else if (consumed >= offset) {
      afterRuns.push(run);
    } else {
      const cut = offset - consumed;
      beforeRuns.push({ text: run.text.slice(0, cut), formatting: { ...run.formatting } });
      afterRuns.push({ text: run.text.slice(cut), formatting: { ...run.formatting } });
    }
    consumed = end;
  }
  before.runs = beforeRuns;
  after.runs = afterRuns;
  return [before, after];
}

// 96 DPI is the fixed pixel/inch ratio SPEC-024 defines "Reset size" by
// (AC-21) — the same ratio browsers use for unstyled `<img>` natural size.
const PX_PER_INCH = 96;
const CM_PER_INCH = 2.54;

/** A natural pixel dimension to centimetres at 96 DPI (AC-21). */
export function pxToCm(px: number): number {
  return (px / PX_PER_INCH) * CM_PER_INCH;
}

/** The inverse of pxToCm, for callers that need to compare against a pixel
 * measurement (e.g. a resize handle's on-screen delta). */
export function cmToPx(cm: number): number {
  return (cm / CM_PER_INCH) * PX_PER_INCH;
}

/** The page's text area width in cm (page width minus its left/right
 * margins) — what an inserted image scales down to fit (AC-5, AC-6). Mirrors
 * tables.ts's equalColumnWidthsCm's own textAreaCm computation. */
export function textAreaWidthCm(model: DocumentModel): number {
  const { pageSize, margins } = model.defaultStyles;
  return Math.max(1, pageSize.widthCm - margins.leftCm - margins.rightCm);
}

/** The display size a freshly inserted image gets (AC-5, AC-6): its natural
 * size, scaled down to fit `textAreaCm` when it would otherwise overflow —
 * aspect ratio always preserved. */
export function fitImageSizeCm(
  naturalWidthPx: number,
  naturalHeightPx: number,
  textAreaCm: number,
): { widthCm: number; heightCm: number } {
  const naturalWidthCm = pxToCm(naturalWidthPx);
  const naturalHeightCm = pxToCm(naturalHeightPx);
  if (naturalWidthCm <= textAreaCm || naturalWidthCm <= 0) {
    return { widthCm: naturalWidthCm, heightCm: naturalHeightCm };
  }
  const scale = textAreaCm / naturalWidthCm;
  return { widthCm: textAreaCm, heightCm: naturalHeightCm * scale };
}

/** The smallest width/height a resize (handle drag or the options panel) can
 * leave an image — no AC specifies one, so this just keeps the image
 * visible/selectable rather than collapsing to nothing. */
export const MIN_IMAGE_SIZE_CM = 0.2;

/** A freshly uploaded/pasted image's media payload — the caller (the file
 * picker/drop/paste handler) has already read the bytes, converted SVG to
 * PNG for a `.docx`-bound document (AC-2) and enforced the size limit
 * (AC-33) before this ever reaches a command. */
export interface NewImageMedia {
  mediaId: string;
  contentType: string;
  /** Base64-encoded image bytes (DocumentModel stays JSON-serialisable). */
  data: string;
}

/** Inserts an image at the caret (AC-4), splitting the caret's paragraph
 * around the insertion point exactly like tables.ts's insertTable — the
 * trailing half always survives so a paragraph follows the image and the
 * caret can never be trapped below it. A non-collapsed selection first
 * collapses to its focus. No-op inside a table cell: images are top-level
 * blocks only (mirrors insertTable's own no-nesting rule). */
export function insertImage(
  model: DocumentModel,
  selection: ModelSelection,
  media: NewImageMedia,
  naturalWidthPx: number,
  naturalHeightPx: number,
  wrap: ImageWrapMode = "inline",
): CommandResult {
  const next = cloneContainers(model);
  const at = clampCaret(next, selection.focus);
  if (at.cell) {
    return { model: next, selection: caretSelection(at) };
  }
  const blocks = next.sections[at.section].blocks;
  const target = blocks[at.block];
  const { widthCm, heightCm } = fitImageSizeCm(naturalWidthPx, naturalHeightPx, textAreaWidthCm(next));
  const image: Image = {
    type: "image",
    mediaId: media.mediaId,
    naturalWidthPx,
    naturalHeightPx,
    widthCm,
    heightCm,
    wrap,
  };
  next.media = [...(next.media ?? []), { mediaId: media.mediaId, contentType: media.contentType, data: media.data }];

  const replacement: (Paragraph | Table | Image)[] = [];
  let imageBlockOffset: number;
  if (target.type !== "paragraph") {
    // The caret addresses a table or (another) image as an atom (no `cell`,
    // offset 0 = before it, 1 = after) — never split, just land the new
    // image on the requested side, with its own trailing paragraph when it
    // lands after (mirrors insertTable's guaranteed-following-paragraph
    // rule so the caret can never be trapped below it).
    if (at.offset === 0) {
      replacement.push(image, target);
      imageBlockOffset = 0;
    } else {
      replacement.push(target, image, emptyParagraph());
      imageBlockOffset = 1;
    }
  } else if (at.offset === 0) {
    replacement.push(image, target);
    imageBlockOffset = 0;
  } else if (at.offset >= blockLength(target)) {
    replacement.push(target, image, emptyParagraph());
    imageBlockOffset = 1;
  } else {
    const [before, after] = splitParagraphAt(target, at.offset);
    replacement.push(before, image, after);
    imageBlockOffset = 1;
  }
  blocks.splice(at.block, 1, ...replacement);
  const imageIndex = at.block + imageBlockOffset;
  const landing: CaretPosition = { section: at.section, block: imageIndex + 1, offset: 0 };
  return { model: next, selection: caretSelection(landing) };
}

/** Where an image command applies. */
export interface ImageLocation {
  section: number;
  block: number;
}

/** Typed `null` seed for an image-location ref/variable in `.vue` scripts
 * (espree bars the annotation — see commands.ts's noSelection). */
export function noImageLocation(): ImageLocation | null {
  return null;
}

/** An image command bound late to the current model/selection/location —
 * how DocumentView.vue's applyImageEdit hands the wrap/resize/alt-text
 * commands to one shared apply step without writing the function-type
 * annotation espree can't parse (mirrors commands.ts's ParagraphCommand). */
export type ImageCommand = (model: DocumentModel, selection: ModelSelection, location: ImageLocation) => CommandResult;

/** Typed `null` seed for ImageCommand parameters in `.vue` scripts. */
export function noImageCommand(): ImageCommand | null {
  return null;
}

/** The image a location addresses, or null. */
export function imageAt(model: DocumentModel, location: ImageLocation): Image | null {
  const target = model.sections[location.section]?.blocks[location.block];
  return isImage(target) ? target : null;
}

/** A media part's bytes as a `data:` URI for the canvas `<img>` — resolved
 * only for rendering; AC-27 bars base64 data URIs from the saved XML, which
 * this never touches. "" for an unresolved/absent part. */
export function mediaDataUri(media: MediaPart | undefined): string {
  return media ? `data:${media.contentType};base64,${media.data}` : "";
}

function replaceImage(model: DocumentModel, location: ImageLocation, next: Image): void {
  model.sections[location.section].blocks.splice(location.block, 1, next);
}

/** Every media id an image block in the document still references. */
function referencedMediaIds(model: DocumentModel): Set<string> {
  const ids = new Set<string>();
  for (const section of model.sections) {
    for (const block of section.blocks) {
      if (isImage(block)) {
        ids.add(block.mediaId);
      }
    }
  }
  return ids;
}

/** Drops every `model.media` entry no image block references any more — the
 * reference-counted cleanup AC-31/AC-32 describe as happening "on save", but
 * performed here (at the point a use is actually removed) so the model
 * itself never carries dead media, and so undo restores it for free (undo
 * snapshots the whole model, see commands.ts's header comment). An image
 * used twice keeps its media entry as long as either use remains (AC-32),
 * since referencedMediaIds counts every remaining image block. */
function pruneUnreferencedMedia(model: DocumentModel): void {
  if (!model.media || model.media.length === 0) {
    return;
  }
  const referenced = referencedMediaIds(model);
  const pruned = model.media.filter((part) => referenced.has(part.mediaId));
  if (pruned.length !== model.media.length) {
    model.media = pruned;
  }
}

/** Sets an image's display size in cm (AC-16, AC-19, AC-20; also what a
 * resize-handle drag commits on release, AC-10..12). Clamped to
 * MIN_IMAGE_SIZE_CM so a drag/field edit can never collapse the image to
 * nothing. */
export function setImageSize(
  model: DocumentModel,
  selection: ModelSelection,
  location: ImageLocation,
  widthCm: number,
  heightCm: number,
): CommandResult {
  const next = cloneContainers(model);
  const image = imageAt(next, location);
  if (!image) {
    return { model: next, selection };
  }
  replaceImage(next, location, {
    ...image,
    widthCm: Math.max(MIN_IMAGE_SIZE_CM, widthCm),
    heightCm: Math.max(MIN_IMAGE_SIZE_CM, heightCm),
  });
  return { model: next, selection };
}

/** "Reset size" (AC-21): restores the image's natural pixel size, converted
 * to cm at 96 DPI. */
export function resetImageSize(
  model: DocumentModel,
  selection: ModelSelection,
  location: ImageLocation,
): CommandResult {
  const next = cloneContainers(model);
  const image = imageAt(next, location);
  if (!image) {
    return { model: next, selection };
  }
  replaceImage(next, location, {
    ...image,
    widthCm: pxToCm(image.naturalWidthPx),
    heightCm: pxToCm(image.naturalHeightPx),
  });
  return { model: next, selection };
}

/** The image toolbar's wrap-mode picker (AC-13, AC-14). */
export function setImageWrap(
  model: DocumentModel,
  selection: ModelSelection,
  location: ImageLocation,
  wrap: ImageWrapMode,
): CommandResult {
  const next = cloneContainers(model);
  const image = imageAt(next, location);
  if (!image) {
    return { model: next, selection };
  }
  replaceImage(next, location, { ...image, wrap });
  return { model: next, selection };
}

/** The "Image options" panel's alt-text field (AC-7, AC-15). An empty string
 * clears it, matching the field's own "cleared input" meaning rather than
 * storing an empty-but-present description. */
export function setImageAltText(
  model: DocumentModel,
  selection: ModelSelection,
  location: ImageLocation,
  altText: string,
): CommandResult {
  const next = cloneContainers(model);
  const image = imageAt(next, location);
  if (!image) {
    return { model: next, selection };
  }
  const updated: Image = { ...image };
  if (altText.length > 0) {
    updated.altText = altText;
  } else {
    delete updated.altText;
  }
  replaceImage(next, location, updated);
  return { model: next, selection };
}

/** Deletes a selected image (AC-30): the caret lands where the image was —
 * the same rule tables.ts's removeTableBlock uses. Its media part is pruned
 * from the model only if no other image block still references it (AC-31,
 * AC-32). */
export function deleteImage(
  model: DocumentModel,
  selection: ModelSelection,
  location: ImageLocation,
): CommandResult {
  const next = cloneContainers(model);
  if (!imageAt(next, location)) {
    return { model: next, selection };
  }
  const blocks = next.sections[location.section].blocks;
  blocks.splice(location.block, 1);
  if (blocks.length === 0) {
    blocks.push(emptyParagraph());
  }
  pruneUnreferencedMedia(next);
  const landing = Math.min(location.block, blocks.length - 1);
  return {
    model: next,
    selection: caretSelection({ section: location.section, block: landing, offset: 0 }),
  };
}
