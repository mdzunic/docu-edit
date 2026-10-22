// Image structure commands (SPEC-024), framework-free like every lib
// module. An image is a block-level object (documentModel.Image) addressed
// the same way a table is — see textEditing.ts's blockLength/clampCaret,
// which treat both as length-1 atoms. This module owns inserting an image at
// the caret, resizing/re-wrapping/re-labelling one, and deleting one with
// reference-counted media cleanup (AC-31, AC-32) — the character-level
// commands stay in commands.ts/textEditing.ts, exactly like tables.ts.
import {
  cloneContainers,
  isImage,
  splitParagraphAt,
  type Block,
  type DocumentModel,
  type Image,
  type ImageWrapMode,
  type Paragraph,
} from "./documentModel";
import { blockLength, clampCaret, type CaretPosition } from "./textEditing";
import { caretSelection, type CommandResult, type ModelSelection } from "./commands";
import { referencedMediaIds } from "./mediaRefs";

function emptyParagraph(): Paragraph {
  return { type: "paragraph", runs: [] };
}

import { pxToCm } from "./core/units";
import { fitImageSizeCm, MIN_IMAGE_SIZE_CM } from "./core/media";

/** The page's text area width in cm (page width minus its left/right
 * margins) — what an inserted image scales down to fit (AC-5, AC-6). Mirrors
 * tables.ts's equalColumnWidthsCm's own textAreaCm computation. */
export function textAreaWidthCm(model: DocumentModel): number {
  const { pageSize, margins } = model.defaultStyles;
  return Math.max(1, pageSize.widthCm - margins.leftCm - margins.rightCm);
}

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

  const replacement: Block[] = [];
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


function replaceImage(model: DocumentModel, location: ImageLocation, next: Image): void {
  model.sections[location.section].blocks.splice(location.block, 1, next);
}

/** Drops every `model.media` entry no image block references any more — the
 * reference-counted cleanup AC-31/AC-32 describe as happening "on save", but
 * performed here (at the point a use is actually removed) so the model
 * itself never carries dead media, and so undo restores it for free (undo
 * snapshots the whole model, see commands.ts's header comment). An image
 * used twice keeps its media entry as long as either use remains (AC-32),
 * since mediaRefs.ts's referencedMediaIds counts every remaining image
 * block. */
function pruneUnreferencedMedia(model: DocumentModel): void {
  if (!model.media || model.media.length === 0) {
    return;
  }
  const referenced = new Set(referencedMediaIds(model));
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
