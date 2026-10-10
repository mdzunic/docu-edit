// "Convert to PDF…" (SPEC-108): one library image onto one assembly sheet,
// written as a one-page PDF — the sheet geometry's end-to-end proof. The
// geometry is core/sheetLayout's, the controls are useSheetSetup's, and the
// file is the existing SPEC-034 writer's: this module adds no serialisation
// of its own. It builds one PdfPageInput (a background rect plus the image
// at the placed rectangle), has the image encoded by encodeImage — after any
// cover/actual crop, so a cut-away region never reaches the file — and hands
// both to buildPdf.
//
// The source is only ever read: its bytes are copied out of the MediaPart,
// and a save stores a NEW media document, titled "<title>.pdf" with
// SPEC-003's duplicate numbering. Both save and download report a failure
// without having written anything.
import { computed, ref, shallowRef, type ComputedRef, type Ref } from "vue";
import { buildPdf, PDF_MIME_TYPE, pdfFileName, type PdfImageData, type PdfPageInput } from "../../lib/pdf/buildPdf";
import { cropImage, encodeImage } from "../../lib/pdf/imageEncode";
import { convertSvgToPng, type ConvertedPng } from "../../lib/core/svgToPng";
import { triggerBlobDownload } from "../../lib/core/triggerDownload";
import { bytesToBase64 } from "../../lib/core/packageParts";
import { decodeImageDimensions } from "../../lib/core/imageDimensions";
import { cmToPt, CM_PER_INCH } from "../../lib/core/units";
import {
  effectiveBackground,
  placeOnSheet,
  resolveSheet,
  type ItemPx,
  type SheetGeometry,
  type SheetPlacement,
  type SheetSetup,
} from "../../lib/core/sheetLayout";
import { getMedia, listDocuments, storeMedia, uniqueTitle } from "../../lib/documentStore";
import type { MediaModel } from "../../lib/media/model";
import { currentImageBytes } from "../../lib/media/download";
import { currentImageSize, hasEdits } from "../../lib/media/editModel";
import { mediaImageDataUri } from "../../lib/media/thumbnail";
import { assembleMediaModel } from "../../lib/media/open";
import { useSheetSetup, type SheetSetupController } from "../../composables/useSheetSetup";

/** The key the one placed image is embedded under. */
export const SHEET_IMAGE_SRC = "sheet-image";

/** The session key the dialog's sheet setup is remembered under. */
export const CONVERT_TO_PDF_SURFACE = "convert-to-pdf";

/** The title a conversion falls back to when its title field is blank. */
const UNTITLED = "Untitled";

/** The browser-bound steps (canvas, image decoding), injectable so the rest
 * of the pipeline is unit-testable without a rendering engine. */
export interface ConvertSteps {
  encode: (src: string, maxWidthPt: number, maxHeightPt: number, maxDpi: number) => Promise<PdfImageData | null>;
  crop: (src: string, crop: SheetPlacement["crop"]) => Promise<string>;
  rasteriseSvg: (bytes: Uint8Array, size: ItemPx) => Promise<ConvertedPng>;
}

/** Typed pass-through for steps arriving through a `.vue` prop (runtime-
 * typed `Object`; espree bars the annotation in `.vue` scripts). */
export function asConvertSteps(value: unknown): ConvertSteps | undefined {
  return value ? (value as ConvertSteps) : undefined;
}

export const BROWSER_CONVERT_STEPS: ConvertSteps = {
  encode: encodeImage,
  crop: cropImage,
  rasteriseSvg: convertSvgToPng,
};

/** The one page the writer receives: sized cmToPt(sheet) × cmToPt(sheet),
 * a background rect over the whole page (white where the output cannot be
 * transparent), and the image at the placed rectangle. Pure. */
export function sheetPdfPage(geometry: SheetGeometry, placement: SheetPlacement, background: string): PdfPageInput {
  const widthPt = cmToPt(geometry.widthCm);
  const heightPt = cmToPt(geometry.heightCm);
  return {
    widthPt,
    heightPt,
    items: [],
    links: [],
    headings: [],
    rects: [{ xPt: 0, yPt: 0, widthPt, heightPt, fill: effectiveBackground(background, "pdf") }],
    images: [
      {
        src: SHEET_IMAGE_SRC,
        xPt: cmToPt(placement.xCm),
        yPt: cmToPt(placement.yCm),
        widthPt: cmToPt(placement.widthCm),
        heightPt: cmToPt(placement.heightCm),
      },
    ],
  };
}

/** An image document's pixel size — of its current image, so an edited
 * image (SPEC-112) places its render: the model's own record, else read
 * from the bytes. Null when neither knows. */
export function sourceItemPx(source: MediaModel, bytes?: Uint8Array): ItemPx | null {
  const known = currentImageSize(source);
  if (known) {
    return known;
  }
  const current = bytes ? { bytes, contentType: source.contentType } : currentImageBytes(source);
  const dimensions = decodeImageDimensions(current.bytes, current.contentType);
  return dimensions ? { widthPx: dimensions.widthPx, heightPx: dimensions.heightPx } : null;
}

/** Converts one image document to PDF bytes. Throws with a user-facing
 * message when it cannot; never touches `source`. */
export async function convertImageToPdf(
  source: MediaModel,
  setup: SheetSetup,
  title: string,
  steps: ConvertSteps = BROWSER_CONVERT_STEPS,
): Promise<Uint8Array> {
  if (source.format === "pdf") {
    throw new Error("That document is already a PDF.");
  }
  // The current image: an edited image converts from its saved render.
  const { bytes, contentType } = currentImageBytes(source);
  const itemPx = sourceItemPx(source, bytes);
  if (!itemPx) {
    throw new Error("The image's size could not be read.");
  }
  const geometry = resolveSheet(setup, itemPx);
  const placement = placeOnSheet(itemPx, geometry.content, setup.fit, setup.align, setup.dpi);
  if (placement.widthCm <= 0 || placement.heightCm <= 0) {
    throw new Error("The image has no visible area on this sheet.");
  }

  let src = `data:${contentType};base64,${bytesToBase64(bytes)}`;
  let crop = placement.crop;
  if (contentType === "image/svg+xml") {
    // Vector in, raster out: the SVG is drawn at the size the sheet places
    // it at (at the chosen dpi), and the crop is scaled onto that raster.
    const size = {
      widthPx: Math.max(1, Math.round((placement.item.widthCm / CM_PER_INCH) * setup.dpi)),
      heightPx: Math.max(1, Math.round((placement.item.heightCm / CM_PER_INCH) * setup.dpi)),
    };
    const png = await steps.rasteriseSvg(bytes, size);
    src = `data:image/png;base64,${bytesToBase64(png.data)}`;
    const sx = png.widthPx / itemPx.widthPx;
    const sy = png.heightPx / itemPx.heightPx;
    crop = { xPx: crop.xPx * sx, yPx: crop.yPx * sy, widthPx: crop.widthPx * sx, heightPx: crop.heightPx * sy };
  }
  if (placement.cropped) {
    src = await steps.crop(src, crop);
  }
  const page = sheetPdfPage(geometry, placement, setup.background);
  const placed = page.images[0];
  const image = await steps.encode(src, placed.widthPt, placed.heightPt, setup.dpi);
  if (!image) {
    throw new Error("The image could not be read.");
  }
  return buildPdf({
    title: title.trim(),
    pages: [page],
    fonts: new Map(),
    images: new Map([[SHEET_IMAGE_SRC, image]]),
  });
}

/** Stores freshly written PDF bytes as a NEW media document titled
 * "<title>.pdf" (numbered past any existing title), its summary parsed from
 * those bytes. Returns the stored model. */
export async function saveConvertedPdf(bytes: Uint8Array, title: string): Promise<MediaModel> {
  const taken = listDocuments().map((entry) => entry.title);
  const model = await assembleMediaModel(uniqueTitle(pdfFileName(title, UNTITLED), taken), "pdf", bytes);
  storeMedia(model);
  return model;
}

/** Hands PDF bytes to the browser as "<title>.pdf". */
export function downloadConvertedPdf(bytes: Uint8Array, title: string): void {
  triggerBlobDownload(bytes, pdfFileName(title, UNTITLED), PDF_MIME_TYPE);
}

export interface ConvertToPdfDeps {
  sourceId: string;
  steps?: ConvertSteps;
}

export interface ConvertToPdfController {
  source: Ref<MediaModel | null>;
  sheet: SheetSetupController;
  title: Ref<string>;
  busy: Ref<boolean>;
  /** The last save/download failure, "" when none. */
  failure: Ref<string>;
  isSvg: ComputedRef<boolean>;
  /** A display URL for the preview's picture. */
  previewUrl: ComputedRef<string>;
  /** The colour the PDF will actually paint behind the image. */
  previewBackground: ComputedRef<string>;
  canConvert: ComputedRef<boolean>;
  /** Why the actions are disabled, "" when they are not. */
  blockedReason: ComputedRef<string>;
  save: () => Promise<MediaModel | null>;
  download: () => Promise<boolean>;
}

export function useConvertToPdf(deps: ConvertToPdfDeps): ConvertToPdfController {
  const steps = deps.steps ?? BROWSER_CONVERT_STEPS;
  const source = shallowRef<MediaModel | null>(getMedia(deps.sourceId) ?? null);
  const itemPx = ref<ItemPx | null>(null);
  const failure = ref("");
  const busy = ref(false);
  const title = ref(source.value?.title ?? "");

  if (source.value) {
    try {
      itemPx.value = sourceItemPx(source.value);
    } catch {
      itemPx.value = null;
    }
  }

  const sheet = useSheetSetup({ surface: CONVERT_TO_PDF_SURFACE, output: "pdf", itemPx });

  const isSvg = computed(() => source.value?.format === "svg" && !hasEdits(source.value));
  // The current image (an edited image's render); a stored media part is
  // already base64, so reading it leaves the part exactly as it is.
  const previewUrl = computed(() => (source.value ? mediaImageDataUri(source.value) : ""));
  const previewBackground = computed(() => effectiveBackground(sheet.setup.value.background, "pdf"));

  const blockedReason = computed(() => {
    if (!source.value) {
      return "That image is no longer in the library.";
    }
    if (!itemPx.value) {
      return "The image's size could not be read.";
    }
    if (!sheet.geometry.value || !sheet.placement.value) {
      return "The sheet is not ready.";
    }
    if (title.value.trim().length === 0) {
      return "Give the PDF a title.";
    }
    return "";
  });
  const canConvert = computed(() => !busy.value && blockedReason.value === "");

  async function run<T>(action: string, then: (bytes: Uint8Array) => Promise<T> | T): Promise<T | null> {
    const model = source.value;
    if (!model || !canConvert.value) {
      return null;
    }
    busy.value = true;
    failure.value = "";
    try {
      const bytes = await convertImageToPdf(model, sheet.setup.value, title.value, steps);
      return await then(bytes);
    } catch (error) {
      const reason = error instanceof Error && error.message ? error.message : String(error);
      failure.value = `Could not ${action}: ${reason}`;
      return null;
    } finally {
      busy.value = false;
    }
  }

  return {
    source,
    sheet,
    title,
    busy,
    failure,
    isSvg,
    previewUrl,
    previewBackground,
    canConvert,
    blockedReason,
    save: () => run("save the PDF", (bytes) => saveConvertedPdf(bytes, title.value)),
    download: async () =>
      (await run("download the PDF", (bytes) => {
        downloadConvertedPdf(bytes, title.value);
        return true;
      })) ?? false,
  };
}
