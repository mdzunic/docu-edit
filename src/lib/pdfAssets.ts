// What a built PDF embeds, loaded for it: the bundled faces its text uses
// and the images its pages place (SPEC-034). Shared by both products'
// exporters — the word processor's usePdfExport and the deck's useDeckExport
// (SPEC-104) — so a deck's PDF loads and embeds fonts exactly as a
// document's does. App layer, like lib/pdfRaster.ts: it fetches the faces
// from the app's own base URL.
import type { PdfImageData, PdfPageInput } from "./core/pdfBuild";
import { faceFileName, faceFontDescriptor, type FaceKey } from "./core/pdfFonts";
import { encodeImage } from "./core/pdfImages";
import { parseSfnt, type SfntFont } from "./core/sfnt";

/** A face load either yields the parsed program or the reason it could not
 * be used. Carrying the reason (rather than a bare null) is what lets a
 * failed export say *why* — an unreachable /fonts/ path, an HTTP status and
 * a corrupt font file are three very different things to chase. */
type FaceResult = { font: SfntFont } | { error: string };

/** Parsed font programs cached for the session — one fetch+parse per face
 * across any number of exports. Failures are NOT cached, so a transient
 * network problem does not poison later attempts. These promises never
 * reject: a failure resolves to its reason. */
const faceCache = new Map<FaceKey, Promise<FaceResult>>();

/** The bundled faces live in public/fonts/, so they are served under the
 * app's base path — not necessarily the origin root. */
function faceUrl(face: FaceKey): string {
  return `${import.meta.env.BASE_URL}fonts/${faceFileName(face)}`;
}

function loadFace(face: FaceKey): Promise<FaceResult> {
  let pending = faceCache.get(face);
  if (!pending) {
    const url = faceUrl(face);
    pending = fetch(url)
      .then(async (response) => {
        if (!response.ok) {
          throw new Error(`HTTP ${response.status}`);
        }
        return { font: parseSfnt(new Uint8Array(await response.arrayBuffer())) };
      })
      .catch((cause: unknown) => {
        faceCache.delete(face);
        const reason = cause instanceof Error ? cause.message : String(cause);
        return { error: `${url}: ${reason}` };
      });
    faceCache.set(face, pending);
  }
  return pending;
}

/** Decodes every image the pages place, once per distinct source. An
 * image that cannot be decoded is left out rather than failing the
 * export — the rest of the document still becomes a PDF. Each source is
 * sampled for the largest size it is printed at. */
export async function loadPdfImages(pages: PdfPageInput[]): Promise<Map<string, PdfImageData>> {
  const largest = new Map<string, { widthPt: number; heightPt: number }>();
  for (const page of pages) {
    for (const item of page.images) {
      const seen = largest.get(item.src);
      largest.set(item.src, {
        widthPt: Math.max(seen?.widthPt ?? 0, item.widthPt),
        heightPt: Math.max(seen?.heightPt ?? 0, item.heightPt),
      });
    }
  }
  const images = new Map<string, PdfImageData>();
  const encoded = await Promise.all(
    Array.from(largest.entries()).map(
      async ([src, size]) =>
        [src, await encodeImage(src, size.widthPt, size.heightPt)] as const,
    ),
  );
  for (const [src, image] of encoded) {
    if (image) {
      images.set(src, image);
    }
  }
  return images;
}

/** Loads and parses every bundled face the pages' text uses. Only a total
 * loss is fatal — a PDF missing one face still embeds the rest. */
export async function loadPdfFonts(pages: PdfPageInput[]): Promise<Map<FaceKey, SfntFont>> {
  const faces = new Set<FaceKey>();
  for (const page of pages) {
    for (const item of page.items) {
      faces.add(item.face);
    }
  }
  const fonts = new Map<FaceKey, SfntFont>();
  const failures: string[] = [];
  const loaded = await Promise.all(
    Array.from(faces).map(async (face) => [face, await loadFace(face)] as const),
  );
  for (const [face, result] of loaded) {
    if ("font" in result) {
      fonts.set(face, result.font);
    } else {
      failures.push(result.error);
    }
  }
  // Only a total loss is fatal — a PDF missing one face still embeds the
  // rest. The reason rides along so the toast names what actually broke.
  if (faces.size > 0 && fonts.size === 0) {
    throw new Error(`no fonts could be loaded for embedding (${failures[0]})`);
  }
  return fonts;
}

/** Faces already registered with `document.fonts` for raster export. */
const rasterFaces = new Set<FaceKey>();

/** Registers each loaded face under the CSS family `faceCssFont` names (the
 * raster export's `fillText` font), from the very bytes the PDF embeds — so a
 * PNG and a PDF of a page draw the same glyphs even where the app never
 * registered the bundled faces itself. Once per face per session. */
export async function registerRasterFaces(fonts: Map<FaceKey, SfntFont>): Promise<void> {
  if (typeof FontFace === "undefined" || typeof document === "undefined" || !document.fonts) {
    return;
  }
  const loads: Promise<unknown>[] = [];
  for (const [face, font] of fonts) {
    if (rasterFaces.has(face)) {
      continue;
    }
    rasterFaces.add(face);
    const { family, weight, style } = faceFontDescriptor(face);
    const fontFace = new FontFace(family, font.bytes.slice().buffer, { weight, style });
    document.fonts.add(fontFace);
    loads.push(fontFace.load().catch(() => undefined));
  }
  await Promise.all(loads);
}
