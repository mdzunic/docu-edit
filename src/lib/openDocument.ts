// Ties a File chosen from the "Open document" picker to the document model:
// read its bytes, detect the format from the bytes themselves (never the
// extension — an ODF package is recognised by its `mimetype` entry), parse
// them, and produce a DocumentModel ready to store and navigate to. Parsing
// errors (DocxOpenError / LegacyDocError / OdtOpenError / FlatOdfError)
// propagate unchanged — the caller only mutates app state once this
// resolves, so a rejected file leaves the previous state untouched.
import {
  DEFAULT_TITLE,
  type CommentThread,
  type DocTheme,
  type DocumentFormat,
  type DocumentModel,
  type DocumentProperties,
  type NamedStyle,
  type NoteContent,
  type NoteSettings,
  type OpaquePart,
  type SourceXml,
} from "./documentModel";
import { type ListDefinition } from "./core/lists";
import { type MediaPart } from "./core/media";
import { parseDocx, DocxOpenError, LegacyDocError } from "./docx/reader";
import { detectOdf, parseOdt, FlatOdfError, OdtOpenError, FLAT_ODF_MESSAGE, ODT_MIMETYPE, OTT_MIMETYPE } from "./odf/reader";
import { unzip, ZipError } from "./core/zip";
import { detectMedia, detectPackage } from "./core/packageDetect";
import type { MediaModel } from "./media/model";
import { assembleMediaModel } from "./media/open";
import { readFileAsArrayBuffer } from "./core/fileReader";
import { parseMd, type ParseMdOptions } from "./md/reader";
import { inferListStyleTypes } from "./styles";
import { OPENABLE_EXTENSIONS } from "./documentKinds";
import type { PresentationModel } from "./presentation/model";
// Only the presentation product's package signatures load with this module:
// its codec (presentation/open, the pptx/odp readers and the Keynote
// importer) is imported in `openFile` once a package IS a deck. Statically,
// every text document's open path paid for the pptx/odp read-write stack —
// the library page and the editor's preload both pulled it in, inside the
// window SPEC-047's firstPageMs times.
import { presentationFormatOf } from "./presentation/detect";
import {
  isIwaPackage,
  isLegacyKeynote,
  KEYNOTE_LEGACY_MESSAGE,
  KEYNOTE_UNREADABLE_MESSAGE,
} from "./presentation/keynote/detect";
import { KeynoteOpenError } from "./presentation/keynote/errors";

export { DocxOpenError, LegacyDocError, OdtOpenError, FlatOdfError, KeynoteOpenError };

/** A file whose format was recognised but isn't openable where it was
 * handed in — an ODF package of an unsupported kind, or a presentation
 * given to a text-only entry point. */
export class UnsupportedFileError extends Error {}
export const UNSUPPORTED_ODF_MESSAGE = "This OpenDocument file isn't a text document or presentation we can open";
export const PRESENTATION_NOT_TEXT_MESSAGE = "This is a presentation file — open it from the library instead";
export const MEDIA_NOT_TEXT_MESSAGE = "This is a PDF or image file — open it from the library instead";

/** SPEC-009 AC-14: shared across every open path (picker, drag-drop,
 * URL fetch) since they all funnel through `openDocumentFile`. */
export const MAX_FILE_SIZE_BYTES = 50 * 1024 * 1024;
export const FILE_TOO_LARGE_MESSAGE = "That file is too large to open (limit 50 MB)";
export class FileTooLargeError extends Error {}

/** SPEC-009 AC-4: drag-and-drop is the one open path with no OS-level file
 * filter (the picker's `accept` attribute has no drop equivalent), so it
 * checks the name itself before ever reading the file. Format is still
 * detected from content, never the extension, once a file passes this —
 * except Markdown (SPEC-050), which has no content signature at all and is
 * admitted by extension as the deliberate, documented exception (see
 * `openDocumentFile`). */
export const UNSUPPORTED_FILE_TYPE_MESSAGE = "We can only open .docx, .odt, .md, .pptx, .odp, .key, .pdf and image files";

const SUPPORTED_EXTENSION_PATTERN = new RegExp(`\\.(${OPENABLE_EXTENSIONS.join("|")})$`, "i");

export function hasSupportedExtension(fileName: string): boolean {
  return SUPPORTED_EXTENSION_PATTERN.test(fileName);
}

/** SPEC-046: the template gallery's "Import template file…" picker, a
 * separate entry point from the regular "Open document" one above — a
 * `.dotx`/`.ott` is a template, not a document to resume editing. */
export function hasTemplateFileExtension(fileName: string): boolean {
  return /\.(dotx|ott)$/i.test(fileName);
}

/** The app-bar title for an opened document: its file name with the
 * extension removed (AC-29). */
export function titleFromFileName(fileName: string): string {
  const withoutExtension = fileName.replace(
    /\.(docx|dotx|odt|ott|md|markdown|pptx|odp|key|pdf|png|jpg|jpeg|gif|svg|webp)$/i,
    "",
  );
  return withoutExtension.length > 0 ? withoutExtension : fileName;
}

/** Assembles a parsed package into a fresh `DocumentModel` (SPEC-004/005):
 * the shared tail of every open path, byte-based and DOM-free beyond what
 * `parsed` itself already required to produce (parseDocx/parseOdt) — no
 * `File`/browser dependency here, which is what lets `readDocxBytes`/
 * `readOdtBytes` below serve a Node caller directly (see
 * src/lib/generate/read.ts, SPEC-045 AC-17). */
export function assembleDocumentModel(
  title: string,
  format: DocumentFormat,
  parsed: {
    sections: DocumentModel["sections"];
    defaultStyles: DocumentModel["defaultStyles"];
    styles: NamedStyle[];
    opaque: OpaquePart[];
    sourceXml: SourceXml;
    media: MediaPart[];
    listDefinitions?: ListDefinition[];
    notes?: NoteContent[];
    footnoteSettings?: NoteSettings;
    endnoteSettings?: NoteSettings;
    comments?: CommentThread[];
    trackChangesRecording?: boolean;
    properties: DocumentProperties;
    /** A `.md` source's verbatim front matter (SPEC-050); never set by a
     * `.docx`/`.odt` parse. */
    frontMatter?: string;
    /** The document's theme (SPEC-077); `dirty` stays unset on open. */
    theme?: DocTheme;
  },
): DocumentModel {
  const model: DocumentModel = {
    id: crypto.randomUUID(),
    title,
    format,
    sections: parsed.sections,
    defaultStyles: parsed.defaultStyles,
  };
  if (parsed.frontMatter !== undefined) {
    model.frontMatter = parsed.frontMatter;
  }
  if (parsed.theme !== undefined) {
    model.theme = parsed.theme;
  }
  // The document's own style catalogue (SPEC-019): what the picker lists
  // and the writers re-declare on save.
  if (parsed.styles.length > 0) {
    model.styles = parsed.styles;
  }
  // Multilevel list definitions (SPEC-032) the document's lists reference.
  if (parsed.listDefinitions && parsed.listDefinitions.length > 0) {
    model.listDefinitions = parsed.listDefinitions;
  }
  // SPEC-008: the retained source package travels with the document for its
  // lifetime, so a later same-format save can round-trip it.
  if (parsed.opaque.length > 0) {
    model.opaque = parsed.opaque;
  }
  if (Object.keys(parsed.sourceXml).length > 0) {
    model.sourceXml = parsed.sourceXml;
  }
  // Images (SPEC-024): media bytes referenced by the model's Image blocks.
  if (parsed.media.length > 0) {
    model.media = parsed.media;
  }
  // Footnotes/endnotes (SPEC-031) the body's reference marks link to.
  if (parsed.notes && parsed.notes.length > 0) {
    model.notes = parsed.notes;
  }
  if (parsed.footnoteSettings) {
    model.footnoteSettings = parsed.footnoteSettings;
  }
  if (parsed.endnoteSettings) {
    model.endnoteSettings = parsed.endnoteSettings;
  }
  // Comments (SPEC-038) anchored in the body's commentAnchors side-lists.
  if (parsed.comments && parsed.comments.length > 0) {
    model.comments = parsed.comments;
  }
  // SPEC-039: the source's persisted recording flag round-trips unchanged
  // but never turns the toggle on itself (see DocumentModel.trackChangesRecording).
  if (parsed.trackChangesRecording) {
    model.trackChangesRecording = true;
  }
  // Document properties and metadata (SPEC-041): omitted entirely when the
  // source file had none, like every other optional model field here.
  if (Object.keys(parsed.properties).length > 0) {
    model.properties = parsed.properties;
  }
  // SPEC-042: neither .docx nor .odt has a "list style" family of its own —
  // recover the Styles pane's "List" grouping from usage, now that styles,
  // paragraphs and list definitions are all parsed (see the doc comment).
  return inferListStyleTypes(model);
}

/** What the kind-agnostic open path hands back: which product owns the
 * file, plus that product's parsed model. The caller stores and routes by
 * `kind` (see documentKinds.ts). */
export type OpenedFile =
  | { kind: "text"; model: DocumentModel }
  | {
      kind: "presentation";
      model: PresentationModel;
      /** SPEC-087's non-blocking notice for a Keynote import that dropped
       * objects/effects; absent when nothing was dropped (and for native
       * .pptx/.odp opens). Surfaced as a toast by `storeOpenedFile`. */
      importNotice?: string;
    }
  | { kind: "media"; model: MediaModel };

// SPEC-047: the interactive open path (picker, drag-drop, URL, library)
// parses a `.docx` with { deferMedia: true } so the ~1s base64 encode of a
// large document's images does not sit between choosing the file and first
// paint. The model comes back with its images' `data` still ""; each page
// force-encodes its own images as it renders (usePaginationView's
// renderEntryFor), and the first idle persist finalises the rest in a
// chunked background pass (documentStore's persistMedia). Headless callers
// (generate, template import) go through the exported helpers below, which
// never defer.
const text = (model: DocumentModel): OpenedFile => ({ kind: "text", model });

/** The shared open pipeline. Unzips at most once for detection (parsers
 * re-unzip on their own, as they always did), then dispatches on the
 * package's own declarations — see detectPackage:
 * - an ODF package opens as .odt (text mimetype), .odp (presentation
 *   mimetype), or is rejected with the honest unsupported-ODF message;
 * - an OPC package whose main part declares PresentationML opens as .pptx;
 * - anything else that is a ZIP takes the .docx path and fails there on its
 *   own terms;
 * - a non-ZIP is flat ODF (rejected specifically), Markdown by extension
 *   (SPEC-050's deliberate exception — Markdown has no content signature),
 *   or takes the .docx path (whose legacy-.doc detection also lives there).
 * `allowPresentations` is false for the text-only entry point, which must
 * keep returning DocumentModel — a presentation there is an error, not a
 * silent misparse. */
async function openFile(file: File, options: ParseMdOptions, allowPresentations: boolean): Promise<OpenedFile> {
  if (file.size > MAX_FILE_SIZE_BYTES) {
    throw new FileTooLargeError(FILE_TOO_LARGE_MESSAGE);
  }
  const buffer = await readFileAsArrayBuffer(file);
  const bytes = new Uint8Array(buffer);
  const title = titleFromFileName(file.name);

  // SPEC-107: media is sniffed BEFORE any unzip attempt, from the bytes and
  // never the extension — media files are not ZIPs, so this is both cheaper
  // and unambiguous. An encrypted or unparseable PDF is refused here
  // (PdfEncryptedError / PdfParseError propagate as the open error), never
  // guessed at.
  const mediaFormat = detectMedia(bytes);
  if (mediaFormat !== null) {
    if (!allowPresentations) {
      // The text-only entry points (template import, style import) must keep
      // returning DocumentModel — a PDF there is an error, not a misparse.
      throw new UnsupportedFileError(MEDIA_NOT_TEXT_MESSAGE);
    }
    return { kind: "media", model: await assembleMediaModel(title, mediaFormat, bytes) };
  }

  let parts: Map<string, Uint8Array> | null = null;
  try {
    parts = await unzip(bytes);
  } catch (error) {
    if (!(error instanceof ZipError)) {
      throw error;
    }
  }
  // SPEC-087: a `.key` that fails content detection must never fall through
  // to the .docx parser and surface a Word error (AC-11) — the one place the
  // extension steers anything, and only towards a better message.
  const namedDotKey = /\.key$/i.test(file.name);
  if (parts === null) {
    if ((await detectOdf(bytes)) === "flat") {
      throw new FlatOdfError(FLAT_ODF_MESSAGE);
    }
    if (/\.(md|markdown)$/i.test(file.name)) {
      return text(assembleDocumentModel(title, "md", await parseMd(bytes, options)));
    }
    if (namedDotKey) {
      // Not even a ZIP — a corrupt or truncated Keynote file.
      throw new KeynoteOpenError(KEYNOTE_UNREADABLE_MESSAGE);
    }
    return text(assembleDocumentModel(title, "docx", await parseDocx(bytes, undefined, { deferMedia: true })));
  }

  // We already inflated the package once above to classify it; hand those
  // parts to the chosen parser so it doesn't decompress the same bytes a
  // second time (a large `.docx`/media-heavy package otherwise pays the
  // whole inflate twice on the open path — SPEC-047).
  const info = detectPackage(parts);
  const presentationFormat = presentationFormatOf(info);
  if (presentationFormat !== null) {
    if (!allowPresentations) {
      throw new UnsupportedFileError(PRESENTATION_NOT_TEXT_MESSAGE);
    }
    const [{ assemblePresentation }, { parsePptx }, { parseOdp }] = await Promise.all([
      import("./presentation/open"),
      import("./presentation/pptx/reader"),
      import("./presentation/odp/reader"),
    ]);
    const parsed = presentationFormat === "pptx" ? await parsePptx(bytes, parts) : await parseOdp(bytes, parts);
    return { kind: "presentation", model: assemblePresentation(title, presentationFormat, parsed) };
  }
  // SPEC-087: Keynote detection is content-based inside the ZIP funnel,
  // after the native presentation dispatch so `presentationFormatOf` and
  // the docx/odf paths are untouched (AC-5..AC-7). iWork '05–'09 gets its
  // own honest rejection (AC-8) before the IWA check.
  if (isLegacyKeynote(parts)) {
    throw new KeynoteOpenError(KEYNOTE_LEGACY_MESSAGE);
  }
  if (isIwaPackage(parts)) {
    if (!allowPresentations) {
      // Text-only entry points reject a Keynote deck exactly like a native
      // presentation (spec Funnel integration).
      throw new UnsupportedFileError(PRESENTATION_NOT_TEXT_MESSAGE);
    }
    // A one-way import: the deck lives on as a from-scratch "pptx"
    // presentation saved via the existing writers.
    const [{ parseKeynote, keynoteImportNotice }, { assemblePresentation }] = await Promise.all([
      import("./presentation/keynote/reader"),
      import("./presentation/open"),
    ]);
    const parsed = await parseKeynote(parts);
    const model = assemblePresentation(title, "pptx", parsed);
    const importNotice = keynoteImportNotice(parsed.dropped);
    return importNotice === null ? { kind: "presentation", model } : { kind: "presentation", model, importNotice };
  }
  if (info.container === "odf") {
    if (info.mimetype === ODT_MIMETYPE || info.mimetype === OTT_MIMETYPE) {
      return text(assembleDocumentModel(title, "odt", await parseOdt(bytes, parts)));
    }
    throw new UnsupportedFileError(UNSUPPORTED_ODF_MESSAGE);
  }
  if (namedDotKey) {
    // A valid ZIP named .key that is no iWork Archive — still a Keynote
    // problem, not a Word document (AC-11).
    throw new KeynoteOpenError(KEYNOTE_UNREADABLE_MESSAGE);
  }
  return text(assembleDocumentModel(title, "docx", await parseDocx(bytes, parts, { deferMedia: true })));
}

/** Reads and parses a picked File of ANY supported kind (not yet stored or
 * navigated to — the caller does that, routing by `kind`). The generic
 * entry points (library picker, drag-and-drop, /open?url=) call this. */
export async function openAnyDocumentFile(file: File, options: ParseMdOptions = {}): Promise<OpenedFile> {
  return openFile(file, options, true);
}

/** Reads and parses a picked File into a new text DocumentModel — the
 * text-only entry point (template import, style import). A presentation
 * handed here is rejected with a specific message rather than misparsed. */
export async function openDocumentFile(file: File, options: ParseMdOptions = {}): Promise<DocumentModel> {
  const opened = await openFile(file, options, false);
  if (opened.kind !== "text") {
    throw new UnsupportedFileError(PRESENTATION_NOT_TEXT_MESSAGE);
  }
  return opened.model;
}

/** Reads and parses a `.dotx`/`.ott` File as a template (SPEC-046 AC-9..11):
 * same content-based format detection and parse path as `openDocumentFile`
 * (a `.dotx` is byte-for-byte an OPC docx package; a `.ott`'s mimetype entry
 * is now recognised by `detectOdf` alongside plain .odt), so the result is
 * assembled the same way — including the fresh id every open already mints,
 * which is what makes this "a new untitled document from it rather than
 * editing the template" (there is no "the template file" in this app to
 * edit in the first place; every open is already a copy into the library).
 *
 * The one thing this adds on top of `openDocumentFile`: a template's
 * package declares itself as a template (a `.dotx`'s `[Content_Types].xml`
 * Override for the main part, a `.ott`'s `mimetype` entry) — carried
 * forward as a same-format document's preserved opaque parts (SPEC-008),
 * that declaration would resurface on every later save, so a "new document
 * from a template" would quietly keep re-saving as a template too. Dropping
 * the retained package here makes every future save rebuild a fresh,
 * ordinary .docx/.odt from the model instead of reusing the template's. */
export async function openTemplateFile(file: File): Promise<DocumentModel> {
  const model = await openDocumentFile(file);
  delete model.opaque;
  delete model.sourceXml;
  return model;
}

/** Reads and parses a `.docx` File into a new DocumentModel (not yet
 * stored or navigated to — the caller does that once this resolves). */
export async function openDocxFile(file: File): Promise<DocumentModel> {
  const buffer = await readFileAsArrayBuffer(file);
  return assembleDocumentModel(titleFromFileName(file.name), "docx", await parseDocx(buffer));
}

/** Reads and parses `.docx` bytes directly into a new DocumentModel, with
 * no `File`/FileReader involved (SPEC-045 AC-17): the Node/Buffer-friendly
 * counterpart of `openDocxFile`, used by src/lib/generate/read.ts. `title`
 * defaults to DEFAULT_TITLE the same way a blank document does — callers
 * that have a meaningful name (e.g. from their own file path) pass it. */
export async function readDocxBytes(
  bytes: Uint8Array | ArrayBuffer,
  title: string = DEFAULT_TITLE,
): Promise<DocumentModel> {
  return assembleDocumentModel(title, "docx", await parseDocx(bytes));
}

/** The `.odt` counterpart of `readDocxBytes` (SPEC-045 AC-17). */
export async function readOdtBytes(
  bytes: Uint8Array | ArrayBuffer,
  title: string = DEFAULT_TITLE,
): Promise<DocumentModel> {
  return assembleDocumentModel(title, "odt", await parseOdt(bytes));
}
