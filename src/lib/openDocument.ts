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
  type DocumentFormat,
  type DocumentModel,
  type DocumentProperties,
  type ListDefinition,
  type MediaPart,
  type NamedStyle,
  type NoteContent,
  type NoteSettings,
  type OpaquePart,
  type SourceXml,
} from "./documentModel";
import { parseDocx, DocxOpenError, LegacyDocError } from "./docx/reader";
import { detectOdf, parseOdt, FlatOdfError, OdtOpenError, FLAT_ODF_MESSAGE } from "./odf/reader";
import { unzip, ZipError } from "./docx/zip";
import { parseMd, type ParseMdOptions } from "./md/reader";
import { inferListStyleTypes } from "./styles";

export { DocxOpenError, LegacyDocError, OdtOpenError, FlatOdfError };

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
export const UNSUPPORTED_FILE_TYPE_MESSAGE = "We can only open .docx, .odt and .md files";

export function hasSupportedExtension(fileName: string): boolean {
  return /\.(docx|odt|md|markdown)$/i.test(fileName);
}

/** Whether `bytes` is a ZIP package, regardless of what kind — the one
 * content signature Markdown detection needs (SPEC-050's "Detection"):
 * a `.md`/`.markdown`-named file is only admitted by extension when it is
 * *not* a ZIP, so a same-named ZIP (a mis-saved/renamed .docx, say) still
 * takes the content-sniffed path below and fails there on its own terms
 * rather than being silently misread as Markdown. */
async function isZipPackage(bytes: Uint8Array): Promise<boolean> {
  try {
    await unzip(bytes);
    return true;
  } catch (error) {
    if (error instanceof ZipError) {
      return false;
    }
    throw error;
  }
}

/** SPEC-046: the template gallery's "Import template file…" picker, a
 * separate entry point from the regular "Open document" one above — a
 * `.dotx`/`.ott` is a template, not a document to resume editing. */
export function hasTemplateFileExtension(fileName: string): boolean {
  return /\.(dotx|ott)$/i.test(fileName);
}

function readFileAsArrayBuffer(file: File): Promise<ArrayBuffer> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.addEventListener("load", () => resolve(reader.result as ArrayBuffer));
    reader.addEventListener("error", () => reject(reader.error ?? new Error("Failed to read file")));
    reader.readAsArrayBuffer(file);
  });
}

/** The app-bar title for an opened document: its file name with the
 * extension removed (AC-29). */
export function titleFromFileName(fileName: string): string {
  const withoutExtension = fileName.replace(/\.(docx|dotx|odt|ott|md|markdown)$/i, "");
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

/** Reads and parses a picked File into a new DocumentModel (not yet stored
 * or navigated to — the caller does that once this resolves). The format is
 * detected from the file's contents: a ZIP whose `mimetype` entry declares
 * OpenDocument text opens as .odt, a flat XML ODF file (.fodt) is rejected
 * with its specific message, and anything else that is a ZIP takes the
 * .docx path. Markdown has no magic bytes and no container (SPEC-050's
 * "Detection"), so content sniffing cannot decide it — the deliberate
 * exception: a file that is not a ZIP package is Markdown if its name ends
 * `.md`/`.markdown`, and takes the .docx path (where parseDocx rejects it on
 * its own terms) otherwise. `hasSupportedExtension` has already rejected
 * anything else for the one open path (drag-and-drop) that checks a name at
 * all; the picker/URL paths reach here for any file they were handed. */
export async function openDocumentFile(file: File, options: ParseMdOptions = {}): Promise<DocumentModel> {
  if (file.size > MAX_FILE_SIZE_BYTES) {
    throw new FileTooLargeError(FILE_TOO_LARGE_MESSAGE);
  }
  const buffer = await readFileAsArrayBuffer(file);
  const bytes = new Uint8Array(buffer);
  const odf = await detectOdf(bytes);
  if (odf === "flat") {
    throw new FlatOdfError(FLAT_ODF_MESSAGE);
  }
  if (odf === "package") {
    return assembleDocumentModel(titleFromFileName(file.name), "odt", await parseOdt(bytes));
  }
  if (/\.(md|markdown)$/i.test(file.name) && !(await isZipPackage(bytes))) {
    return assembleDocumentModel(titleFromFileName(file.name), "md", await parseMd(bytes, options));
  }
  return assembleDocumentModel(titleFromFileName(file.name), "docx", await parseDocx(bytes));
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

/** The `.md` counterpart of `readDocxBytes` (SPEC-050, matching SPEC-045
 * AC-17's Node-friendly convention). */
export async function readMdBytes(
  bytes: Uint8Array | ArrayBuffer,
  title: string = DEFAULT_TITLE,
  options: ParseMdOptions = {},
): Promise<DocumentModel> {
  return assembleDocumentModel(title, "md", await parseMd(bytes, options));
}
