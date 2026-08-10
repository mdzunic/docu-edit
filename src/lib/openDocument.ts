// Ties a File chosen from the "Open document" picker to the document model:
// read its bytes, detect the format from the bytes themselves (never the
// extension — an ODF package is recognised by its `mimetype` entry), parse
// them, and produce a DocumentModel ready to store and navigate to. Parsing
// errors (DocxOpenError / LegacyDocError / OdtOpenError / FlatOdfError)
// propagate unchanged — the caller only mutates app state once this
// resolves, so a rejected file leaves the previous state untouched.
import type {
  CommentThread,
  DocumentFormat,
  DocumentModel,
  DocumentProperties,
  ListDefinition,
  MediaPart,
  NamedStyle,
  NoteContent,
  NoteSettings,
  OpaquePart,
  SourceXml,
} from "./documentModel";
import { parseDocx, DocxOpenError, LegacyDocError } from "./docx/reader";
import { detectOdf, parseOdt, FlatOdfError, OdtOpenError, FLAT_ODF_MESSAGE } from "./odf/reader";
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
 * detected from content, never the extension, once a file passes this. */
export const UNSUPPORTED_FILE_TYPE_MESSAGE = "We can only open .docx and .odt files";

export function hasSupportedExtension(fileName: string): boolean {
  return /\.(docx|odt)$/i.test(fileName);
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
  const withoutExtension = fileName.replace(/\.(docx|odt)$/i, "");
  return withoutExtension.length > 0 ? withoutExtension : fileName;
}

function buildModel(
  fileName: string,
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
  },
): DocumentModel {
  const model: DocumentModel = {
    id: crypto.randomUUID(),
    title: titleFromFileName(fileName),
    format,
    sections: parsed.sections,
    defaultStyles: parsed.defaultStyles,
  };
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
 * with its specific message, and everything else takes the .docx path. */
export async function openDocumentFile(file: File): Promise<DocumentModel> {
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
    return buildModel(file.name, "odt", await parseOdt(bytes));
  }
  return buildModel(file.name, "docx", await parseDocx(bytes));
}

/** Reads and parses a `.docx` File into a new DocumentModel (not yet
 * stored or navigated to — the caller does that once this resolves). */
export async function openDocxFile(file: File): Promise<DocumentModel> {
  const buffer = await readFileAsArrayBuffer(file);
  return buildModel(file.name, "docx", await parseDocx(buffer));
}
