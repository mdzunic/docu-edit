// Ties a File chosen from the "Open document" picker to the document model:
// read its bytes, parse them, and produce a DocumentModel ready to store and
// navigate to. Parsing errors (DocxOpenError / LegacyDocError) propagate
// unchanged — the caller only mutates app state once this resolves, so a
// rejected file leaves the previous state untouched (AC-33).
import type { DocumentModel } from "./documentModel";
import { parseDocx, DocxOpenError, LegacyDocError } from "./docx/reader";

export { DocxOpenError, LegacyDocError };

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
  const withoutExtension = fileName.replace(/\.docx$/i, "");
  return withoutExtension.length > 0 ? withoutExtension : fileName;
}

/** Reads and parses a `.docx` File into a new DocumentModel (not yet
 * stored or navigated to — the caller does that once this resolves). */
export async function openDocxFile(file: File): Promise<DocumentModel> {
  const buffer = await readFileAsArrayBuffer(file);
  const parsed = await parseDocx(buffer);
  return {
    id: crypto.randomUUID(),
    title: titleFromFileName(file.name),
    format: "docx",
    sections: parsed.sections,
    defaultStyles: parsed.defaultStyles,
  };
}
