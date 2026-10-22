// Browser File → bytes plumbing, shared by the open-document picker
// (lib/openDocument.ts) and the image-insert path (lib/imageInsert.ts).
// Product-neutral: nothing here knows what the bytes are.

/** Reads a picked/dropped/pasted File into an ArrayBuffer, rejecting with
 * the FileReader's own error (or a generic one) when the read fails. */
export function readFileAsArrayBuffer(file: File): Promise<ArrayBuffer> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.addEventListener("load", () => resolve(reader.result as ArrayBuffer));
    reader.addEventListener("error", () => reject(reader.error ?? new Error("Failed to read file")));
    reader.readAsArrayBuffer(file);
  });
}
