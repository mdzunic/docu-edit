// Fetches a document from a remote URL and opens it exactly as a picked
// file would (SPEC-009's `/open?url=` route). Every failure — network
// error, CORS block, a non-OK status, or a response that plainly isn't a
// document — is normalised into a single OpenFromUrlError so the view can
// show one message ("Couldn't open that link") with the underlying reason.
import type { DocumentModel } from "./documentModel";
import {
  FILE_TOO_LARGE_MESSAGE,
  MAX_FILE_SIZE_BYTES,
  openDocumentFile,
} from "./openDocument";

export class OpenFromUrlError extends Error {}

const DOCX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const ODT_CONTENT_TYPE = "application/vnd.oasis.opendocument.text";
// Servers that serve document downloads without a precise content type
// commonly fall back to one of these — treated as "could be a document",
// deferring to the extension check and, ultimately, content sniffing.
const GENERIC_BINARY_CONTENT_TYPES = ["application/zip", "application/octet-stream"];

function hasDocumentExtension(url: string): boolean {
  try {
    return /\.(docx|odt)$/i.test(new URL(url).pathname);
  } catch {
    return false;
  }
}

function isSupportedContentType(contentType: string | null): boolean {
  if (!contentType) {
    return true;
  }
  const type = contentType.split(";")[0]?.trim().toLowerCase() ?? "";
  return type === DOCX_CONTENT_TYPE || type === ODT_CONTENT_TYPE || GENERIC_BINARY_CONTENT_TYPES.includes(type);
}

function fileNameFromContentDisposition(header: string | null): string | undefined {
  if (!header) {
    return undefined;
  }
  const match = /filename\*?=(?:UTF-8''|")?([^";]+)"?/i.exec(header);
  if (!match?.[1]) {
    return undefined;
  }
  try {
    return decodeURIComponent(match[1]);
  } catch {
    return match[1];
  }
}

function fileNameFromUrl(url: string): string {
  try {
    const { pathname } = new URL(url);
    const last = pathname.split("/").filter(Boolean).pop();
    return last ? decodeURIComponent(last) : "document";
  } catch {
    return "document";
  }
}

function deriveFileName(response: Response, url: string): string {
  return fileNameFromContentDisposition(response.headers.get("content-disposition")) ?? fileNameFromUrl(url);
}

/** Fetches, validates and parses the document at `url` (SPEC-009 AC-9..
 * AC-14). Delegates to `openDocumentFile` once the bytes are in hand, so
 * format detection, the size limit and parsing errors all match the
 * file-picker path exactly. */
export async function openDocumentFromUrl(url: string): Promise<DocumentModel> {
  let response: Response;
  try {
    response = await fetch(url);
  } catch (error) {
    const reason = error instanceof Error ? error.message : String(error);
    throw new OpenFromUrlError(reason || "The request failed — check the URL and your connection");
  }

  if (!response.ok) {
    throw new OpenFromUrlError(`The server responded with ${response.status} ${response.statusText}`.trim());
  }

  const contentType = response.headers.get("content-type");
  if (!isSupportedContentType(contentType) && !hasDocumentExtension(url)) {
    throw new OpenFromUrlError(`The link returned ${contentType ?? "an unrecognised file type"}, not a .docx or .odt document`);
  }

  const contentLength = response.headers.get("content-length");
  if (contentLength && Number(contentLength) > MAX_FILE_SIZE_BYTES) {
    throw new OpenFromUrlError(FILE_TOO_LARGE_MESSAGE);
  }

  const buffer = await response.arrayBuffer();
  const file = new File([buffer], deriveFileName(response, url));
  try {
    return await openDocumentFile(file);
  } catch (error) {
    const reason = error instanceof Error ? error.message : String(error);
    throw new OpenFromUrlError(reason);
  }
}
