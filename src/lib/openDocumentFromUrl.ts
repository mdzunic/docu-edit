// Fetches a document from a remote URL and opens it exactly as a picked
// file would (SPEC-009's `/open?url=` route). Every failure — network
// error, CORS block, a non-OK status, or a response that plainly isn't a
// document — is normalised into a single OpenFromUrlError so the view can
// show one message ("Couldn't open that link") with the underlying reason.
import {
  FILE_TOO_LARGE_MESSAGE,
  MAX_FILE_SIZE_BYTES,
  openAnyDocumentFile,
  type OpenedFile,
} from "./openDocument";
import { OPENABLE_EXTENSIONS } from "./documentKinds";
import { LIBRARY_IMAGE_MIME_TYPES } from "./core/media";

export class OpenFromUrlError extends Error {}

const DOCX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const ODT_CONTENT_TYPE = "application/vnd.oasis.opendocument.text";
const PPTX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const ODP_CONTENT_TYPE = "application/vnd.oasis.opendocument.presentation";
// SPEC-050: Markdown's registered content type. Plain "text/plain" is
// deliberately not included here — unlike the extension check below (which
// only ever matches an actual .md/.markdown URL), that would accept any
// plain-text link at all and defeat the "not a document" error entirely.
const MARKDOWN_CONTENT_TYPE = "text/markdown";
// Servers that serve document downloads without a precise content type
// commonly fall back to one of these — treated as "could be a document",
// deferring to the extension check and, ultimately, content sniffing.
const GENERIC_BINARY_CONTENT_TYPES = ["application/zip", "application/octet-stream"];

// Derived from the registry (documentKinds.ts), so this path gains new
// kinds' extensions without local knowledge of them (SPEC-107 AC-1).
const OPENABLE_EXTENSION_PATTERN = new RegExp(`\\.(${OPENABLE_EXTENSIONS.join("|")})$`, "i");

function hasDocumentExtension(url: string): boolean {
  try {
    return OPENABLE_EXTENSION_PATTERN.test(new URL(url).pathname);
  } catch {
    return false;
  }
}

function isSupportedContentType(contentType: string | null): boolean {
  if (!contentType) {
    return true;
  }
  const type = contentType.split(";")[0]?.trim().toLowerCase() ?? "";
  return (
    type === DOCX_CONTENT_TYPE ||
    type === ODT_CONTENT_TYPE ||
    type === MARKDOWN_CONTENT_TYPE ||
    type === PPTX_CONTENT_TYPE ||
    type === ODP_CONTENT_TYPE ||
    type === "application/pdf" ||
    (LIBRARY_IMAGE_MIME_TYPES as readonly string[]).includes(type) ||
    GENERIC_BINARY_CONTENT_TYPES.includes(type)
  );
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
 * AC-14). Delegates to `openAnyDocumentFile` once the bytes are in hand, so
 * format detection, the size limit and parsing errors all match the
 * file-picker path exactly — presentations included; the caller routes by
 * the returned `kind`. */
export async function openDocumentFromUrl(url: string): Promise<OpenedFile> {
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
    throw new OpenFromUrlError(
      `The link returned ${contentType ?? "an unrecognised file type"}, not a .docx, .odt, .md, .pptx, .odp, .pdf or image file`,
    );
  }

  const contentLength = response.headers.get("content-length");
  if (contentLength && Number(contentLength) > MAX_FILE_SIZE_BYTES) {
    throw new OpenFromUrlError(FILE_TOO_LARGE_MESSAGE);
  }

  const buffer = await response.arrayBuffer();
  const file = new File([buffer], deriveFileName(response, url));
  try {
    // SPEC-050: a relative image `src` in a Markdown file opened this way
    // resolves against the document's own URL — the one open path that has
    // one (a picked/dropped file has no base to resolve a relative path
    // against at all).
    return await openAnyDocumentFile(file, { baseUrl: url });
  } catch (error) {
    const reason = error instanceof Error ? error.message : String(error);
    throw new OpenFromUrlError(reason);
  }
}
