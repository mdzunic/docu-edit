import { describe, it, expect, afterEach, vi } from "vitest";
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { openDocumentFromUrl, OpenFromUrlError } from "./openDocumentFromUrl";
import { FILE_TOO_LARGE_MESSAGE, MAX_FILE_SIZE_BYTES } from "./openDocument";

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "docx", "fixtures");
const docxBytes = readFileSync(join(fixturesDir, "word-authored.docx"));
const DOCX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

function mockFetchResolvedWith(response: Partial<Response>) {
  vi.stubGlobal("fetch", vi.fn().mockResolvedValue(response));
}

afterEach(() => {
  vi.unstubAllGlobals();
});

describe("openDocumentFromUrl", () => {
  it("fetches and opens a .docx file, exactly like the file picker (AC-9)", async () => {
    mockFetchResolvedWith({
      ok: true,
      status: 200,
      statusText: "OK",
      headers: new Headers({ "content-type": DOCX_CONTENT_TYPE }),
      arrayBuffer: () => Promise.resolve(docxBytes.buffer.slice(docxBytes.byteOffset, docxBytes.byteOffset + docxBytes.byteLength)),
    });

    const doc = await openDocumentFromUrl("https://example.com/reports/Quarterly.docx");
    expect(doc.format).toBe("docx");
    expect(doc.title).toBe("Quarterly");
  });

  it("names the document from a Content-Disposition header when present", async () => {
    mockFetchResolvedWith({
      ok: true,
      status: 200,
      statusText: "OK",
      headers: new Headers({
        "content-type": DOCX_CONTENT_TYPE,
        "content-disposition": 'attachment; filename="Team Notes.docx"',
      }),
      arrayBuffer: () => Promise.resolve(docxBytes.buffer.slice(docxBytes.byteOffset, docxBytes.byteOffset + docxBytes.byteLength)),
    });

    const doc = await openDocumentFromUrl("https://example.com/download?id=42");
    expect(doc.title).toBe("Team Notes");
  });

  it("wraps a network failure or CORS block in OpenFromUrlError with the reason (AC-11, AC-12)", async () => {
    vi.stubGlobal("fetch", vi.fn().mockRejectedValue(new TypeError("Failed to fetch")));
    await expect(openDocumentFromUrl("https://blocked.example.com/doc.docx")).rejects.toThrow(OpenFromUrlError);
    await expect(openDocumentFromUrl("https://blocked.example.com/doc.docx")).rejects.toThrow("Failed to fetch");
  });

  it("wraps a non-OK response status in OpenFromUrlError (AC-11)", async () => {
    mockFetchResolvedWith({ ok: false, status: 404, statusText: "Not Found", headers: new Headers() });
    await expect(openDocumentFromUrl("https://example.com/missing.docx")).rejects.toThrow(OpenFromUrlError);
    await expect(openDocumentFromUrl("https://example.com/missing.docx")).rejects.toThrow(/404/);
  });

  it("rejects a non-document content type when the URL has no supported extension either (AC-13)", async () => {
    mockFetchResolvedWith({
      ok: true,
      status: 200,
      statusText: "OK",
      headers: new Headers({ "content-type": "text/html" }),
    });
    await expect(openDocumentFromUrl("https://example.com/page")).rejects.toThrow(OpenFromUrlError);
    await expect(openDocumentFromUrl("https://example.com/page")).rejects.toThrow(/text\/html/);
  });

  it("falls back to the URL extension when the content type is generic", async () => {
    mockFetchResolvedWith({
      ok: true,
      status: 200,
      statusText: "OK",
      headers: new Headers({ "content-type": "application/octet-stream" }),
      arrayBuffer: () => Promise.resolve(docxBytes.buffer.slice(docxBytes.byteOffset, docxBytes.byteOffset + docxBytes.byteLength)),
    });
    const doc = await openDocumentFromUrl("https://example.com/files/report.docx");
    expect(doc.format).toBe("docx");
  });

  it("refuses a response advertising more than 50 MB via Content-Length, before downloading it (AC-14)", async () => {
    const arrayBuffer = vi.fn();
    mockFetchResolvedWith({
      ok: true,
      status: 200,
      statusText: "OK",
      headers: new Headers({
        "content-type": DOCX_CONTENT_TYPE,
        "content-length": String(MAX_FILE_SIZE_BYTES + 1),
      }),
      arrayBuffer,
    });
    await expect(openDocumentFromUrl("https://example.com/huge.docx")).rejects.toThrow(FILE_TOO_LARGE_MESSAGE);
    expect(arrayBuffer).not.toHaveBeenCalled();
  });

  it("still rejects an oversized body when Content-Length is absent (AC-14)", async () => {
    mockFetchResolvedWith({
      ok: true,
      status: 200,
      statusText: "OK",
      headers: new Headers({ "content-type": DOCX_CONTENT_TYPE }),
      arrayBuffer: () => Promise.resolve(new ArrayBuffer(MAX_FILE_SIZE_BYTES + 1)),
    });
    await expect(openDocumentFromUrl("https://example.com/huge.docx")).rejects.toThrow(FILE_TOO_LARGE_MESSAGE);
  });
});
