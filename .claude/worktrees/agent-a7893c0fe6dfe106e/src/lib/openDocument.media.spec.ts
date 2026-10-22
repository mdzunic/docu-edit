// SPEC-107 on the open funnel: every open path reaches media through
// openAnyDocumentFile, which sniffs the six media formats from the bytes
// BEFORE any unzip attempt and never from the extension; `.md`, `.csv` and
// `.tsv` keep their extension exception (SPEC-126); the text-only entry point
// refuses media with a specific message; one size limit covers every kind.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import {
  DocxOpenError,
  FILE_TOO_LARGE_MESSAGE,
  FileTooLargeError,
  MAX_FILE_SIZE_BYTES,
  MEDIA_NOT_TEXT_MESSAGE,
  openAnyDocumentFile,
  openDocumentFile,
  UnsupportedFileError,
} from "./openDocument";
import { openDocumentFromUrl } from "./openDocumentFromUrl";
import { PdfEncryptedError } from "./core/pdfDocument";
import { equalBytes } from "./core/equality";
import { mediaSourceBytes } from "./media/download";
import type { MediaFormat } from "./media/model";
import {
  buildEncryptedPdfBytes,
  buildGifFixtureBytes,
  buildJpegFixtureBytes,
  buildPngFixtureBytes,
  buildSvgFixtureBytes,
  buildWebpFixtureBytes,
} from "./media/fixtures/build";

const PAGES_PDF = new Uint8Array(readFileSync(join(__dirname, "../../e2e/fixtures/spec107-pages.pdf")));

afterEach(() => {
  vi.unstubAllGlobals();
});

async function fixtureBytes(format: MediaFormat): Promise<Uint8Array> {
  switch (format) {
    case "pdf":
      return PAGES_PDF;
    case "png":
      return buildPngFixtureBytes();
    case "jpeg":
      return buildJpegFixtureBytes();
    case "gif":
      return buildGifFixtureBytes();
    case "svg":
      return buildSvgFixtureBytes();
    case "webp":
      return buildWebpFixtureBytes();
  }
}

function file(bytes: Uint8Array, name: string): File {
  return new File([bytes as BlobPart], name);
}

describe("openAnyDocumentFile — media (SPEC-107)", () => {
  it.each([
    ["pdf", "Report.pdf", "application/pdf"],
    ["png", "Shot.png", "image/png"],
    ["jpeg", "Photo.jpg", "image/jpeg"],
    ["gif", "Anim.gif", "image/gif"],
    ["svg", "Mark.svg", "image/svg+xml"],
    ["webp", "Shot.webp", "image/webp"],
  ] as const)("opens a %s file as a media document, stored verbatim", async (format, name, contentType) => {
    const bytes = await fixtureBytes(format);
    const opened = await openAnyDocumentFile(file(bytes, name));
    expect(opened.kind).toBe("media");
    if (opened.kind !== "media") {
      return;
    }
    expect(opened.model.format).toBe(format);
    expect(opened.model.contentType).toBe(contentType);
    expect(opened.model.title).toBe(name.replace(/\.[a-z]+$/, ""));
    expect(opened.model.byteLength).toBe(bytes.length);
    expect(equalBytes(mediaSourceBytes(opened.model), bytes)).toBe(true);
  });

  it("detects from the bytes, never the extension", async () => {
    const png = await openAnyDocumentFile(file(await buildPngFixtureBytes(), "chart.pdf"));
    expect(png.kind === "media" && png.model.format).toBe("png");
    const pdf = await openAnyDocumentFile(file(PAGES_PDF, "scan.jpg"));
    expect(pdf.kind === "media" && pdf.model.format).toBe("pdf");
    const svg = await openAnyDocumentFile(file(buildSvgFixtureBytes(), "logo"));
    expect(svg.kind === "media" && svg.model.format).toBe("svg");
  });

  it("gives no media extension an exception: a .png that is not a PNG is not opened as one", async () => {
    await expect(openAnyDocumentFile(file(new TextEncoder().encode("hello"), "fake.png"))).rejects.toThrow(DocxOpenError);
    await expect(openAnyDocumentFile(file(new TextEncoder().encode("# Notes"), "notes.svg"))).rejects.toThrow(DocxOpenError);
  });

  it(".md, .csv and .tsv keep their extension exception", async () => {
    const opened = await openAnyDocumentFile(file(new TextEncoder().encode("# Notes\n\nPlain text."), "notes.md"));
    expect(opened.kind).toBe("text");
    expect(opened.kind === "text" && opened.model.format).toBe("md");
    for (const format of ["csv", "tsv"] as const) {
      const table = await openAnyDocumentFile(file(new TextEncoder().encode("a,b\tc\n1,2\t3\n"), `table.${format}`));
      expect(table.kind === "spreadsheet" && table.model.format).toBe(format);
    }
  });

  it("refuses an encrypted PDF with the password message", async () => {
    const error = await openAnyDocumentFile(file(buildEncryptedPdfBytes(), "locked.pdf")).then(
      () => null,
      (thrown: unknown) => thrown,
    );
    expect(error).toBeInstanceOf(PdfEncryptedError);
    expect((error as Error).message).toBe("This PDF is password-protected");
  });

  it("applies the one 50 MB limit, with the one message, to media too", async () => {
    expect(MAX_FILE_SIZE_BYTES).toBe(50 * 1024 * 1024);
    const oversized = new Uint8Array(MAX_FILE_SIZE_BYTES + 1);
    oversized.set(PAGES_PDF);
    const error = await openAnyDocumentFile(file(oversized, "huge.pdf")).then(
      () => null,
      (thrown: unknown) => thrown,
    );
    expect(error).toBeInstanceOf(FileTooLargeError);
    expect((error as Error).message).toBe(FILE_TOO_LARGE_MESSAGE);
  });
});

describe("openDocumentFile's text-only contract (SPEC-107)", () => {
  it("rejects a PDF and an image with a specific message instead of misparsing them", async () => {
    for (const [bytes, name] of [
      [PAGES_PDF, "Report.pdf"],
      [await buildPngFixtureBytes(), "Shot.png"],
    ] as const) {
      const error = await openDocumentFile(file(bytes, name)).then(
        () => null,
        (thrown: unknown) => thrown,
      );
      expect(error).toBeInstanceOf(UnsupportedFileError);
      expect((error as Error).message).toBe(MEDIA_NOT_TEXT_MESSAGE);
    }
  });
});

describe("/open?url= reaches media through the same funnel (SPEC-107)", () => {
  it("opens a fetched PDF served as application/pdf", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn().mockResolvedValue({
        ok: true,
        status: 200,
        statusText: "OK",
        headers: new Headers({ "content-type": "application/pdf" }),
        arrayBuffer: () => Promise.resolve(PAGES_PDF.slice().buffer),
      }),
    );
    const opened = await openDocumentFromUrl("https://example.com/papers/Report.pdf");
    expect(opened.kind).toBe("media");
    expect(opened.kind === "media" && opened.model.pdf?.pageCount).toBe(3);
    expect(opened.model.title).toBe("Report");
  });

  it("opens a fetched WebP served as image/webp", async () => {
    const webp = buildWebpFixtureBytes();
    vi.stubGlobal(
      "fetch",
      vi.fn().mockResolvedValue({
        ok: true,
        status: 200,
        statusText: "OK",
        headers: new Headers({ "content-type": "image/webp" }),
        arrayBuffer: () => Promise.resolve(webp.slice().buffer),
      }),
    );
    const opened = await openDocumentFromUrl("https://example.com/shot");
    expect(opened.kind === "media" && opened.model.format).toBe("webp");
  });
});
