// SPEC-087 — the Keynote import path through the open funnel: content-based
// detection (both container layouts), the rejection messages and their typed
// error, the assembled model's from-scratch shape, the import notice
// threading, and the widened open surface (.key in the extension list,
// accept string and drop message).
import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import {
  hasSupportedExtension,
  KeynoteOpenError,
  openAnyDocumentFile,
  openDocumentFile,
  PRESENTATION_NOT_TEXT_MESSAGE,
  titleFromFileName,
  UNSUPPORTED_FILE_TYPE_MESSAGE,
} from "./openDocument";
import { OPEN_FILE_ACCEPT, OPENABLE_EXTENSIONS } from "./documentKinds";
import {
  KEYNOTE_LEGACY_MESSAGE,
  KEYNOTE_UNREADABLE_MESSAGE,
  NOT_KEYNOTE_MESSAGE,
} from "./presentation/keynote/reader";
import { buildKeynoteFromArchives, buildSyntheticKeynote, FIXTURE_PNG_BASE64, message, varintField } from "./presentation/keynote/fixtures";
import { zip } from "./core/zipWriter";

const e2eFixturesDir = join(dirname(fileURLToPath(import.meta.url)), "..", "..", "e2e", "fixtures");

function asFile(bytes: Uint8Array, name: string): File {
  return new File([bytes], name);
}

describe("the widened open surface (AC-1..AC-4)", () => {
  it("OPENABLE_EXTENSIONS gains \"key\" (AC-1)", () => {
    expect(OPENABLE_EXTENSIONS).toContain("key");
  });

  it("the pickers' accept string lists .key after the native presentation formats (AC-2)", () => {
    // SPEC-116 adds the six workbook extensions after .key.
    expect(OPEN_FILE_ACCEPT).toBe(
      ".docx,.odt,.md,.markdown,.pptx,.odp,.key,.xlsx,.xlsm,.xltx,.xltm,.ods,.ots,.pdf,.png,.jpg,.jpeg,.gif,.svg,.webp",
    );
  });

  it("the drop error names .key (AC-3, in SPEC-116's current wording)", () => {
    expect(UNSUPPORTED_FILE_TYPE_MESSAGE).toBe(
      "We can only open .docx, .odt, .md, .pptx, .odp, .key, .xlsx, .ods, .pdf and image files",
    );
  });

  it("hasSupportedExtension accepts .key, case-insensitively", () => {
    expect(hasSupportedExtension("deck.key")).toBe(true);
    expect(hasSupportedExtension("DECK.KEY")).toBe(true);
  });

  it("titleFromFileName strips .key (AC-56)", () => {
    expect(titleFromFileName("mini-deck.key")).toBe("mini-deck");
  });
});

describe("content-based detection in the ZIP funnel (AC-5..AC-7)", () => {
  it("a ZIP containing Index/Document.iwa takes the Keynote path regardless of extension (AC-5)", async () => {
    const bytes = await buildSyntheticKeynote();
    const opened = await openAnyDocumentFile(asFile(bytes, "renamed.zip"));
    expect(opened.kind).toBe("presentation");
    if (opened.kind === "presentation") {
      expect(opened.model.format).toBe("pptx");
      expect(opened.model.slides).toHaveLength(2);
    }
  });

  it("a ZIP whose Index.zip holds Document.iwa takes the Keynote path too (AC-6)", async () => {
    const bytes = await buildSyntheticKeynote({ nested: true });
    const opened = await openAnyDocumentFile(asFile(bytes, "mini-deck.key"));
    expect(opened.kind).toBe("presentation");
  });

  it("native .pptx and .odp detection is unaffected (AC-7)", async () => {
    const pptx = await openAnyDocumentFile(asFile(readFileSync(join(e2eFixturesDir, "mini-deck.pptx")), "mini-deck.pptx"));
    expect(pptx.kind).toBe("presentation");
    if (pptx.kind === "presentation") {
      expect(pptx.model.format).toBe("pptx");
      expect(pptx.model.opaque?.length ?? 0).toBeGreaterThan(0);
    }
    const odp = await openAnyDocumentFile(asFile(readFileSync(join(e2eFixturesDir, "mini-deck.odp")), "mini-deck.odp"));
    expect(odp.kind).toBe("presentation");
    if (odp.kind === "presentation") {
      expect(odp.model.format).toBe("odp");
    }
  });

  it("docx detection is unaffected (AC-7)", async () => {
    const docxFixture = join(dirname(fileURLToPath(import.meta.url)), "docx", "fixtures", "word-authored.docx");
    const opened = await openAnyDocumentFile(asFile(readFileSync(docxFixture), "word-authored.docx"));
    expect(opened.kind).toBe("text");
  });
});

describe("rejections (AC-8..AC-12)", () => {
  it("rejects an iWork '05–'09 package with the exact legacy message (AC-8)", async () => {
    const legacy = await zip([{ name: "index.apxl", data: new TextEncoder().encode("<presentation/>") }]);
    const open = openAnyDocumentFile(asFile(legacy, "old-deck.key"));
    await expect(open).rejects.toThrow(KEYNOTE_LEGACY_MESSAGE);
    await expect(openAnyDocumentFile(asFile(legacy, "old-deck.key"))).rejects.toBeInstanceOf(KeynoteOpenError);
  });

  it("rejects a gzipped legacy index the same way (AC-8)", async () => {
    const legacy = await zip([{ name: "index.apxl.gz", data: Uint8Array.from([0x1f, 0x8b, 0x00]) }]);
    await expect(openAnyDocumentFile(asFile(legacy, "old-deck.key"))).rejects.toThrow(KEYNOTE_LEGACY_MESSAGE);
  });

  it("rejects an IWA package whose root archive is not a Keynote presentation (AC-9)", async () => {
    const foreign = await buildKeynoteFromArchives([{ id: 1, type: 10000, payload: message(varintField(1, 1)) }]);
    const open = openAnyDocumentFile(asFile(foreign, "letter.pages"));
    await expect(open).rejects.toThrow(NOT_KEYNOTE_MESSAGE);
  });

  it("rejects an IWA package with an undecodable object graph (AC-10)", async () => {
    const garbled = await zip([{ name: "Index/Document.iwa", data: Uint8Array.from([9, 9, 9, 9]) }]);
    const open = openAnyDocumentFile(asFile(garbled, "deck.key"));
    await expect(open).rejects.toThrow(NOT_KEYNOTE_MESSAGE);
    await expect(openAnyDocumentFile(asFile(garbled, "deck.key"))).rejects.toBeInstanceOf(KeynoteOpenError);
  });

  it("rejects a .key that is not even a ZIP as an unreadable Keynote file, not a Word error (AC-11)", async () => {
    const notZip = new TextEncoder().encode("not an archive at all");
    const open = openAnyDocumentFile(asFile(notZip, "broken.key"));
    await expect(open).rejects.toThrow(KEYNOTE_UNREADABLE_MESSAGE);
    await expect(openAnyDocumentFile(asFile(notZip, "broken.key"))).rejects.toBeInstanceOf(KeynoteOpenError);
  });

  it("rejects a valid ZIP named .key with no iWork content the same way (AC-11)", async () => {
    const plainZip = await zip([{ name: "readme.txt", data: new TextEncoder().encode("hello") }]);
    await expect(openAnyDocumentFile(asFile(plainZip, "odd.key"))).rejects.toThrow(KEYNOTE_UNREADABLE_MESSAGE);
  });

  it("a text-only entry point rejects a Keynote deck like any presentation", async () => {
    const bytes = await buildSyntheticKeynote();
    await expect(openDocumentFile(asFile(bytes, "mini-deck.key"))).rejects.toThrow(PRESENTATION_NOT_TEXT_MESSAGE);
  });
});

describe("the assembled model and the import notice (AC-13..AC-16, AC-37/AC-38)", () => {
  it("yields a from-scratch pptx model titled from the file name (AC-13..AC-16)", async () => {
    const opened = await openAnyDocumentFile(asFile(await buildSyntheticKeynote(), "mini-deck.key"));
    expect(opened.kind).toBe("presentation");
    if (opened.kind !== "presentation") {
      return;
    }
    expect(opened.model.format).toBe("pptx");
    expect(opened.model.title).toBe("mini-deck");
    expect(opened.model.opaque).toBeUndefined();
    for (const slide of opened.model.slides) {
      expect(slide.partName).toBeUndefined();
      expect(slide.sourceXml).toBeUndefined();
    }
  });

  it("threads the import notice when objects were dropped (AC-37)", async () => {
    const bytes = await buildSyntheticKeynote({ withForeignDrawable: true, withTransition: true });
    const opened = await openAnyDocumentFile(asFile(bytes, "mini-deck.key"));
    expect(opened.kind).toBe("presentation");
    if (opened.kind === "presentation") {
      expect(opened.importNotice).toBe("Imported from Keynote. 1 object and 1 effect couldn't be imported.");
    }
  });

  it("threads no notice when nothing was dropped (AC-38)", async () => {
    const opened = await openAnyDocumentFile(asFile(await buildSyntheticKeynote(), "mini-deck.key"));
    expect(opened.kind === "presentation" && opened.importNotice).toBeFalsy();
  });

  it("the imported deck saves as a from-scratch .pptx that re-opens with the same content (AC-55's unit half)", async () => {
    const opened = await openAnyDocumentFile(asFile(await buildSyntheticKeynote(), "mini-deck.key"));
    expect(opened.kind).toBe("presentation");
    if (opened.kind !== "presentation") {
      return;
    }
    const { writePptx } = await import("./presentation/pptx/writer");
    const { readPresentationBytes } = await import("./presentation/open");
    const { bytes } = await writePptx(opened.model);
    const reopened = await readPresentationBytes(bytes, "pptx", "mini-deck");

    expect(reopened.slides).toHaveLength(2);
    const texts = reopened.slides.flatMap((slide) =>
      slide.frames.flatMap((frame) =>
        frame.kind === "text" ? frame.paragraphs.map((p) => p.runs.map((r) => r.text).join("")) : [],
      ),
    );
    expect(texts).toEqual(
      expect.arrayContaining([
        "Keynote fixture deck",
        "Imported by docu-edit",
        "Second slide",
        "First bullet point",
        "Second bullet point",
      ]),
    );
    // The PNG media part survives byte-identical.
    const media = reopened.media?.find((part) => part.contentType === "image/png");
    expect(media?.data).toBe(FIXTURE_PNG_BASE64);
    expect(reopened.slides[1].frames.some((frame) => frame.kind === "image")).toBe(true);
    expect(reopened.slides[1].notes).toEqual(["A speaker note for the importer"]);
  });
});
