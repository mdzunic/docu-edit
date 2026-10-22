// BUG-1 regression: a control character in run text used to be written raw
// into `word/document.xml`/`content.xml`, producing a package that is not
// well-formed XML — Word refuses it, and this app's own reader refuses it
// with "This file isn't a Word document we can open", blaming the file for
// what the save did to it. The user's document was destroyed at save time.
//
// This asserts the whole loop the user actually experiences (type/paste →
// save → reopen), not just the sanitizer in isolation (core/xmlText.spec.ts),
// because the bug lived in the seam between them.
import { describe, expect, it } from "vitest";
import { createDocument, type Paragraph } from "./documentModel";
import { caretSelection } from "./commands";
import { pastePlainText } from "./clipboard";
import { insertText } from "./textEditing";
import { writeDocx } from "./docx/writer";
import { writeOdt } from "./odf/writer";
import { parseDocx } from "./docx/reader";
import { parseOdt } from "./odf/reader";
import { unzip } from "./core/zip";

const ch = (code: number) => String.fromCharCode(code);
const SOH = ch(0x01);
const VT = ch(0x0b);
const LONE_SURROGATE = ch(0xd800);

function bodyText(blocks: readonly unknown[]): string {
  return (blocks as Paragraph[])
    .filter((b) => b.type === "paragraph")
    .flatMap((b) => b.runs.map((r) => r.text))
    .join("");
}

describe("BUG-1 — control characters never reach a saved package", () => {
  const dirty = `before${SOH}mid${VT}end${LONE_SURROGATE}tail`;

  it("insertText strips them at the one insertion choke point", () => {
    const model = createDocument("doc");
    const caret = { section: 0, block: 0, offset: 0 };
    insertText(model, caret, dirty);
    expect(bodyText(model.sections[0].blocks)).toBe("beforemid\nendtail");
  });

  it("a paste carrying Word's U+000B line break becomes a real line break", () => {
    // The realistic case: copying two lines out of Word puts U+000B between
    // them in the plain-text flavour.
    const model = createDocument("doc");
    const at = { section: 0, block: 0, offset: 0 };
    const { model: next } = pastePlainText(model, caretSelection(at), `line one${VT}line two`);
    expect(bodyText(next.sections[0].blocks)).toBe("line one\nline two");
  });

  it("saves a .docx whose document.xml is well-formed, and reopens it", async () => {
    const model = createDocument("doc");
    insertText(model, { section: 0, block: 0, offset: 0 }, dirty);

    const { bytes } = await writeDocx(model);
    const parts = await unzip(bytes);
    const xml = new TextDecoder().decode(parts.get("word/document.xml")!);
    expect(xml.includes(SOH)).toBe(false);
    expect(xml.includes(VT)).toBe(false);

    // The assertion that failed before the fix: DocxOpenError, on the file
    // the app had just written.
    const reopened = await parseDocx(bytes);
    expect(bodyText(reopened.sections[0].blocks)).toBe("beforemid\nendtail");
  });

  it("saves an .odt whose content.xml is well-formed, and reopens it", async () => {
    const model = createDocument("doc");
    insertText(model, { section: 0, block: 0, offset: 0 }, dirty);

    const { bytes } = await writeOdt(model);
    const parts = await unzip(bytes);
    const xml = new TextDecoder().decode(parts.get("content.xml")!);
    expect(xml.includes(SOH)).toBe(false);
    expect(xml.includes(VT)).toBe(false);

    const reopened = await parseOdt(bytes);
    expect(bodyText(reopened.sections[0].blocks)).toBe("beforemid\nendtail");
  });

  // Where the invariant's boundary sits, stated rather than asserted: it is
  // held at ingress, so code that hand-builds a Run literal bypasses it.
  // Nothing in the app does — every insertion path routes through insertText,
  // and every XML-backed reader is protected by DOMParser rejecting the
  // source outright. Deliberately not pinned as a test: an
  // added write-side backstop would be an improvement, and a test asserting
  // today's gap would stand in its way.

  it("a .md source is sanitized too — the one reader whose input isn't XML", async () => {
    const { parseMd } = await import("./md/reader");
    const source = new TextEncoder().encode(`para ${SOH}one${VT}still one\n`);
    const parsed = await parseMd(source);

    // Sanitized at the decode, so the VT becomes a newline in the *Markdown
    // source* — and a lone newline inside a paragraph is a soft wrap in
    // Markdown, which this reader joins with a space (see joinBlockLines).
    // That is the format's own rule applying to a character that should
    // never have been in the file, and it keeps the words apart rather than
    // running them together; what matters is that nothing illegal survives.
    expect(bodyText(parsed.sections[0].blocks)).toBe("para one still one");

    const model = createDocument("from md");
    model.sections = parsed.sections;
    const { bytes } = await writeDocx(model);
    await expect(parseDocx(bytes)).resolves.toBeTruthy();
  });
});
