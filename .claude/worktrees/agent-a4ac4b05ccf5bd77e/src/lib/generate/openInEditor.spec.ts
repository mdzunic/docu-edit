// SPEC-045 AC-16: a generated Blob can be opened directly in the in-browser
// editor (SPEC-002) with no download round trip. This drives the exact
// functions HomeView.vue's own file-picker handler calls after a user picks
// a file — openDocumentFile then storeDocument — but starting from an
// in-memory Blob this API just produced, never a URL.createObjectURL +
// anchor-download + re-selected-File cycle. Runs under vitest's jsdom
// environment, which is what supplies Blob/File/FileReader here — the
// same DOM surface a real browser gives the app.
import { describe, expect, it } from "vitest";
import { DocumentBuilder } from "./builder";
import { openDocumentFile } from "../openDocument";
import { getDocument, storeDocument } from "../documentStore";
import { equalStructural } from "../equality";
import { DOCX_MIME_TYPE } from "../docx/writer";

describe("generated Blob -> in-browser editor (AC-16)", () => {
  it("opens with openDocumentFile/storeDocument, no download round trip", async () => {
    const builder = new DocumentBuilder({ title: "Generated Report" });
    builder.addHeading("Quarterly Report", 1).addParagraph("Prepared by the finance team.");

    // The only thing a real browser flow has that this doesn't: the Blob
    // came from a <a download> + re-picked File instead of straight from
    // toBlob(). Wrapping it in a File is required by openDocumentFile's
    // signature, not a disk trip — nothing here touches storage or a URL.
    const blob = await builder.toBlob("docx");
    expect(blob.type).toBe(DOCX_MIME_TYPE);
    const file = new File([blob], "Generated Report.docx", { type: DOCX_MIME_TYPE });

    const opened = await openDocumentFile(file);
    storeDocument(opened);

    const stored = getDocument(opened.id);
    expect(stored).toBeDefined();
    expect(stored?.title).toBe("Generated Report");
    expect(equalStructural(stored?.sections, builder.toDoc().sections)).toBe(true);
  });
});
