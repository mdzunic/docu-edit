import { describe, it, expect } from "vitest";
import {
  buildCorePropertiesXml,
  buildCustomPropertiesXml,
  parseCoreProperties,
  parseCustomProperties,
} from "./docProps";
import { parseXml } from "./xml";
import { parseDocx } from "./reader";
import { writeDocx } from "./writer";
import { unzip } from "./zip";
import { createDocument, type DocumentModel } from "../documentModel";
import { caretSelection, insertText as insertTextCommand } from "../commands";
import { insertComment } from "../comments";
import { applyTrackChanges } from "../trackChangesDiff";

describe("buildCorePropertiesXml / parseCoreProperties round trip (SPEC-041 AC-2, AC-3, AC-6)", () => {
  const modified = new Date("2026-08-03T12:34:56Z");

  it("writes and re-reads every Description-tab field", () => {
    const xml = buildCorePropertiesXml(
      {
        title: "Quarterly Report",
        subject: "Q3 results",
        author: "Ada Lovelace",
        lastModifiedBy: "Grace Hopper",
        keywords: "finance, q3",
        comments: "Draft for review",
        created: "2026-01-01T00:00:00Z",
      },
      modified,
    );
    const properties = parseCoreProperties(xml);
    expect(properties.title).toBe("Quarterly Report");
    expect(properties.subject).toBe("Q3 results");
    expect(properties.author).toBe("Ada Lovelace");
    expect(properties.lastModifiedBy).toBe("Grace Hopper");
    expect(properties.keywords).toBe("finance, q3");
    expect(properties.comments).toBe("Draft for review");
    expect(properties.created).toBe("2026-01-01T00:00:00Z");
    expect(properties.modified).toBe("2026-08-03T12:34:56Z");
  });

  it("omits dcterms:created when unset rather than fabricating one (round-trip fidelity)", () => {
    const xml = buildCorePropertiesXml({ title: "Untitled" }, modified);
    expect(xml).not.toContain("dcterms:created");
    expect(parseCoreProperties(xml).created).toBeUndefined();
  });

  it("escapes XML metacharacters", () => {
    const xml = buildCorePropertiesXml({ title: 'Q3 <Report> & "Friends"' }, modified);
    const doc = parseXml(xml);
    expect(doc.getElementsByTagNameNS("http://purl.org/dc/elements/1.1/", "title")[0].textContent).toBe(
      'Q3 <Report> & "Friends"',
    );
  });

  it("preserves an unmodelled core.xml field as an extra field and re-emits it verbatim (AC-6)", () => {
    const source =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"` +
      ` xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"` +
      ` xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
      `<dc:title>Report</dc:title>` +
      `<cp:category>Finance</cp:category>` +
      `<cp:revision>3</cp:revision>` +
      `</cp:coreProperties>`;
    const properties = parseCoreProperties(source);
    expect(properties.extraFields).toHaveLength(2);
    expect(properties.extraFields?.map((field) => field.name).sort()).toEqual(["category", "revision"]);
    expect(properties.extraFields?.find((field) => field.name === "category")?.value).toBe("Finance");
    expect(properties.extraFields?.find((field) => field.name === "revision")?.value).toBe("3");

    const rebuilt = buildCorePropertiesXml(
      { title: properties.title ?? "", extraFields: properties.extraFields },
      modified,
    );
    const rebuiltProperties = parseCoreProperties(rebuilt);
    expect(rebuiltProperties.extraFields).toHaveLength(2);
    const category = rebuiltProperties.extraFields?.find((field) => field.name === "category");
    expect(category?.xml).toContain("Finance");
    expect(category?.value).toBe("Finance");
    const revision = rebuiltProperties.extraFields?.find((field) => field.name === "revision");
    expect(revision?.xml).toContain("3");
    expect(revision?.value).toBe("3");
  });

  it("tolerates malformed XML by returning no properties", () => {
    expect(parseCoreProperties("not xml")).toEqual({});
  });
});

describe("buildCustomPropertiesXml / parseCustomProperties round trip (SPEC-041 AC-7..10)", () => {
  it("round-trips text, number, date and boolean properties", () => {
    const xml = buildCustomPropertiesXml(
      [
        { name: "Client", type: "text", value: "Acme Corp" },
        { name: "Budget", type: "number", value: "42.5" },
        { name: "Due", type: "date", value: "2026-12-31" },
        { name: "Approved", type: "boolean", value: "true" },
      ],
      [],
    );
    expect(xml).toBeDefined();
    const { customProperties, unsupportedCustomProperties } = parseCustomProperties(xml!);
    expect(unsupportedCustomProperties).toEqual([]);
    expect(customProperties).toEqual([
      { name: "Client", type: "text", value: "Acme Corp" },
      { name: "Budget", type: "number", value: "42.5" },
      { name: "Due", type: "date", value: "2026-12-31" },
      { name: "Approved", type: "boolean", value: "true" },
    ]);
  });

  it("returns undefined when there is nothing to write", () => {
    expect(buildCustomPropertiesXml([], [])).toBeUndefined();
  });

  it("preserves an unsupported property's source element verbatim and assigns fresh properties a non-colliding pid (AC-10)", () => {
    const unsupported = [
      {
        name: "Rate",
        rawType: "cy",
        rawXml:
          '<property xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Rate"><vt:cy>1.5</vt:cy></property>',
      },
    ];
    const xml = buildCustomPropertiesXml([{ name: "Client", type: "text", value: "Acme" }], unsupported);
    expect(xml).toContain("<vt:cy>1.5</vt:cy>"); // preserved verbatim, unparsed
    // pid 2 is taken by the preserved property, so the fresh one gets 3.
    expect(xml).toContain('pid="3" name="Client"');

    const parsed = parseCustomProperties(xml!);
    expect(parsed.customProperties).toEqual([{ name: "Client", type: "text", value: "Acme" }]);
    expect(parsed.unsupportedCustomProperties).toEqual(unsupported);
  });

  it("treats an unrecognised vt: type as unsupported (AC-10)", () => {
    const xml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"` +
      ` xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">` +
      `<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Cost"><vt:cy>100</vt:cy></property>` +
      `</Properties>`;
    const { customProperties, unsupportedCustomProperties } = parseCustomProperties(xml);
    expect(customProperties).toEqual([]);
    expect(unsupportedCustomProperties).toEqual([{ name: "Cost", rawType: "cy", rawXml: expect.stringContaining("vt:cy") }]);
  });
});

describe("writeDocx/parseDocx — document properties integration (SPEC-041 AC-5, AC-8, AC-12..14)", () => {
  function blankModel(): DocumentModel {
    return { ...createDocument("props-doc"), format: "docx" };
  }

  it("writes Description-tab edits to docProps/core.xml and reads them back on reopen (AC-5)", async () => {
    const model = blankModel();
    model.properties = {
      title: "Quarterly Report",
      subject: "Finance",
      author: "Ada Lovelace",
      lastModifiedBy: "Grace Hopper",
      keywords: "q3, finance",
      comments: "Draft",
    };
    const { bytes } = await writeDocx(model);
    const reopened = await parseDocx(bytes);
    expect(reopened.properties.title).toBe("Quarterly Report");
    expect(reopened.properties.subject).toBe("Finance");
    expect(reopened.properties.author).toBe("Ada Lovelace");
    expect(reopened.properties.lastModifiedBy).toBe("Grace Hopper");
    expect(reopened.properties.keywords).toBe("q3, finance");
    expect(reopened.properties.comments).toBe("Draft");
  });

  it("round-trips adding, editing and removing custom properties (AC-8)", async () => {
    const model = blankModel();
    model.properties = {
      customProperties: [
        { name: "Client", type: "text", value: "Acme" },
        { name: "Budget", type: "number", value: "100" },
      ],
    };
    const firstSave = await writeDocx(model);
    const reopened = await parseDocx(firstSave.bytes);
    expect(reopened.properties.customProperties).toEqual(model.properties.customProperties);

    // Edit one, remove the other, add a new one.
    const edited: DocumentModel = {
      ...blankModel(),
      properties: {
        customProperties: [
          { name: "Client", type: "text", value: "Acme Corp" },
          { name: "Approved", type: "boolean", value: "true" },
        ],
      },
    };
    const secondSave = await writeDocx(edited);
    const reopenedAgain = await parseDocx(secondSave.bytes);
    expect(reopenedAgain.properties.customProperties).toEqual([
      { name: "Client", type: "text", value: "Acme Corp" },
      { name: "Approved", type: "boolean", value: "true" },
    ]);

    // Removing every custom property drops the part entirely — no dangling
    // Content-Types Override for a part the save no longer writes.
    const cleared: DocumentModel = { ...blankModel(), properties: {} };
    const thirdSave = await writeDocx(cleared);
    const parts = await unzip(thirdSave.bytes);
    expect(parts.has("docProps/custom.xml")).toBe(false);
    expect(new TextDecoder().decode(parts.get("[Content_Types].xml")!)).not.toContain("custom-properties");
  });

  it("removes a previously-saved custom property from an already-declared custom.xml (same-format save)", async () => {
    const model = blankModel();
    model.properties = { customProperties: [{ name: "Client", type: "text", value: "Acme" }] };
    const firstSave = await writeDocx(model);
    const opened = await parseDocx(firstSave.bytes);
    const reopenedModel: DocumentModel = {
      ...blankModel(),
      opaque: opened.opaque,
      properties: { customProperties: [] },
    };
    const secondSave = await writeDocx(reopenedModel);
    const parts = await unzip(secondSave.bytes);
    expect(parts.has("docProps/custom.xml")).toBe(false);
  });

  it("'Remove personal information on save' strips the author/last-modified-by and anonymizes comment authors in the written bytes (AC-13, AC-14)", async () => {
    let model = blankModel();
    model.sections = [{ blocks: [{ type: "paragraph", runs: [{ text: "Hello world", formatting: {} }] }] }];
    model = insertComment(
      model,
      { anchor: { section: 0, block: 0, offset: 0 }, focus: { section: 0, block: 0, offset: 5 } },
      "Ada Lovelace",
      "note",
    ).model;
    model.properties = {
      author: "Ada Lovelace",
      lastModifiedBy: "Grace Hopper",
      removePersonalInfoOnSave: true,
    };
    const { bytes } = await writeDocx(model);
    const parts = await unzip(bytes);
    const decoder = new TextDecoder();
    expect(decoder.decode(parts.get("docProps/core.xml")!)).not.toContain("Lovelace");
    const commentsXml = decoder.decode(parts.get("word/comments.xml")!);
    expect(commentsXml).not.toContain("Lovelace");
    expect(commentsXml).toContain("Author 1");

    const reopened = await parseDocx(bytes);
    expect(reopened.properties.author).toBeUndefined();
    expect(reopened.properties.lastModifiedBy).toBeUndefined();
    expect(reopened.comments[0].messages[0].author).toBe("Author 1");

    // The live in-memory model is never mutated by a save.
    expect(model.properties.author).toBe("Ada Lovelace");
  });

  it("strips tracked-change author names from the written bytes when the toggle is on", async () => {
    const before = blankModel();
    before.sections = [{ blocks: [{ type: "paragraph", runs: [{ text: "Hello", formatting: {} }] }] }];
    const after = insertTextCommand(before, caretSelection({ section: 0, block: 0, offset: 5 }), " world").model;
    const tracked = applyTrackChanges(before, after, "Ada Lovelace", "2026-01-01T00:00:00.000Z");
    tracked.format = "docx";
    tracked.properties = { removePersonalInfoOnSave: true };
    const { bytes } = await writeDocx(tracked);
    const documentXml = new TextDecoder().decode((await unzip(bytes)).get("word/document.xml")!);
    expect(documentXml).not.toContain("Lovelace");
    expect(documentXml).toContain("Author 1");
  });
});
