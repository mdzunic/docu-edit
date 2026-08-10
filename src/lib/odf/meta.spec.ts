import { describe, it, expect } from "vitest";
import { buildMetaXml, parseOdfCustomProperties, parseOdfMeta } from "./meta";
import { parseXml } from "../docx/xml";
import { parseOdt } from "./reader";
import { writeOdt } from "./writer";
import { unzip } from "../docx/zip";
import { createDocument, type DocumentModel } from "../documentModel";
import { caretSelection, insertText as insertTextCommand } from "../commands";
import { insertComment } from "../comments";
import { applyTrackChanges } from "../trackChangesDiff";

describe("buildMetaXml / parseOdfMeta round trip (SPEC-041 AC-2, AC-3, AC-6)", () => {
  const modified = new Date("2026-08-03T12:34:56Z");

  it("writes and re-reads every Description-tab field, including multiple keywords", () => {
    const xml = buildMetaXml(
      {
        title: "Field Report",
        subject: "Site visit",
        author: "Ada Lovelace",
        lastModifiedBy: "Grace Hopper",
        keywords: "site, visit, q3",
        comments: "Draft for review",
        created: "2026-01-01T00:00:00Z",
      },
      modified,
    );
    const properties = parseOdfMeta(xml);
    expect(properties.title).toBe("Field Report");
    expect(properties.subject).toBe("Site visit");
    // author -> meta:initial-creator (the original author); lastModifiedBy
    // -> dc:creator (ODF's own "who last modified this" field) — see
    // meta.ts's header comment on the counter-intuitive naming.
    expect(properties.author).toBe("Ada Lovelace");
    expect(properties.lastModifiedBy).toBe("Grace Hopper");
    expect(properties.keywords).toBe("site, visit, q3");
    expect(properties.comments).toBe("Draft for review");
    expect(properties.created).toBe("2026-01-01T00:00:00Z");
    expect(properties.modified).toBe("2026-08-03T12:34:56Z");
  });

  it("omits meta:creation-date when unset rather than fabricating one (round-trip fidelity)", () => {
    const xml = buildMetaXml({ title: "Untitled" }, modified);
    expect(xml).not.toContain("meta:creation-date");
    expect(parseOdfMeta(xml).created).toBeUndefined();
  });

  it("escapes XML metacharacters", () => {
    const xml = buildMetaXml({ title: 'Q3 <Report> & "Friends"' }, modified);
    const doc = parseXml(xml);
    expect(doc.getElementsByTagNameNS("http://purl.org/dc/elements/1.1/", "title")[0].textContent).toBe(
      'Q3 <Report> & "Friends"',
    );
  });

  it("preserves an unmodelled office:meta field as an extra field (AC-6)", () => {
    const source =
      `<?xml version="1.0" encoding="UTF-8"?>` +
      `<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"` +
      ` xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" office:version="1.2">` +
      `<office:meta><dc:title xmlns:dc="http://purl.org/dc/elements/1.1/">Report</dc:title>` +
      `<meta:generator>LibreOffice/7.6</meta:generator></office:meta>` +
      "</office:document-meta>";
    const properties = parseOdfMeta(source);
    expect(properties.extraFields).toHaveLength(1);
    expect(properties.extraFields?.[0].name).toBe("generator");
    expect(properties.extraFields?.[0].xml).toContain("LibreOffice/7.6");
    expect(properties.extraFields?.[0].value).toBe("LibreOffice/7.6");
  });

  it("tolerates malformed XML by returning no properties", () => {
    expect(parseOdfMeta("not xml")).toEqual({});
  });
});

describe("buildMetaXml custom properties / parseOdfCustomProperties (SPEC-041 AC-7..10)", () => {
  const modified = new Date("2026-08-03T12:34:56Z");

  it("round-trips text, number, date and boolean properties", () => {
    const xml = buildMetaXml(
      {
        title: "Doc",
        customProperties: [
          { name: "Client", type: "text", value: "Acme Corp" },
          { name: "Budget", type: "number", value: "42.5" },
          { name: "Due", type: "date", value: "2026-12-31" },
          { name: "Approved", type: "boolean", value: "true" },
        ],
      },
      modified,
    );
    const { customProperties, unsupportedCustomProperties } = parseOdfCustomProperties(xml);
    expect(unsupportedCustomProperties).toEqual([]);
    expect(customProperties).toEqual([
      { name: "Client", type: "text", value: "Acme Corp" },
      { name: "Budget", type: "number", value: "42.5" },
      { name: "Due", type: "date", value: "2026-12-31" },
      { name: "Approved", type: "boolean", value: "true" },
    ]);
  });

  it("preserves an unsupported meta:user-defined value-type verbatim (AC-10)", () => {
    const xml =
      `<?xml version="1.0" encoding="UTF-8"?>` +
      `<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"` +
      ` xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" office:version="1.2">` +
      `<office:meta><meta:user-defined meta:name="Duration" meta:value-type="time">PT1H</meta:user-defined></office:meta>` +
      "</office:document-meta>";
    const { customProperties, unsupportedCustomProperties } = parseOdfCustomProperties(xml);
    expect(customProperties).toEqual([]);
    expect(unsupportedCustomProperties).toEqual([
      { name: "Duration", rawType: "time", rawXml: expect.stringContaining("PT1H") },
    ]);
  });
});

describe("writeOdt/parseOdt — document properties integration (SPEC-041 AC-5, AC-8, AC-12..14)", () => {
  function blankModel(): DocumentModel {
    return { ...createDocument("props-doc"), format: "odt" };
  }

  it("writes Description-tab edits to meta.xml and reads them back on reopen (AC-5)", async () => {
    const model = blankModel();
    model.properties = {
      title: "Field Report",
      subject: "Site visit",
      author: "Ada Lovelace",
      lastModifiedBy: "Grace Hopper",
      keywords: "site, visit",
      comments: "Draft",
    };
    const { bytes } = await writeOdt(model);
    const reopened = await parseOdt(bytes);
    expect(reopened.properties.title).toBe("Field Report");
    expect(reopened.properties.subject).toBe("Site visit");
    expect(reopened.properties.author).toBe("Ada Lovelace");
    expect(reopened.properties.lastModifiedBy).toBe("Grace Hopper");
    expect(reopened.properties.keywords).toBe("site, visit");
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
    const { bytes } = await writeOdt(model);
    const reopened = await parseOdt(bytes);
    expect(reopened.properties.customProperties).toEqual(model.properties.customProperties);

    const edited: DocumentModel = {
      ...blankModel(),
      properties: {
        customProperties: [
          { name: "Client", type: "text", value: "Acme Corp" },
          { name: "Approved", type: "boolean", value: "true" },
        ],
      },
    };
    const secondSave = await writeOdt(edited);
    const reopenedAgain = await parseOdt(secondSave.bytes);
    expect(reopenedAgain.properties.customProperties).toEqual([
      { name: "Client", type: "text", value: "Acme Corp" },
      { name: "Approved", type: "boolean", value: "true" },
    ]);
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
    const { bytes } = await writeOdt(model);
    const parts = await unzip(bytes);
    const decoder = new TextDecoder();
    expect(decoder.decode(parts.get("meta.xml")!)).not.toContain("Lovelace");
    const contentXml = decoder.decode(parts.get("content.xml")!);
    expect(contentXml).not.toContain("Lovelace");
    expect(contentXml).toContain("Author 1");

    const reopened = await parseOdt(bytes);
    expect(reopened.properties.author).toBeUndefined();
    expect(reopened.properties.lastModifiedBy).toBeUndefined();
    expect(reopened.comments[0].messages[0].author).toBe("Author 1");
    expect(model.properties.author).toBe("Ada Lovelace");
  });

  it("strips tracked-change author names from the written bytes when the toggle is on", async () => {
    const before = blankModel();
    before.sections = [{ blocks: [{ type: "paragraph", runs: [{ text: "Hello", formatting: {} }] }] }];
    const after = insertTextCommand(before, caretSelection({ section: 0, block: 0, offset: 5 }), " world").model;
    const tracked = applyTrackChanges(before, after, "Ada Lovelace", "2026-01-01T00:00:00.000Z");
    tracked.format = "odt";
    tracked.properties = { removePersonalInfoOnSave: true };
    const { bytes } = await writeOdt(tracked);
    const contentXml = new TextDecoder().decode((await unzip(bytes)).get("content.xml")!);
    expect(contentXml).not.toContain("Lovelace");
    expect(contentXml).toContain("Author 1");
  });
});
