import { describe, expect, it } from "vitest";
import { createDocument } from "./documentModel";
import { deleteTemplate, getUserTemplate, instantiateTemplate, listUserTemplates, renameTemplate, saveAsTemplate } from "./templateStore";

describe("templateStore (AC-5..8)", () => {
  it("saveAsTemplate stores an independent copy under a fresh id, titled from the name given", () => {
    const doc = createDocument("source-doc");
    doc.title = "Q3 Report";
    doc.sections[0].blocks = [{ type: "paragraph", runs: [{ text: "Hello", formatting: {} }] }];
    const record = saveAsTemplate(doc, "My Template");
    expect(record.id).not.toBe(doc.id);
    expect(record.title).toBe("My Template");
    expect(record.model.title).toBe("My Template");
    expect(getUserTemplate(record.id)).toEqual(record);
    // independent copy: mutating the source afterwards must not affect it
    doc.sections[0].blocks = [];
    expect(record.model.sections[0].blocks).toHaveLength(1);
  });

  it("saveAsTemplate falls back to the document's own title when the name is blank", () => {
    const doc = createDocument("source-doc-2");
    doc.title = "Untitled Draft";
    const record = saveAsTemplate(doc, "   ");
    expect(record.title).toBe("Untitled Draft");
  });

  it("listUserTemplates lists saved templates, sorted non-increasing by save time", () => {
    const a = saveAsTemplate(createDocument("a"), "Template A");
    const b = saveAsTemplate(createDocument("b"), "Template B");
    const listed = listUserTemplates();
    expect(listed.map((t) => t.id)).toEqual(expect.arrayContaining([a.id, b.id]));
    for (let i = 1; i < listed.length; i++) {
      expect(listed[i].createdAt).toBeLessThanOrEqual(listed[i - 1].createdAt);
    }
  });

  it("renameTemplate trims and rejects an empty title", () => {
    const record = saveAsTemplate(createDocument("rename-src"), "Original");
    expect(renameTemplate(record.id, "  Renamed  ")).toBe("Renamed");
    expect(getUserTemplate(record.id)?.title).toBe("Renamed");
    expect(renameTemplate(record.id, "   ")).toBeNull();
    expect(getUserTemplate(record.id)?.title).toBe("Renamed");
  });

  it("renameTemplate on an unknown id returns null", () => {
    expect(renameTemplate("does-not-exist", "X")).toBeNull();
  });

  it("deleteTemplate removes it from the list", () => {
    const record = saveAsTemplate(createDocument("delete-src"), "To Delete");
    expect(getUserTemplate(record.id)).toBeDefined();
    deleteTemplate(record.id);
    expect(getUserTemplate(record.id)).toBeUndefined();
  });

  it("instantiateTemplate returns a fresh document with the template's content and a new id", () => {
    const doc = createDocument("instantiate-src");
    doc.sections[0].blocks = [{ type: "paragraph", runs: [{ text: "Body text", formatting: {} }] }];
    const record = saveAsTemplate(doc, "Instantiable");
    const instance = instantiateTemplate(record.id);
    expect(instance.id).not.toBe(record.id);
    expect(instance.sections[0].blocks).toEqual(record.model.sections[0].blocks);
    // independent copy: mutating the instance must not affect the stored template
    (instance.sections[0].blocks[0] as { runs: { text: string }[] }).runs[0].text = "Changed";
    const stored = getUserTemplate(record.id);
    expect((stored!.model.sections[0].blocks[0] as { runs: { text: string }[] }).runs[0].text).toBe("Body text");
  });
});
