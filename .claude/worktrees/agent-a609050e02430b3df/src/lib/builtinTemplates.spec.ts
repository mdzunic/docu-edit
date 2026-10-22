import { describe, expect, it } from "vitest";
import { BUILTIN_TEMPLATES } from "./builtinTemplates";
import { writeDocx } from "./docx/writer";
import { writeOdt } from "./odf/writer";
import { parseDocx } from "./docx/reader";
import { parseOdt } from "./odf/reader";

describe("BUILTIN_TEMPLATES (AC-2, AC-3, AC-4)", () => {
  it("ships at least six templates covering letter, memo, report, invoice, meeting minutes, CV", () => {
    expect(BUILTIN_TEMPLATES.length).toBeGreaterThanOrEqual(6);
    const ids = BUILTIN_TEMPLATES.map((t) => t.id);
    expect(ids).toEqual(expect.arrayContaining(["letter", "memo", "report", "invoice", "meeting-minutes", "cv"]));
  });

  it("every template has a unique id and non-empty title/preview", () => {
    const ids = BUILTIN_TEMPLATES.map((t) => t.id);
    expect(new Set(ids).size).toBe(ids.length);
    for (const template of BUILTIN_TEMPLATES) {
      expect(template.title.length).toBeGreaterThan(0);
      expect(template.preview.length).toBeGreaterThan(0);
    }
  });

  it("build() returns a fresh id each time", () => {
    const template = BUILTIN_TEMPLATES[0];
    const a = template.build();
    const b = template.build();
    expect(a.id).not.toBe(b.id);
  });

  it.each(BUILTIN_TEMPLATES.map((t) => [t.id, t] as const))("%s renders correctly and round-trips through .docx", async (_id, template) => {
    const model = template.build();
    expect(model.sections[0].blocks.length).toBeGreaterThan(0);
    const { bytes, dropped } = await writeDocx(model);
    expect(bytes.length).toBeGreaterThan(0);
    expect(dropped).toEqual([]);
    const reopened = await parseDocx(bytes);
    expect(reopened.sections[0].blocks.length).toBe(model.sections[0].blocks.length);
  });

  it.each(BUILTIN_TEMPLATES.map((t) => [t.id, t] as const))("%s renders correctly and round-trips through .odt", async (_id, template) => {
    const model = template.build();
    const { bytes, dropped } = await writeOdt(model);
    expect(bytes.length).toBeGreaterThan(0);
    expect(dropped).toEqual([]);
    const reopened = await parseOdt(bytes);
    expect(reopened.sections[0].blocks.length).toBe(model.sections[0].blocks.length);
  });
});
