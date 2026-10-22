// SPEC-135's xlsx annotations fixture (spec135-annotations.xlsx, built by
// fixtures/build.ts): package-level facts only — the builder is
// deterministic, the committed e2e file is its bytes, the package is whole
// with exactly the listed parts, every XML part (the VML too) parses, and
// the Excel 365 shapes the annotation reader is proven on are really in the
// parts. How the app reads them is xlsx/annotations.spec.ts's.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { equalBytes } from "../core/equality";
import { parseXml } from "../core/xml";
import { unzip } from "../core/zip";
import { buildSpec135AnnotationsXlsxBytes, SPEC135_ANNOTATIONS_XLSX } from "./fixtures/build";
import { packageIntegrity } from "./fixtures/packageIntegrity";

const E2E_FIXTURES = join(__dirname, "../../../e2e/fixtures");
const MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const THREADED_NS = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments";
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const EXCEL_NS = "urn:schemas-microsoft-com:office:excel";
const VML_NS = "urn:schemas-microsoft-com:vml";
const decoder = new TextDecoder();
const committed = (name: string): Uint8Array => new Uint8Array(readFileSync(join(E2E_FIXTURES, name)));
const textOf = (parts: Map<string, Uint8Array>, name: string): string => decoder.decode(parts.get(name));
const xmlOf = (parts: Map<string, Uint8Array>, name: string): Element => parseXml(textOf(parts, name)).documentElement;
const elements = (root: Element, ns: string, local: string): Element[] => Array.from(root.getElementsByTagNameNS(ns, local));

describe("spec135-annotations.xlsx", () => {
  const F = SPEC135_ANNOTATIONS_XLSX;
  const parts = async (): Promise<Map<string, Uint8Array>> => unzip(committed(F.file));

  it("is deterministic, the committed bytes, whole, with exactly the listed parts, every XML part parsing", async () => {
    const built = await buildSpec135AnnotationsXlsxBytes();
    expect(equalBytes(built, await buildSpec135AnnotationsXlsxBytes())).toBe(true);
    expect(equalBytes(committed(F.file), built)).toBe(true);
    expect(await packageIntegrity(built)).toEqual([]);
    const all = await parts();
    expect([...all.keys()]).toEqual([...F.parts]);
    for (const [name] of all) {
      expect(() => parseXml(textOf(all, name)), name).not.toThrow();
    }
  });

  it("relates the comments, VML, threaded and ctrlProps parts from Summary and the persons part from the workbook", async () => {
    const all = await parts();
    const sheetRels = elements(xmlOf(all, F.sheets.Summary.rels), PKG_REL_NS, "Relationship").map((rel) => [
      rel.getAttribute("Id"),
      rel.getAttribute("Target"),
      rel.getAttribute("TargetMode"),
    ]);
    expect(sheetRels).toEqual(
      expect.arrayContaining([
        [F.vml.rId, "../drawings/vmlDrawing1.vml", null],
        [F.comments.rId, "../comments1.xml", null],
        [F.threadedComments.rId, "../threadedComments/threadedComment1.xml", null],
        [F.checkbox.rId, "../ctrlProps/ctrlProp1.xml", null],
        ...F.hyperlinks.flatMap((link) => ("rId" in link ? [[link.rId, link.target, "External"]] : [])),
      ]),
    );
    const workbookRels = elements(xmlOf(all, "xl/_rels/workbook.xml.rels"), PKG_REL_NS, "Relationship");
    expect(workbookRels.some((rel) => rel.getAttribute("Type") === F.person.type && rel.getAttribute("Target") === "persons/person.xml")).toBe(true);
    const sheet = xmlOf(all, F.sheets.Summary.partName);
    expect(elements(sheet, MAIN_NS, "legacyDrawing")[0]?.getAttributeNS(REL_NS, "id")).toBe(F.vml.rId);
    const hyperlinks = elements(sheet, MAIN_NS, "hyperlink").map((link) => ({
      ref: link.getAttribute("ref"),
      rId: link.getAttributeNS(REL_NS, "id"),
      location: link.getAttribute("location"),
      tooltip: link.getAttribute("tooltip"),
    }));
    expect(hyperlinks).toEqual(
      F.hyperlinks.map((link) => ({
        ref: link.ref,
        rId: "rId" in link ? link.rId : null,
        location: "location" in link ? link.location : null,
        tooltip: "tooltip" in link ? link.tooltip : null,
      })),
    );
  });

  it("holds the authors, the rich note, both placeholders and the threads as Excel 365 writes them", async () => {
    const all = await parts();
    const comments = xmlOf(all, F.comments.partName);
    expect(elements(comments, MAIN_NS, "author").map((author) => author.textContent)).toEqual([...F.authors]);
    const list = elements(comments, MAIN_NS, "comment");
    expect(list.map((comment) => comment.getAttribute("ref"))).toEqual(["B2", "D2", "E2", "B4"]);
    const b2 = list[0]!;
    const runs = elements(b2, MAIN_NS, "r");
    expect(runs.map((run) => run.textContent)).toEqual(F.notes.B2.runs.map((run) => run.text));
    expect(elements(runs[0]!, MAIN_NS, "b")).toHaveLength(1);
    for (const ref of ["D2", "E2"] as const) {
      const placeholder = list.find((comment) => comment.getAttribute("ref") === ref)!;
      const thread = F.threads[ref];
      expect(F.authors[Number(placeholder.getAttribute("authorId"))]).toBe(`tc=${thread.threadId}`);
      expect(placeholder.textContent).toBe(
        F.placeholder.preamble + `Comment:\n    ${thread.messages[0].text}` + thread.messages.slice(1).map((reply) => `\nReply:\n    ${reply.text}`).join(""),
      );
    }
    const threaded = elements(xmlOf(all, F.threadedComments.partName), THREADED_NS, "threadedComment");
    expect(threaded.map((message) => [message.getAttribute("ref"), message.getAttribute("id"), message.getAttribute("parentId"), message.getAttribute("done")])).toEqual([
      ...F.threads.D2.messages.map((message, index) => ["D2", message.id, index === 0 ? null : F.threads.D2.threadId, null]),
      ...F.threads.E2.messages.map((message, index) => ["E2", message.id, index === 0 ? null : F.threads.E2.threadId, index === 0 ? "1" : null]),
    ]);
    const persons = elements(xmlOf(all, F.person.partName), THREADED_NS, "person");
    expect(persons.map((person) => [person.getAttribute("id"), person.getAttribute("displayName")])).toEqual(F.persons.map((person) => [person.personId, person.displayName]));
  });

  it("holds four note shapes (B4 shown, 150 × 75 pt) and a checkbox shape with its own shapetype in one VML part", async () => {
    const all = await parts();
    const vml = textOf(all, F.vml.partName);
    expect(vml.startsWith("<xml ")).toBe(true);
    expect(vml.includes("\r\n")).toBe(true);
    const root = parseXml(vml).documentElement;
    const shapes = elements(root, VML_NS, "shape");
    const data = shapes.map((shape) => elements(shape, EXCEL_NS, "ClientData")[0]!);
    expect(data.map((item) => item.getAttribute("ObjectType"))).toEqual(["Note", "Note", "Note", "Note", "Checkbox"]);
    expect(data.slice(0, 4).map((item) => [elements(item, EXCEL_NS, "Row")[0]!.textContent, elements(item, EXCEL_NS, "Column")[0]!.textContent])).toEqual([
      ["1", "1"],
      ["3", "1"],
      ["1", "3"],
      ["1", "4"],
    ]);
    expect(data.map((item) => elements(item, EXCEL_NS, "Visible").length)).toEqual([0, 1, 0, 0, 0]);
    expect(shapes[1]!.getAttribute("style")).toContain(`width:${F.notes.B4.sizePt.width}pt;height:${F.notes.B4.sizePt.height}pt`);
    expect(elements(root, VML_NS, "shapetype").map((type) => type.getAttribute("id"))).toEqual(["_x0000_t202", "_x0000_t201"]);
    expect(shapes[4]!.getAttributeNS("urn:schemas-microsoft-com:office:office", "spid")).toBe(F.checkbox.shapeId);
  });
});
