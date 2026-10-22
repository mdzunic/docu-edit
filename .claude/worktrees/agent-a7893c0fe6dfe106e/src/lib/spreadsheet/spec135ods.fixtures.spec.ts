// SPEC-135's LibreOffice fixture (spec135-annotations.ods, built by
// fixtures/build.ts): package-level facts only — the builder is
// deterministic, the committed e2e file is its bytes, the package is whole
// with exactly the listed parts, every XML part parses, and the Calc shapes
// the ods annotation reader is proven on (an annotation first in its cell,
// with and without `office:display`; `text:a` links, one paragraph holding
// two) are really in content.xml where SPEC135_ANNOTATIONS_ODS says. Nothing
// here asserts how the app reads them (ods/annotations.spec.ts does).
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { equalBytes } from "../core/equality";
import { parseXml } from "../core/xml";
import { unzip } from "../core/zip";
import { colToLetters } from "./address";
import { buildSpec135AnnotationsOdsBytes, SPEC135_ANNOTATIONS_ODS } from "./fixtures/build";
import { packageIntegrity } from "./fixtures/packageIntegrity";

const E2E_FIXTURES = join(__dirname, "../../../e2e/fixtures");
const TABLE_NS = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const OFFICE_NS = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT_NS = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const XLINK_NS = "http://www.w3.org/1999/xlink";
const decoder = new TextDecoder();
const committed = (name: string): Uint8Array => new Uint8Array(readFileSync(join(E2E_FIXTURES, name)));
const textOf = (parts: Map<string, Uint8Array>, name: string): string => decoder.decode(parts.get(name));
const children = (parent: Element): Element[] => Array.from(parent.children);

/** A table's cells by A1, walking the repeat counts (DOM, small tables only:
 * the trailing runs are counted, never expanded). */
function cellsByA1(table: Element): Map<string, Element> {
  const out = new Map<string, Element>();
  let row = 0;
  for (const rowElement of children(table).filter((child) => child.localName === "table-row")) {
    let col = 0;
    for (const cell of children(rowElement)) {
      const repeat = Number(cell.getAttributeNS(TABLE_NS, "number-columns-repeated") ?? "1") || 1;
      if (cell.childElementCount > 0 || cell.attributes.length > (repeat > 1 ? 1 : 0)) {
        out.set(`${colToLetters(col)}${row + 1}`, cell);
      }
      col += repeat;
    }
    expect(col).toBe(16_384);
    row += Number(rowElement.getAttributeNS(TABLE_NS, "number-rows-repeated") ?? "1") || 1;
  }
  expect(row).toBe(1_048_576);
  return out;
}

describe("spec135-annotations.ods", () => {
  const O = SPEC135_ANNOTATIONS_ODS;
  const parts = async (): Promise<Map<string, Uint8Array>> => unzip(committed(O.file));
  const tables = async (): Promise<Element[]> => {
    const content = parseXml(textOf(await parts(), "content.xml")).documentElement;
    return Array.from(content.getElementsByTagNameNS(TABLE_NS, "table"));
  };

  it("is deterministic, the committed bytes, whole, with exactly the listed parts, every XML part parsing", async () => {
    const built = await buildSpec135AnnotationsOdsBytes();
    expect(equalBytes(built, await buildSpec135AnnotationsOdsBytes())).toBe(true);
    expect(equalBytes(committed(O.file), built)).toBe(true);
    expect(await packageIntegrity(built)).toEqual([]);
    const all = await parts();
    expect([...all.keys()]).toEqual([...O.parts]);
    for (const [name] of all) {
      if (name.endsWith(".xml") || name.endsWith(".rdf")) {
        expect(() => parseXml(textOf(all, name)), name).not.toThrow();
      }
    }
    expect(textOf(all, "meta.xml")).toContain(`meta:table-count="2" meta:cell-count="11" meta:object-count="0"`);
    expect(textOf(all, "meta.xml")).toContain(`<meta:generator>LibreOffice/24.8`);
    expect(textOf(all, "settings.xml")).toContain(`<config:config-item config:name="ActiveTable" config:type="string">Summary</config:config-item>`);
  });

  it("holds the tables in order, every value cell where the facts say", async () => {
    const [summary, data] = await tables();
    expect([summary, data].map((table) => table!.getAttributeNS(TABLE_NS, "name"))).toEqual([...O.tableNames]);
    for (const [table, values] of [
      [summary!, O.values.Summary],
      [data!, O.values.Data],
    ] as const) {
      const cells = cellsByA1(table);
      expect([...cells.keys()].toSorted()).toEqual(Object.keys(values).toSorted());
      for (const [a1, value] of Object.entries(values)) {
        const cell = cells.get(a1)!;
        if (typeof value === "number") {
          expect(cell.getAttributeNS(OFFICE_NS, "value"), a1).toBe(String(value));
        } else {
          expect(cell.getAttributeNS(OFFICE_NS, "value-type"), a1).toBe("string");
          const text = Array.from(cell.getElementsByTagNameNS(TEXT_NS, "p"))
            .filter((p) => p.parentElement === cell)
            .map((p) => p.textContent)
            .join("\n");
          expect(text, a1).toBe(value);
        }
      }
    }
  });

  it("each annotation is its cell's first child, verbatim — B2 without office:display, B4 shown with two paragraphs", async () => {
    const content = textOf(await parts(), "content.xml");
    const [summary] = await tables();
    const cells = cellsByA1(summary!);
    for (const note of [O.notes.B2, O.notes.B4]) {
      expect(content).toContain(`calcext:value-type="float">${note.xml}<text:p>`);
      const annotation = cells.get(note.a1)!.firstElementChild!;
      expect([annotation.namespaceURI, annotation.localName]).toEqual([OFFICE_NS, "annotation"]);
      expect(annotation.hasAttributeNS(OFFICE_NS, "display")).toBe(note.visible);
      expect(children(annotation).map((child) => child.textContent)).toEqual([note.author, note.createdAt, ...note.paragraphs]);
    }
    expect(O.notes.B4.xml).toContain(`<office:annotation office:display="true" draw:style-name="gr1"`);
    expect(content.match(/<office:annotation[\s>]/g)).toHaveLength(2);
  });

  it("links are text:a with xlink:href then xlink:type — A7 titled, B7 internal, C9 one paragraph with two", async () => {
    const content = textOf(await parts(), "content.xml");
    const [summary] = await tables();
    const cells = cellsByA1(summary!);
    for (const link of [O.links.A7, O.links.B7, O.links.C9]) {
      expect(content).toContain(`<table:table-cell office:value-type="string" calcext:value-type="string">${link.paragraph}</table:table-cell>`);
      const anchors = Array.from(cells.get(link.a1)!.getElementsByTagNameNS(TEXT_NS, "a"));
      expect(anchors[0]!.getAttributeNS(XLINK_NS, "href")).toBe(link.href);
    }
    expect(cells.get("A7")!.getElementsByTagNameNS(TEXT_NS, "a")[0]!.getAttributeNS(OFFICE_NS, "title")).toBe(O.links.A7.title);
    const c9 = Array.from(cells.get("C9")!.getElementsByTagNameNS(TEXT_NS, "a"));
    expect(c9.map((anchor) => [anchor.getAttributeNS(XLINK_NS, "href"), anchor.textContent])).toEqual([
      [O.links.C9.href, "One"],
      [O.links.C9.second.href, O.links.C9.second.text],
    ]);
    expect(cells.get("C9")!.getElementsByTagNameNS(TEXT_NS, "p")).toHaveLength(1);
  });
});
