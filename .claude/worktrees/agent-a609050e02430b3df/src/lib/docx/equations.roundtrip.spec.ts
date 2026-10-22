// SPEC-056 — equations round-trip through .docx as OMML (AC-30, AC-33, AC-49).
import { describe, expect, it } from "vitest";
import { parseDocx } from "./reader";
import { writeDocx } from "./writer";
import { equalStructural, firstStructuralDifference } from "../equality";
import {
  createDocument,
  createEquationRun,
  type DocumentModel,
  type EquationSpec,
  type MathNode,
  type Run,
} from "../documentModel";
import { parseLinear } from "../math/linearParse";

function docWithEquation(body: MathNode[], display: "inline" | "block"): DocumentModel {
  const model = createDocument("eq-doc");
  model.format = "docx";
  const para = model.sections[0].blocks[0];
  if (para.type !== "paragraph") throw new Error("expected paragraph");
  para.runs = [
    { text: "See ", formatting: {} },
    createEquationRun({ display, body }),
    { text: " above.", formatting: {} },
  ];
  return model;
}

async function equationAfterRoundTrip(model: DocumentModel): Promise<EquationSpec> {
  const { bytes, dropped } = await writeDocx(model);
  expect(dropped).toEqual([]); // same-format save drops nothing (AC-33)
  const parsed = await parseDocx(bytes);
  const runs = (parsed.sections[0].blocks[0] as { runs: Run[] }).runs;
  const eqRun = runs.find((r) => r.equation);
  if (!eqRun?.equation) throw new Error("equation not read back");
  return eqRun.equation;
}

describe("equation .docx round trip", () => {
  const cases: Array<[string, string, "inline" | "block"]> = [
    ["inline fraction", "\\frac{a}{b}", "inline"],
    ["display quadratic", "x = \\frac{-b}{2a}", "block"],
    ["scripts and radical", "\\sqrt{x^2 + 1}", "inline"],
    ["n-ary sum", "\\sum_{i=1}^{n} i", "inline"],
    ["matrix in delimiters", "\\left( \\matrix{a & b \\\\ c & d} \\right)", "block"],
    ["accent and function", "\\hat{x} \\sin", "inline"],
  ];

  for (const [label, src, display] of cases) {
    it(`preserves a ${label}`, async () => {
      const body = parseLinear(src);
      const expected: EquationSpec = { display, body };
      const back = await equationAfterRoundTrip(docWithEquation(body, display));
      const diff = firstStructuralDifference(expected, back);
      expect(diff, diff ? JSON.stringify(diff) : "").toBeNull();
      expect(back.display).toBe(display);
    });
  }

  it("preserves an unsupported OMML construct verbatim as opaqueMath (AC-23/25/49)", async () => {
    const opaque: MathNode = {
      kind: "opaqueMath",
      source: `<m:groupChr xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:e><m:r><m:t>y</m:t></m:r></m:e></m:groupChr>`,
      linear: "y",
    };
    const back = await equationAfterRoundTrip(docWithEquation([opaque], "inline"));
    // opaqueMath.source compared with canonicalizeXml (AC-49) via equalStructural.
    expect(equalStructural([opaque], back.body)).toBe(true);
    expect(back.body[0].kind).toBe("opaqueMath");
  });
});
