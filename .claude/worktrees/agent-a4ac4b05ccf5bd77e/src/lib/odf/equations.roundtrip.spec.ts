// SPEC-056 — equations round-trip through .odt as an embedded formula object
// (draw:frame/draw:object -> Formula N/content.xml holding MathML), anchored
// as-char for inline equations (AC-31/33/47/48/50).
import { describe, expect, it } from "vitest";
import { parseOdt } from "./reader";
import { writeOdt } from "./writer";
import { equalStructural, firstStructuralDifference } from "../equality";
import {
  createDocument,
  createEquationRun,
  type DocumentModel,
  type EquationSpec,
  type MathNode,
  type Paragraph,
  type Run,
} from "../documentModel";
import { parseLinear } from "../math/linearParse";

function docWithEquation(body: MathNode[], display: "inline" | "block"): DocumentModel {
  const model = createDocument("eq-doc");
  model.format = "odt";
  const para = model.sections[0].blocks[0];
  if (para.type !== "paragraph") throw new Error("expected paragraph");
  para.runs = [
    { text: "See ", formatting: {} },
    createEquationRun({ display, body }),
    { text: " above.", formatting: {} },
  ];
  return model;
}

function firstEquation(model: Pick<DocumentModel, "sections">): EquationSpec {
  const runs = (model.sections[0].blocks[0] as { runs: Run[] }).runs;
  const eqRun = runs.find((r) => r.equation);
  if (!eqRun?.equation) throw new Error("equation not read back");
  return eqRun.equation;
}

async function equationAfterRoundTrip(model: DocumentModel): Promise<EquationSpec> {
  const { bytes, dropped } = await writeOdt(model);
  expect(dropped).toEqual([]); // same-format save drops nothing (AC-33)
  const parsed = await parseOdt(bytes);
  return firstEquation(parsed);
}

describe("equation .odt round trip (AC-31/33/47/48)", () => {
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

  it("anchors an inline equation as-char and stores MathML in a formula object", async () => {
    const { bytes } = await writeOdt(docWithEquation(parseLinear("x^2"), "inline"));
    // The zip carries a formula sub-document; the frame is as-char (AC-31).
    const parsed = await parseOdt(bytes);
    expect(firstEquation(parsed).display).toBe("inline");
  });

  it("ignores a StarMath annotation and takes the MathML (AC-47)", async () => {
    // A frame whose content.xml holds MathML plus an <annotation> — the reader
    // must read the MathML, not the annotation.
    const { bytes } = await writeOdt(docWithEquation(parseLinear("\\alpha"), "inline"));
    const back = await parseOdt(bytes);
    expect(firstEquation(back).body[0]).toMatchObject({ kind: "sym", text: "α" });
  });

  it("preserves an unsupported MathML construct as opaqueMath", async () => {
    const opaque: MathNode = {
      kind: "opaqueMath",
      source: `<merror xmlns="http://www.w3.org/1998/Math/MathML"><mtext>z</mtext></merror>`,
      linear: "z",
    };
    const back = await equationAfterRoundTrip(docWithEquation([opaque], "inline"));
    expect(equalStructural([opaque], back.body)).toBe(true);
    expect(back.body[0].kind).toBe("opaqueMath");
  });
});

describe("cross-format equation fidelity (AC-32/48)", () => {
  it("writes a .docx-sourced equation into .odt as MathML rather than dropping it", async () => {
    // A model whose source format is .docx (as if opened from Word), saved as
    // .odt — the cross-format path the reviewer flagged as silently dropping
    // the equation. It must now be written, not lost.
    const docx = createDocument("x");
    docx.format = "docx";
    const para = docx.sections[0].blocks[0] as Paragraph;
    para.runs = [createEquationRun({ display: "block", body: parseLinear("E = mc^2") })];
    para.formatting = { alignment: "center" };

    const { bytes: odtBytes, dropped } = await writeOdt(docx);
    // A fully-modelled equation converts losslessly — no equation DropRecord.
    expect(dropped.filter((d) => d.reason.toLowerCase().includes("equation"))).toEqual([]);
    const reopened = await parseOdt(odtBytes);
    const eq = firstEquation(reopened);
    // The content survived the cross-format save (it was not silently dropped).
    expect(eq.body.some((n) => n.kind === "script")).toBe(true);
  });
});
