import { describe, it, expect } from "vitest";
import { parseDocx } from "./docx/reader";
import { writeDocx } from "./docx/writer";
import { parseOdt } from "./odf/reader";
import { writeOdt } from "./odf/writer";
import { firstStructuralDifference } from "./equality";
import {
  createDocument,
  createFieldSpecRun,
  isTof,
  type DocumentModel,
  type Paragraph,
  type TableOfFigures,
  type Run,
  type FieldSpec,
} from "./documentModel";
import { CAPTION_STYLE_ID } from "./styles";

type CaptionRefShow = "full" | "labelAndNumber" | "numberOnly" | "textOnly";

function captionParagraph(label: string, cached: string, text: string, anchor: string): Paragraph {
  return {
    type: "paragraph",
    styleId: CAPTION_STYLE_ID,
    runs: [
      { text: `${label} `, formatting: {} },
      { text: "1", formatting: {}, field: { kind: "seq", label, format: "decimal", cached } },
      { text: ": ", formatting: {} },
      { text, formatting: {} },
    ],
    bookmarks: [{ name: anchor, offset: 0, hidden: true }],
  };
}

function refRun(show: CaptionRefShow, target: string, cached: string): Run {
  return createFieldSpecRun({ kind: "ref", target, show, cached });
}

/** A document with an image, its caption, a cross-reference to it, and a table
 * of figures listing the caption. */
function captionModel(): DocumentModel {
  const model = createDocument("word-captions");
  const anchor = "_cap_Figure_1";
  const tof: TableOfFigures = {
    type: "tof",
    label: "Figure",
    options: { pageNumbers: true, dotLeader: true, hyperlinks: true },
    entries: [{ text: "Figure 1: Revenue", page: 1, targetBookmark: anchor }],
  };
  model.sections[0].blocks = [
    { type: "paragraph", runs: [{ text: "A body paragraph.", formatting: {} }] },
    captionParagraph("Figure", "1", "Revenue", anchor),
    { type: "paragraph", runs: [{ text: "See ", formatting: {} }, refRun("labelAndNumber", anchor, "Figure 1"), { text: ".", formatting: {} }] },
    tof,
  ];
  return model;
}

function captionOf(model: { sections: DocumentModel["sections"] }): Paragraph {
  const block = model.sections[0].blocks.find(
    (b) => b.type === "paragraph" && b.styleId === CAPTION_STYLE_ID,
  );
  return block as Paragraph;
}

function refFieldOf(model: { sections: DocumentModel["sections"] }): FieldSpec | undefined {
  for (const block of model.sections[0].blocks) {
    if (block.type === "paragraph") {
      const run = block.runs.find((r) => r.field?.kind === "ref");
      if (run) {
        return run.field;
      }
    }
  }
  return undefined;
}

describe("caption / seq / table-of-figures round trip (SPEC-055)", () => {
  it("round-trips seq, caption reference and the table of figures through docx (AC-45, AC-46, AC-51, AC-52)", async () => {
    const opened = await parseDocx((await writeDocx(captionModel())).bytes);
    // reopen once more to prove the writer's output re-reads identically.
    const reopened = await parseDocx(
      (await writeDocx({ ...captionModel(), sections: opened.sections })).bytes,
    );
    const diff = firstStructuralDifference(opened.sections, reopened.sections);
    expect(diff).toBeNull();

    // the caption's seq field survived at its saved number
    const caption = captionOf(opened);
    expect(caption.styleId).toBe(CAPTION_STYLE_ID);
    expect(caption.runs.find((r) => r.field?.kind === "seq")?.field).toEqual({
      kind: "seq",
      label: "Figure",
      format: "decimal",
      cached: "1",
    });

    // the caption cross-reference kept its label-and-number format
    expect(refFieldOf(opened)).toMatchObject({ kind: "ref", show: "labelAndNumber" });

    // the table of figures round-tripped with its entry
    const tof = opened.sections[0].blocks.find(isTof) as TableOfFigures;
    expect(tof.label).toBe("Figure");
    expect(tof.entries).toHaveLength(1);
    expect(tof.entries[0].text).toBe("Figure 1: Revenue");
  });

  it("round-trips seq, caption reference and the table of figures through odt (AC-47, AC-48)", async () => {
    const opened = await parseOdt((await writeOdt(captionModel())).bytes);

    const caption = captionOf(opened);
    expect(caption.runs.find((r) => r.field?.kind === "seq")?.field).toEqual({
      kind: "seq",
      label: "Figure",
      format: "decimal",
      cached: "1",
    });

    expect(refFieldOf(opened)).toMatchObject({ kind: "ref", show: "labelAndNumber" });

    const tof = opened.sections[0].blocks.find(isTof) as TableOfFigures;
    expect(tof.label).toBe("Figure");
    expect(tof.entries[0].text).toBe("Figure 1: Revenue");
  });
});
