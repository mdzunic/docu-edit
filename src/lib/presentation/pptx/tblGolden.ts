// The NORMATIVE canonical DrawingML table markup (SPEC-091), hand-authored
// against ECMA-376 §21.1.3 (a:tbl) / §19.3.1.20 (p:graphicFrame): a 2-column,
// 2-row table with a header-row look, a table style id, a solid cell fill and a
// centred cell anchor. pptx/tbl.spec.ts pins the serialiser to it (both ways):
// serializeTableFrame(GOLDEN_TABLE, 5) === TBL_GOLDEN, and the reader decodes
// TBL_GOLDEN back to GOLDEN_TABLE (id aside). Cross-app verification (opening
// the round-tripped deck in PowerPoint/Impress) is manual QA in the PR (AC-16e).
import type { TableFrame } from "../model";

/** The clean model TBL_GOLDEN encodes (its `id` is re-minted on read). */
export const GOLDEN_TABLE: Omit<TableFrame, "id"> = {
  kind: "table",
  xCm: 2,
  yCm: 3,
  widthCm: 8,
  heightCm: 1.8,
  columnWidthsCm: [
    3.5,
    4.5
  ],
  firstRow: true,
  styleId: "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}",
  rows: [
    {
      heightCm: 0.9,
      cells: [
        {
          paragraphs: [
            {
              runs: [
                {
                  text: "Name"
                }
              ]
            }
          ]
        },
        {
          paragraphs: [
            {
              runs: [
                {
                  text: "Value"
                }
              ]
            }
          ]
        }
      ]
    },
    {
      heightCm: 0.9,
      cells: [
        {
          paragraphs: [
            {
              runs: [
                {
                  text: "Alpha"
                }
              ]
            }
          ],
          fill: {
            type: "solid",
            color: "#ffff00"
          },
          verticalAlign: "center"
        },
        {
          paragraphs: [
            {
              runs: [
                {
                  text: "1"
                }
              ]
            }
          ]
        }
      ]
    }
  ]
};

/** The canonical `p:graphicFrame` markup for GOLDEN_TABLE with p:cNvPr id=5. */
export const TBL_GOLDEN =
  "<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id=\"5\" name=\"Table 5\"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp=\"1\"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x=\"720000\" y=\"1080000\"/><a:ext cx=\"2880000\" cy=\"648000\"/></p:xfrm><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/table\"><a:tbl><a:tblPr firstRow=\"1\"><a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId></a:tblPr><a:tblGrid><a:gridCol w=\"1260000\"/><a:gridCol w=\"1620000\"/></a:tblGrid><a:tr h=\"324000\"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:buNone/></a:pPr><a:r><a:rPr lang=\"en-US\" sz=\"1800\"><a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill></a:rPr><a:t>Name</a:t></a:r></a:p></a:txBody><a:tcPr marL=\"46800\" marR=\"46800\" marT=\"46800\" marB=\"46800\"></a:tcPr></a:tc><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:buNone/></a:pPr><a:r><a:rPr lang=\"en-US\" sz=\"1800\"><a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill></a:rPr><a:t>Value</a:t></a:r></a:p></a:txBody><a:tcPr marL=\"46800\" marR=\"46800\" marT=\"46800\" marB=\"46800\"></a:tcPr></a:tc></a:tr><a:tr h=\"324000\"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:buNone/></a:pPr><a:r><a:rPr lang=\"en-US\" sz=\"1800\"><a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill></a:rPr><a:t>Alpha</a:t></a:r></a:p></a:txBody><a:tcPr marL=\"46800\" marR=\"46800\" marT=\"46800\" marB=\"46800\" anchor=\"ctr\"><a:solidFill><a:srgbClr val=\"FFFF00\"/></a:solidFill></a:tcPr></a:tc><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:buNone/></a:pPr><a:r><a:rPr lang=\"en-US\" sz=\"1800\"><a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill></a:rPr><a:t>1</a:t></a:r></a:p></a:txBody><a:tcPr marL=\"46800\" marR=\"46800\" marT=\"46800\" marB=\"46800\"></a:tcPr></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>";
