// The NORMATIVE canonical ODF table markup (SPEC-091), hand-authored against
// OASIS ODF 1.3 §9.1 (table:table): a 2-column, 2-row table with a header row,
// a solid cell fill and a centred cell anchor, plus the automatic styles it
// references. odp/tableXml.spec.ts pins the serialiser both ways:
// serializeOdpTableFrame(GOLDEN_TABLE) === { frame: TABLE_GOLDEN, autoStyles:
// TABLE_GOLDEN_STYLES }, and the reader decodes TABLE_GOLDEN back to
// GOLDEN_TABLE (id aside). Cross-app QA is manual in the PR (AC-16e).
import type { TableFrame } from "../model";

/** The clean model TABLE_GOLDEN encodes (its `id` is re-minted on read). */
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
  rows: [
    {
      heightCm: 0.9,
      isHeader: true,
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

/** The canonical `draw:frame` markup for GOLDEN_TABLE. */
export const TABLE_GOLDEN =
  "<draw:frame svg:x=\"2cm\" svg:y=\"3cm\" svg:width=\"8cm\" svg:height=\"1.8cm\"><table:table><table:table-column table:style-name=\"de-C1\"/><table:table-column table:style-name=\"de-C2\"/><table:table-header-rows><table:table-row table:style-name=\"de-R1\"><table:table-cell><text:p>Name</text:p></table:table-cell><table:table-cell><text:p>Value</text:p></table:table-cell></table:table-row></table:table-header-rows><table:table-row table:style-name=\"de-R1\"><table:table-cell table:style-name=\"de-CE1\"><text:p>Alpha</text:p></table:table-cell><table:table-cell><text:p>1</text:p></table:table-cell></table:table-row></table:table></draw:frame>";

/** The automatic styles TABLE_GOLDEN references. */
export const TABLE_GOLDEN_STYLES =
  "<style:style style:name=\"de-C1\" style:family=\"table-column\"><style:table-column-properties style:column-width=\"3.5cm\"/></style:style><style:style style:name=\"de-C2\" style:family=\"table-column\"><style:table-column-properties style:column-width=\"4.5cm\"/></style:style><style:style style:name=\"de-R1\" style:family=\"table-row\"><style:table-row-properties style:row-height=\"0.9cm\" style:min-row-height=\"0.9cm\"/></style:style><style:style style:name=\"de-CE1\" style:family=\"table-cell\"><style:table-cell-properties fo:background-color=\"#ffff00\" style:vertical-align=\"middle\"/></style:style>";
