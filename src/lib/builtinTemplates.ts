// The six shipped templates (SPEC-046 AC-2): plain DocumentModel content —
// no merge chips of their own, so they render and edit exactly like any
// other document — built the same way createDocument() builds a blank one.
// The gallery instantiates one by calling `build()`, which always returns a
// fresh model with a new id.
import { createDocument, type DocumentModel, type Paragraph, type Run, type Table } from "./documentModel";

function run(text: string, bold = false): Run {
  return { text, formatting: bold ? { bold: true } : {} };
}

function paragraph(runs: Run[], headingLevel?: number): Paragraph {
  const p: Paragraph = { type: "paragraph", runs };
  if (headingLevel !== undefined) {
    p.headingLevel = headingLevel;
  }
  return p;
}

function heading(text: string, level = 1): Paragraph {
  return paragraph([run(text, true)], level);
}

function text(value: string): Paragraph {
  return paragraph([run(value)]);
}

function blank(): Paragraph {
  return paragraph([]);
}

function table(header: string[], rows: string[][], widthCm = 4): Table {
  const columnWidthsCm = header.map(() => widthCm);
  return {
    type: "table",
    columnWidthsCm,
    rows: [
      { isHeader: true, cells: header.map((label) => ({ blocks: [paragraph([run(label, true)])] })) },
      ...rows.map((cells) => ({ cells: cells.map((value) => ({ blocks: [text(value)] })) })),
    ],
  };
}

export interface BuiltinTemplate {
  id: string;
  title: string;
  /** A short line shown under the title in the gallery card — not a
   * rendered thumbnail (SPEC-046 leaves the preview form unspecified; a
   * text summary needs no canvas/rendering pipeline and is always
   * accurate). */
  preview: string;
  build: () => DocumentModel;
}

function withBlocks(title: string, blocks: (Paragraph | Table)[]): DocumentModel {
  const model = createDocument(crypto.randomUUID());
  model.title = title;
  model.sections[0].blocks = blocks;
  return model;
}

export const BUILTIN_TEMPLATES: BuiltinTemplate[] = [
  {
    id: "letter",
    title: "Letter",
    preview: "A formal business letter with a date, address block and closing.",
    build: () =>
      withBlocks("Letter", [
        text("[Your Name]"),
        text("[Your Address]"),
        text("[City, State ZIP]"),
        blank(),
        text("[Date]"),
        blank(),
        text("[Recipient Name]"),
        text("[Company Name]"),
        text("[Address]"),
        blank(),
        text("Dear [Recipient Name],"),
        blank(),
        text(
          "I am writing to [state the purpose of your letter]. [Add supporting details in one or two more paragraphs.]",
        ),
        blank(),
        text("Thank you for your time and consideration."),
        blank(),
        text("Sincerely,"),
        blank(),
        text("[Your Name]"),
      ]),
  },
  {
    id: "memo",
    title: "Memo",
    preview: "An internal memo with To/From/Date/Re fields.",
    build: () =>
      withBlocks("Memo", [
        heading("MEMORANDUM", 1),
        blank(),
        paragraph([run("To: ", true), run("[Recipient]")]),
        paragraph([run("From: ", true), run("[Your Name]")]),
        paragraph([run("Date: ", true), run("[Date]")]),
        paragraph([run("Re: ", true), run("[Subject]")]),
        blank(),
        text("[Body of the memo. State the purpose up front, then the details.]"),
      ]),
  },
  {
    id: "report",
    title: "Report",
    preview: "A structured report with summary, sections and conclusion.",
    build: () =>
      withBlocks("Report", [
        heading("[Report Title]", 1),
        text("[Author] · [Date]"),
        blank(),
        heading("Executive Summary", 2),
        text("[A short summary of the report's purpose and key findings.]"),
        blank(),
        heading("Background", 2),
        text("[Context the reader needs before the findings.]"),
        blank(),
        heading("Findings", 2),
        text("[The report's main content.]"),
        blank(),
        heading("Conclusion", 2),
        text("[Summary of conclusions and any recommended next steps.]"),
      ]),
  },
  {
    id: "invoice",
    title: "Invoice",
    preview: "A billing invoice with a line-item table and totals.",
    build: () =>
      withBlocks("Invoice", [
        heading("Invoice", 1),
        paragraph([run("Invoice #: ", true), run("[Invoice Number]")]),
        paragraph([run("Date: ", true), run("[Date]")]),
        blank(),
        paragraph([run("Bill To: ", true), run("[Client Name]")]),
        text("[Client Address]"),
        blank(),
        table(
          ["Description", "Quantity", "Unit Price", "Amount"],
          [
            ["[Item description]", "1", "[0.00]", "[0.00]"],
            ["", "", "", ""],
          ],
        ),
        blank(),
        paragraph([run("Total: ", true), run("[0.00]")]),
        blank(),
        text("Thank you for your business."),
      ]),
  },
  {
    id: "meeting-minutes",
    title: "Meeting minutes",
    preview: "Attendees, agenda, discussion notes and action items.",
    build: () =>
      withBlocks("Meeting minutes", [
        heading("Meeting Minutes", 1),
        paragraph([run("Date: ", true), run("[Date]")]),
        paragraph([run("Attendees: ", true), run("[Names]")]),
        blank(),
        heading("Agenda", 2),
        text("1. [Agenda item]"),
        text("2. [Agenda item]"),
        blank(),
        heading("Discussion", 2),
        text("[Notes from the discussion.]"),
        blank(),
        heading("Action Items", 2),
        table(["Action", "Owner", "Due"], [["[Action item]", "[Owner]", "[Date]"]]),
      ]),
  },
  {
    id: "cv",
    title: "CV",
    preview: "A curriculum vitae with experience, education and skills.",
    build: () =>
      withBlocks("CV", [
        heading("[Your Name]", 1),
        text("[Email] · [Phone] · [Location]"),
        blank(),
        heading("Experience", 2),
        paragraph([run("[Job Title], [Company]", true)]),
        text("[Start] – [End]"),
        text("[One or two lines describing your role and achievements.]"),
        blank(),
        heading("Education", 2),
        paragraph([run("[Degree], [Institution]", true)]),
        text("[Start] – [End]"),
        blank(),
        heading("Skills", 2),
        text("[Skill], [Skill], [Skill]"),
      ]),
  },
];
