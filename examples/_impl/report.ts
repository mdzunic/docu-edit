// SPEC-045 AC-21: a multi-page report with a table of contents. Run via
// `node examples/report.js`.
import { writeFile, mkdir } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { DocumentBuilder } from "../../src/lib/generate/builder";

const chapters = [
  {
    title: "Market Overview",
    sections: [
      { heading: "Industry Trends", body: "Demand for automated assembly grew 14% year over year." },
      { heading: "Competitive Landscape", body: "Three new entrants shipped comparable hardware this quarter." },
    ],
  },
  {
    title: "Financial Summary",
    sections: [
      { heading: "Revenue", body: "Revenue reached $4.2M, ahead of the $3.8M forecast." },
      { heading: "Expenses", body: "Operating expenses held flat against the prior quarter." },
    ],
  },
  {
    title: "Outlook",
    sections: [
      { heading: "Next Quarter", body: "Two new product lines are scheduled to ship." },
    ],
  },
];

const builder = new DocumentBuilder({ title: "Quarterly Report" })
  .addHeading("Quarterly Report", 1)
  .addParagraph("Prepared for the board of directors.")
  .addTableOfContents({ levels: [1, 2] });

for (const chapter of chapters) {
  builder.addPageBreak().addHeading(chapter.title, 1);
  for (const section of chapter.sections) {
    builder.addHeading(section.heading, 2).addParagraph(section.body);
  }
}

const outDir = fileURLToPath(new URL("../output/", import.meta.url));
await mkdir(outDir, { recursive: true });
const outPath = fileURLToPath(new URL("../output/report.docx", import.meta.url));
await writeFile(outPath, await builder.toBuffer("docx"));
console.log(`Wrote ${outPath}`);
