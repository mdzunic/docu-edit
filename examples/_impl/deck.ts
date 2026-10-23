// SPEC-106: a data-driven quarterly review deck — a title slide, a
// highlights slide with speaker notes, one slide per region built from the
// numbers, a scorecard laid out as text boxes on a blank slide, and a
// closing slide. Run via `node examples/deck.js`; writes
// examples/output/quarterly-review.pptx, .odp and .html.
import { mkdir, writeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { PresentationBuilder } from "../../src/lib/generate/presentation";

const quarter = {
  name: "Q3 2026",
  prepared: "Prepared 30 September 2026",
  revenueGrowthPct: 12,
  regions: [
    { name: "Europe", revenueM: 4.2, growthPct: 20, wins: ["Two new retail partners", "Nordics launch"] },
    { name: "Americas", revenueM: 6.1, growthPct: 9, wins: ["Renewed the largest account"] },
    { name: "APAC", revenueM: 2.3, growthPct: 0, wins: [] },
  ],
  scorecard: [
    { label: "Revenue", value: "$12.6M", good: true },
    { label: "Gross margin", value: "61%", good: true },
    { label: "Churn", value: "3.4%", good: false },
  ],
};

const growth = (pct: number): string => (pct === 0 ? "flat" : `${pct > 0 ? "up" : "down"} ${Math.abs(pct)}%`);
const best = quarter.regions.reduce((a, b) => (b.growthPct > a.growthPct ? b : a));

const deck = new PresentationBuilder({ title: `${quarter.name} review` })
  .addTitleSlide([{ text: `${quarter.name} review`, bold: true }], quarter.prepared)
  .addBulletSlide("Highlights", [
    `Revenue ${growth(quarter.revenueGrowthPct)}`,
    quarter.regions.map((region) => `${region.name} ${growth(region.growthPct)}`),
  ])
  .setNotes([`Lead with ${best.name}.`, "Keep APAC short."]);

for (const region of quarter.regions) {
  deck.addBulletSlide(region.name, [
    { runs: ["Revenue ", { text: `$${region.revenueM.toFixed(1)}M`, bold: true }] },
    [`${growth(region.growthPct)} on Q2`],
    ...(region.wins.length > 0 ? ["Wins", region.wins] : []),
  ]);
}

// A scorecard: one labelled box per metric, three across, coloured by result.
deck.addBlankSlide().setBackground("#f4f6f8").addTextBox([{ text: "Scorecard", fontSizePt: 36, bold: true }], {
  x: { cm: 2 },
  y: { cm: 1.5 },
  width: { cm: 29 },
  height: { cm: 2.5 },
});
quarter.scorecard.forEach((metric, i) => {
  const x = { cm: 2 + i * 10 };
  deck
    .addTextBox([{ text: metric.label, fontSizePt: 20 }], { x, y: { cm: 6 }, width: { cm: 9 }, height: { cm: 1.5 } })
    .addTextBox([{ text: metric.value, fontSizePt: 44, bold: true, color: metric.good ? "#1e7a3c" : "#b42318" }], {
      x,
      y: { cm: 8 },
      width: { cm: 9 },
      height: { cm: 3 },
    });
});

deck.addTitleSlide("Questions?", `Next review: ${quarter.name.replace("Q3", "Q4")}`);

const outDir = fileURLToPath(new URL("../output/", import.meta.url));
await mkdir(outDir, { recursive: true });
for (const format of ["pptx", "odp"] as const) {
  const outPath = `${outDir}quarterly-review.${format}`;
  await writeFile(outPath, await deck.toBuffer(format));
  console.log(`Wrote ${outPath}`);
}
const htmlPath = `${outDir}quarterly-review.html`;
await writeFile(htmlPath, deck.toHtml({ includeNotes: true }));
console.log(`Wrote ${htmlPath}`);
