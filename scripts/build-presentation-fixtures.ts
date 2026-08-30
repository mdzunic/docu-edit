// Regenerates the checked-in presentation e2e fixtures from the same
// DOM-free builders the unit specs use, so e2e assertions and unit
// assertions can never drift apart:
//   npx vite-node scripts/build-presentation-fixtures.ts
import { mkdirSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import {
  buildOdpAnimationsFixtureBytes,
  buildOdpChartsFixtureBytes,
  buildOdpFixtureBytes,
  buildOdpTablesFixtureBytes,
  buildPptxAnimationsFixtureBytes,
  buildPptxChartsFixtureBytes,
  buildPptxFixtureBytes,
  buildPptxTablesFixtureBytes,
  buildSpec089OdpFixtureBytes,
  buildSpec089PptxFixtureBytes,
} from "../src/lib/presentation/fixtures/build";

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "..", "e2e", "fixtures");
mkdirSync(fixturesDir, { recursive: true });

const pptx = await buildPptxFixtureBytes();
writeFileSync(join(fixturesDir, "mini-deck.pptx"), pptx);
const odp = await buildOdpFixtureBytes();
writeFileSync(join(fixturesDir, "mini-deck.odp"), odp);
// SPEC-086 animation fixtures (plain body+shape slide + canonical + exotic).
const animPptx = await buildPptxAnimationsFixtureBytes();
writeFileSync(join(fixturesDir, "spec086-animations.pptx"), animPptx);
const animOdp = await buildOdpAnimationsFixtureBytes();
writeFileSync(join(fixturesDir, "spec086-animations.odp"), animOdp);
// SPEC-089 layout catalogue fixtures.
const layoutsPptx = await buildSpec089PptxFixtureBytes();
writeFileSync(join(fixturesDir, "spec089-layouts.pptx"), layoutsPptx);
const layoutsOdp = await buildSpec089OdpFixtureBytes();
writeFileSync(join(fixturesDir, "spec089-layouts.odp"), layoutsOdp);
// SPEC-091 table fixtures (a modelled 3×3 table + one out-of-subset table).
const tablesPptx = await buildPptxTablesFixtureBytes();
writeFileSync(join(fixturesDir, "spec091-tables.pptx"), tablesPptx);
const tablesOdp = await buildOdpTablesFixtureBytes();
writeFileSync(join(fixturesDir, "spec091-tables.odp"), tablesOdp);
// SPEC-092 chart fixtures (column/line/pie + one out-of-catalogue chart).
const chartsPptx = await buildPptxChartsFixtureBytes();
writeFileSync(join(fixturesDir, "spec092-charts.pptx"), chartsPptx);
const chartsOdp = await buildOdpChartsFixtureBytes();
writeFileSync(join(fixturesDir, "spec092-charts.odp"), chartsOdp);
console.log(
  `wrote mini-deck.pptx (${pptx.byteLength}B), mini-deck.odp (${odp.byteLength}B), ` +
    `spec086-animations.pptx (${animPptx.byteLength}B), spec086-animations.odp (${animOdp.byteLength}B), ` +
    `spec089-layouts.pptx (${layoutsPptx.byteLength}B), spec089-layouts.odp (${layoutsOdp.byteLength}B), ` +
    `spec091-tables.pptx (${tablesPptx.byteLength}B), spec091-tables.odp (${tablesOdp.byteLength}B), ` +
    `spec092-charts.pptx (${chartsPptx.byteLength}B), spec092-charts.odp (${chartsOdp.byteLength}B)`,
);
// The core idle scheduler's keep-alive timer would otherwise hold the
// process open after the work is done (a known trap for Node callers).
process.exit(0);
