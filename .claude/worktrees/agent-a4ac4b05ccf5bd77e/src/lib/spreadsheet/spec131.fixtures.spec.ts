// SPEC-131's drawing fixtures (spec131-drawings.xlsx and .ods, built by
// fixtures/build.ts): each committed e2e file is the builder's bytes, the
// package is whole, and the producer shapes the readers and writers are
// proven on are really in the parts — Excel's three MCE wrappers (a14 form
// control with its empty Fallback, sle15 slicer with its text-box Fallback,
// aink ink with its picture Fallback), every anchor kind, the group, the
// hidden shape, the EMF picture, the a14/a16 extLsts on cNvPr and the
// worksheet's mc:Ignorable; LibreOffice's cell-anchored frame with
// table:end-cell-address, a cell-anchored shape without it, table:shapes, a
// draw:transform rotation, a draw:g, a draw:control over office:forms and
// the gr1… graphic styles.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { equalBytes } from "../core/equality";
import { unzip } from "../core/zip";
import { buildSpec131DrawingsOdsBytes, buildSpec131DrawingsXlsxBytes, SPEC131_DRAWINGS_ODS, SPEC131_DRAWINGS_XLSX, SPEC131_PICTURE_PNG } from "./fixtures/build";
import { packageIntegrity } from "./fixtures/packageIntegrity";

const E2E_FIXTURES = join(__dirname, "../../../e2e/fixtures");
const decoder = new TextDecoder();
const committed = (name: string): Uint8Array => new Uint8Array(readFileSync(join(E2E_FIXTURES, name)));

describe("spec131-drawings.xlsx", () => {
  const F = SPEC131_DRAWINGS_XLSX;

  it("is the builder's bytes, whole, with Excel's parts in order", async () => {
    const bytes = committed("spec131-drawings.xlsx");
    expect(equalBytes(bytes, await buildSpec131DrawingsXlsxBytes())).toBe(true);
    expect(await packageIntegrity(bytes)).toEqual([]);
    expect([...(await unzip(bytes)).keys()]).toEqual([...F.parts]);
  });

  it("holds Excel's drawing shapes: every anchor kind, the three MCE wrappers, the group, the hidden shape, the EMF picture, the cNvPr extLsts", async () => {
    const parts = await unzip(committed("spec131-drawings.xlsx"));
    const drawing = decoder.decode(parts.get(F.drawingPart));
    const sheet = decoder.decode(parts.get("xl/worksheets/sheet1.xml"));
    expect(sheet).toContain(`mc:Ignorable="x14ac xr xr2 xr3"`);
    expect(sheet).toMatch(/<pageMargins[^>]*\/><drawing r:id="rId1"\/><\/worksheet>$/);
    expect(drawing.match(/<xdr:twoCellAnchor editAs="oneCell">/g)).toHaveLength(6); // picture, slicer ×2, ink ×2, EMF
    expect(drawing).toContain(`<xdr:twoCellAnchor editAs="absolute">`);
    expect(drawing.match(/<xdr:oneCellAnchor>/g)).toHaveLength(1);
    expect(drawing.match(/<xdr:absoluteAnchor>/g)).toHaveLength(1);
    expect(drawing).toContain(`<a:prstGeom prst="roundRect">`);
    expect(drawing).toContain(`<a:t>Assumptions</a:t>`);
    expect(drawing).toContain(`<xdr:cxnSp macro="">`);
    expect(drawing).toContain(`<a:prstGeom prst="straightConnector1">`);
    expect(drawing.match(/<xdr:grpSp>/g)).toHaveLength(1);
    expect(drawing).toMatch(/<xdr:cNvPr id="10" name="Rectangle 9" hidden="1">/);
    expect(drawing).toContain(`Requires="a14"><xdr:twoCellAnchor editAs="absolute">`);
    expect(drawing).toContain(`<a14:compatExt spid="_x0000_s1025"/>`);
    expect(drawing).toContain(`</mc:Choice><mc:Fallback/></mc:AlternateContent>`);
    expect(drawing).toContain(`Requires="sle15">`);
    expect(drawing).toContain(`<a:t>${F.slicerText}</a:t>`);
    expect(drawing).toContain(`Requires="aink">`);
    expect(drawing).toContain(`<xdr:contentPart xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId3"/>`);
    expect(drawing.match(/a16:creationId/g)!.length).toBeGreaterThanOrEqual(12);
    const rels = decoder.decode(parts.get(F.drawingRels));
    expect(rels).toContain(`Target="../media/image2.emf"`);
    // The picture and the ink's Fallback share one media part.
    expect(rels.match(/Target="\.\.\/media\/image1\.png"/g)).toHaveLength(2);
    const types = decoder.decode(parts.get("[Content_Types].xml"));
    expect(types).toContain(`<Default Extension="emf" ContentType="image/x-emf"/>`);
    expect(types).toContain(`<Override PartName="/${F.drawingPart}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`);
    expect(equalBytes(parts.get(F.media.png)!, SPEC131_PICTURE_PNG)).toBe(true);
  });
});

describe("spec131-drawings.ods", () => {
  const O = SPEC131_DRAWINGS_ODS;

  it("is the builder's bytes, whole", async () => {
    const bytes = committed("spec131-drawings.ods");
    expect(equalBytes(bytes, await buildSpec131DrawingsOdsBytes())).toBe(true);
    expect(await packageIntegrity(bytes)).toEqual([]);
  });

  it("holds LibreOffice's drawing shapes", async () => {
    const parts = await unzip(committed("spec131-drawings.ods"));
    const content = decoder.decode(parts.get("content.xml"));
    expect(content).toContain(`<table:table-cell><draw:frame table:end-cell-address="Sheet1.C7"`);
    expect(content).toContain(`xlink:href="${O.picture}"`);
    expect(content).toContain(`<svg:desc>Company logo</svg:desc>`);
    expect(content).toMatch(/<table:table-cell><draw:custom-shape draw:z-index="1" draw:name="Assumptions box"/);
    expect(content).toContain(`</office:forms><table:shapes><draw:custom-shape draw:z-index="2"`);
    expect(content).toContain(`draw:transform="rotate (0.523598775598299) translate (`);
    expect(content).toMatch(/<draw:g table:end-cell-address="Sheet1\.F19"[^>]*><draw:custom-shape[^]*?<\/draw:custom-shape><draw:custom-shape/);
    expect(content).toContain(`<draw:control table:end-cell-address="Sheet1.I4"`);
    for (const style of ["gr1", "gr2", "gr3", "gr4", "gr5", "gr6"]) {
      expect(content).toContain(`<style:style style:name="${style}" style:family="graphic">`);
    }
    const manifest = decoder.decode(parts.get("META-INF/manifest.xml"));
    expect(manifest).toContain(`manifest:full-path="${O.picture}" manifest:media-type="image/png"`);
  });
});
