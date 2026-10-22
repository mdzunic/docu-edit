#!/usr/bin/env node
// One-off generator for the two SPEC-008 test fixtures:
//   src/lib/docx/fixtures/word-charts-and-controls.docx
//   src/lib/odf/fixtures/lo-embedded.odt
// Not part of the app build — run manually
// (`node scripts/build-spec008-fixtures.mjs`) if the fixtures ever need to
// be regenerated or extended.
//
// Like the SPEC-004/005 fixtures (see build-docx-fixtures.mjs), these are
// hand-built packages, not genuine exports: this environment has no Word or
// LibreOffice. word-charts-and-controls.docx carries the three things
// SPEC-008's AC names — a DrawingML chart (part + inline w:drawing), a
// block-level w:sdt content control, and a customXml part — plus a
// docProps/core.xml so a same-format save replaces it rather than patching
// [Content_Types].xml/_rels/.rels. lo-embedded.odt carries an embedded
// image (Pictures/ part + inline draw:frame) and a table:table, with the
// graphic/table automatic styles those nodes reference.
import { deflateRawSync } from "node:zlib";
import { writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import path from "node:path";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function u32(value) {
  const buf = Buffer.alloc(4);
  buf.writeUInt32LE(value >>> 0, 0);
  return buf;
}
function u16(value) {
  const buf = Buffer.alloc(2);
  buf.writeUInt16LE(value, 0);
  return buf;
}

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    table[n] = c >>> 0;
  }
  return table;
})();

function crc32(bytes) {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc = CRC_TABLE[(crc ^ byte) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

/** entries: { name, content: string | Buffer, store?: boolean }[] */
function buildZip(entries) {
  const localParts = [];
  const centralParts = [];
  let offset = 0;

  for (const entry of entries) {
    const nameBytes = Buffer.from(entry.name, "utf-8");
    const rawBytes = Buffer.isBuffer(entry.content) ? entry.content : Buffer.from(entry.content, "utf-8");
    const dataBytes = entry.store ? rawBytes : deflateRawSync(rawBytes);
    const method = entry.store ? 0 : 8;
    const crc = crc32(rawBytes);

    const localHeader = Buffer.concat([
      u32(0x04034b50),
      u16(20),
      u16(0),
      u16(method),
      u16(0),
      u16(0),
      u32(crc),
      u32(dataBytes.length),
      u32(rawBytes.length),
      u16(nameBytes.length),
      u16(0),
      nameBytes,
    ]);
    localParts.push(localHeader, dataBytes);

    const centralHeader = Buffer.concat([
      u32(0x02014b50),
      u16(20),
      u16(20),
      u16(0),
      u16(method),
      u16(0),
      u16(0),
      u32(crc),
      u32(dataBytes.length),
      u32(rawBytes.length),
      u16(nameBytes.length),
      u16(0),
      u16(0),
      u16(0),
      u16(0),
      u32(0),
      u32(offset),
      nameBytes,
    ]);
    centralParts.push(centralHeader);
    offset += localHeader.length + dataBytes.length;
  }

  const centralDirectory = Buffer.concat(centralParts);
  const eocd = Buffer.concat([
    u32(0x06054b50),
    u16(0),
    u16(0),
    u16(entries.length),
    u16(entries.length),
    u32(centralDirectory.length),
    u32(offset),
    u16(0),
  ]);
  return Buffer.concat([...localParts, centralDirectory, eocd]);
}

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

// ---------------------------------------------------------------------------
// word-charts-and-controls.docx
// ---------------------------------------------------------------------------

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const docxContentTypes =
  `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
  '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
  '<Override PartName="/word/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>' +
  '<Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>' +
  '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
  "</Types>";

const docxPackageRels =
  `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/>` +
  '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
  "</Relationships>";

const docxDocumentRels =
  `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId1" Type="${R}/styles" Target="styles.xml"/>` +
  `<Relationship Id="rId4" Type="${R}/chart" Target="charts/chart1.xml"/>` +
  `<Relationship Id="rId5" Type="${R}/customXml" Target="../customXml/item1.xml"/>` +
  "</Relationships>";

const docxDocument =
  `${XML_DECL}<w:document xmlns:w="${W}"` +
  ` xmlns:r="${R}"` +
  ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"' +
  ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
  ' xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
  "<w:body>" +
  '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Annual Metrics</w:t></w:r></w:p>' +
  "<w:p><w:r><w:t>Editable intro paragraph.</w:t></w:r></w:p>" +
  '<w:p><w:r><w:t xml:space="preserve">Quarterly results: </w:t></w:r>' +
  '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
  '<wp:extent cx="5486400" cy="3200400"/><wp:docPr id="1" name="Chart 1"/>' +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
  '<c:chart r:id="rId4"/></a:graphicData></a:graphic>' +
  "</wp:inline></w:drawing></w:r></w:p>" +
  '<w:sdt><w:sdtPr><w:alias w:val="Status"/><w:tag w:val="status"/></w:sdtPr>' +
  "<w:sdtContent><w:p><w:r><w:t>Controlled block text</w:t></w:r></w:p></w:sdtContent></w:sdt>" +
  "<w:p><w:r><w:t>Closing remarks.</w:t></w:r></w:p>" +
  "<w:sectPr>" +
  '<w:pgSz w:w="11906" w:h="16838"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>' +
  "</w:sectPr>" +
  "</w:body></w:document>";

const docxStyles =
  `${XML_DECL}<w:styles xmlns:w="${W}">` +
  "<w:docDefaults><w:rPrDefault><w:rPr>" +
  '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/>' +
  "</w:rPr></w:rPrDefault></w:docDefaults>" +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>' +
  "</w:styles>";

const docxChart =
  `${XML_DECL}<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"` +
  ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
  "<c:chart><c:title><c:overlay c:val=\"0\"/></c:title><c:plotArea><c:layout/>" +
  '<c:barChart><c:barDir c:val="col"/><c:grouping c:val="clustered"/>' +
  "<c:ser><c:idx c:val=\"0\"/><c:order c:val=\"0\"/>" +
  "<c:val><c:numLit><c:pt c:idx=\"0\"><c:v>42</c:v></c:pt><c:pt c:idx=\"1\"><c:v>17</c:v></c:pt></c:numLit></c:val>" +
  "</c:ser></c:barChart></c:plotArea><c:plotVisOnly c:val=\"1\"/></c:chart></c:chartSpace>";

const docxCustomXmlItem =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<metrics xmlns="urn:docu-edit:spec008"><quarter>Q3</quarter><region>EMEA</region></metrics>';

const docxCustomXmlItemProps =
  `${XML_DECL}<ds:datastoreItem ds:itemID="{8E4B62D1-3C1A-4C57-9E1B-5A1DBB4A2F10}"` +
  ' xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">' +
  '<ds:schemaRefs><ds:schemaRef ds:uri="urn:docu-edit:spec008"/></ds:schemaRefs></ds:datastoreItem>';

const docxCustomXmlRels =
  `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps1.xml"/>' +
  "</Relationships>";

const docxCore =
  `${XML_DECL}<cp:coreProperties` +
  ' xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"' +
  ' xmlns:dc="http://purl.org/dc/elements/1.1/"' +
  ' xmlns:dcterms="http://purl.org/dc/terms/"' +
  ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
  "<dc:title>Annual Metrics</dc:title>" +
  '<dcterms:created xsi:type="dcterms:W3CDTF">2026-05-01T09:00:00Z</dcterms:created>' +
  '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-05-01T09:00:00Z</dcterms:modified>' +
  "</cp:coreProperties>";

writeFileSync(
  path.join(__dirname, "..", "src", "lib", "docx", "fixtures", "word-charts-and-controls.docx"),
  buildZip([
    { name: "[Content_Types].xml", content: docxContentTypes },
    { name: "_rels/.rels", content: docxPackageRels },
    { name: "word/document.xml", content: docxDocument },
    { name: "word/_rels/document.xml.rels", content: docxDocumentRels },
    { name: "word/styles.xml", content: docxStyles },
    { name: "word/charts/chart1.xml", content: docxChart },
    { name: "customXml/item1.xml", content: docxCustomXmlItem },
    { name: "customXml/itemProps1.xml", content: docxCustomXmlItemProps },
    { name: "customXml/_rels/item1.xml.rels", content: docxCustomXmlRels },
    { name: "docProps/core.xml", content: docxCore },
  ]),
);

// ---------------------------------------------------------------------------
// lo-embedded.odt
// ---------------------------------------------------------------------------

const ODT_MIMETYPE = "application/vnd.oasis.opendocument.text";
const ODF_DECL = '<?xml version="1.0" encoding="UTF-8"?>';
const ODF_XMLNS =
  ' xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"' +
  ' xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"' +
  ' xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"' +
  ' xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"' +
  ' xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"' +
  ' xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"' +
  ' xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"' +
  ' xmlns:xlink="http://www.w3.org/1999/xlink"';

const odtContent =
  `${ODF_DECL}<office:document-content${ODF_XMLNS} office:version="1.3">` +
  "<office:font-face-decls>" +
  '<style:font-face style:name="Liberation Serif" svg:font-family="\'Liberation Serif\'" style:font-family-generic="roman" style:font-pitch="variable"/>' +
  "</office:font-face-decls>" +
  "<office:automatic-styles>" +
  '<style:style style:name="P1" style:family="paragraph" style:parent-style-name="Standard">' +
  '<style:paragraph-properties fo:text-align="center"/></style:style>' +
  '<style:style style:name="fr1" style:family="graphic">' +
  '<style:graphic-properties fo:margin-left="0.1cm" fo:margin-right="0.1cm"/></style:style>' +
  '<style:style style:name="Table1" style:family="table">' +
  '<style:table-properties style:width="17cm" table:align="margins"/></style:style>' +
  '<style:style style:name="Table1.A" style:family="table-column">' +
  '<style:table-column-properties style:column-width="8.5cm"/></style:style>' +
  "</office:automatic-styles>" +
  "<office:body><office:text>" +
  '<text:h text:style-name="Heading_20_1" text:outline-level="1">Field Report</text:h>' +
  '<text:p text:style-name="Standard">Editable intro paragraph.</text:p>' +
  '<text:p text:style-name="Standard">Photo: <draw:frame draw:style-name="fr1" draw:name="Image1" text:anchor-type="as-char" svg:width="2cm" svg:height="2cm" draw:z-index="0">' +
  '<draw:image xlink:href="Pictures/image1.png" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>' +
  "</draw:frame> shows the site.</text:p>" +
  '<table:table table:name="Table1" table:style-name="Table1">' +
  '<table:table-column table:style-name="Table1.A" table:number-columns-repeated="2"/>' +
  '<table:table-row><table:table-cell office:value-type="string"><text:p>Cell A</text:p></table:table-cell>' +
  '<table:table-cell office:value-type="string"><text:p>Cell B</text:p></table:table-cell></table:table-row>' +
  "</table:table>" +
  '<text:p text:style-name="P1">Closing summary.</text:p>' +
  "</office:text></office:body></office:document-content>";

const odtStyles =
  `${ODF_DECL}<office:document-styles${ODF_XMLNS} office:version="1.3">` +
  "<office:styles>" +
  '<style:default-style style:family="paragraph">' +
  '<style:text-properties fo:font-family="Liberation Serif" fo:font-size="12pt"/>' +
  "</style:default-style>" +
  '<style:style style:name="Standard" style:family="paragraph" style:class="text"/>' +
  '<style:style style:name="Heading_20_1" style:display-name="Heading 1" style:family="paragraph" style:parent-style-name="Standard" style:default-outline-level="1">' +
  '<style:text-properties fo:font-size="18pt" fo:font-weight="bold"/></style:style>' +
  "</office:styles>" +
  "<office:automatic-styles>" +
  '<style:page-layout style:name="pm1">' +
  '<style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm"' +
  ' fo:margin-top="2cm" fo:margin-right="2cm" fo:margin-bottom="2cm" fo:margin-left="2cm"/>' +
  "</style:page-layout>" +
  "</office:automatic-styles>" +
  "<office:master-styles>" +
  '<style:master-page style:name="Standard" style:page-layout-name="pm1"/>' +
  "</office:master-styles>" +
  "</office:document-styles>";

const odtMeta =
  `${ODF_DECL}<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"` +
  ' xmlns:dc="http://purl.org/dc/elements/1.1/" office:version="1.3">' +
  "<office:meta><dc:title>Field Report</dc:title><dc:date>2026-05-02T10:00:00Z</dc:date></office:meta>" +
  "</office:document-meta>";

const odtSettings =
  `${ODF_DECL}<office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:version="1.3"/>`;

const odtManifest =
  `${ODF_DECL}<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">` +
  `<manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="${ODT_MIMETYPE}"/>` +
  '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>' +
  '<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>' +
  '<manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>' +
  '<manifest:file-entry manifest:full-path="settings.xml" manifest:media-type="text/xml"/>' +
  '<manifest:file-entry manifest:full-path="Pictures/image1.png" manifest:media-type="image/png"/>' +
  "</manifest:manifest>";

// A valid 1x1 transparent PNG — small, binary, and never valid UTF-8, so a
// byte-identity failure cannot hide behind text re-encoding.
const onePixelPng = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==",
  "base64",
);

writeFileSync(
  path.join(__dirname, "..", "src", "lib", "odf", "fixtures", "lo-embedded.odt"),
  buildZip([
    { name: "mimetype", content: ODT_MIMETYPE, store: true },
    { name: "META-INF/manifest.xml", content: odtManifest },
    { name: "content.xml", content: odtContent },
    { name: "styles.xml", content: odtStyles },
    { name: "meta.xml", content: odtMeta },
    { name: "settings.xml", content: odtSettings },
    { name: "Pictures/image1.png", content: onePixelPng },
  ]),
);

console.log("SPEC-008 fixtures written.");
