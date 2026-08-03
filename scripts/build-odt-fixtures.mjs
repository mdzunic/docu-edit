#!/usr/bin/env node
// One-off generator for the SPEC-005 test fixtures under
// src/lib/odf/fixtures/. Not part of the app build — run manually
// (`node scripts/build-odt-fixtures.mjs`) if the fixtures ever need to be
// regenerated or extended.
//
// Like the .docx fixtures (see build-docx-fixtures.mjs), these are NOT files
// exported from real LibreOffice/OpenOffice installs (this environment has
// no network access and neither suite installed) — they are hand-built ODF
// packages that reproduce each producer's structure and known quirks:
//
// - libreoffice-writer.odt: ODF 1.3 as LibreOffice 7.x writes it — automatic
//   P*/T* styles in content.xml, named styles with percentage heading font
//   sizes ("130%") resolved through style:parent-style-name chains,
//   "first space literal, the rest text:s" whitespace encoding, Mpm1 page
//   layout referenced from the Standard master page.
// - openoffice-era.odt: ODF 1.1 as OpenOffice.org 2.x/3.x wrote it —
//   office:version="1.1", inch-based page metrics, named paragraph styles
//   referenced directly from content, a text:h that relies on the style's
//   style:default-outline-level instead of a text:outline-level attribute,
//   and pretty-printed literal whitespace inside a paragraph.
//
// Both are valid ZIP packages (real CRCs, `mimetype` stored uncompressed as
// the first entry per ODF packaging rules), so `mimetype`-based format
// detection sees exactly what it would in a genuine export.
import { deflateRawSync } from "node:zlib";
import { writeFileSync, mkdirSync } from "node:fs";
import { fileURLToPath } from "node:url";
import path from "node:path";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const outDir = path.join(__dirname, "..", "src", "lib", "odf", "fixtures");
mkdirSync(outDir, { recursive: true });

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

function crc32(buf) {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) {
    c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  }
  return (c ^ 0xffffffff) >>> 0;
}

// entries: {name, content, store?} — store:true keeps the entry uncompressed
// (required for the mimetype entry, which must also come first).
function buildZip(entries) {
  const localParts = [];
  const centralParts = [];
  let offset = 0;

  for (const entry of entries) {
    const nameBytes = Buffer.from(entry.name, "utf-8");
    const rawBytes = Buffer.from(entry.content, "utf-8");
    const method = entry.store ? 0 : 8;
    const dataBytes = entry.store ? rawBytes : deflateRawSync(rawBytes);
    const crc = crc32(rawBytes);
    const version = entry.store ? 10 : 20;

    const localHeader = Buffer.concat([
      u32(0x04034b50),
      u16(version),
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
      u16(version),
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
  const localData = Buffer.concat(localParts);
  const eocd = Buffer.concat([
    u32(0x06054b50),
    u16(0),
    u16(0),
    u16(entries.length),
    u16(entries.length),
    u32(centralDirectory.length),
    u32(localData.length),
    u16(0),
  ]);

  return Buffer.concat([localData, centralDirectory, eocd]);
}

const ODT_MIMETYPE = "application/vnd.oasis.opendocument.text";

const NS = [
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"',
  'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"',
  'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"',
  'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"',
  'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"',
  'xmlns:xlink="http://www.w3.org/1999/xlink"',
].join(" ");

function manifestXml() {
  return `<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
 <manifest:file-entry manifest:full-path="/" manifest:media-type="${ODT_MIMETYPE}"/>
 <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
 <manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>
 <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
</manifest:manifest>`;
}

function metaXml(generator, version) {
  return `<?xml version="1.0" encoding="UTF-8"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" office:version="${version}">
 <office:meta><meta:generator>${generator}</meta:generator></office:meta>
</office:document-meta>`;
}

// ---------------------------------------------------------------------
// Fixture A: LibreOffice Writer (ODF 1.3)
// ---------------------------------------------------------------------

const libreOfficeStyles = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles ${NS} office:version="1.3">
 <office:font-face-decls>
  <style:font-face style:name="Liberation Serif" svg:font-family="'Liberation Serif'" style:font-family-generic="roman" style:font-pitch="variable"/>
  <style:font-face style:name="Liberation Sans" svg:font-family="'Liberation Sans'" style:font-family-generic="swiss" style:font-pitch="variable"/>
 </office:font-face-decls>
 <office:styles>
  <style:default-style style:family="paragraph">
   <style:paragraph-properties style:text-autospace="ideograph-alpha" style:line-break="strict"/>
   <style:text-properties style:font-name="Liberation Serif" fo:font-size="12pt" fo:language="en" fo:country="US"/>
  </style:default-style>
  <style:style style:name="Standard" style:family="paragraph" style:class="text"/>
  <style:style style:name="Heading" style:family="paragraph" style:parent-style-name="Standard" style:next-style-name="Text_20_body" style:class="text">
   <style:paragraph-properties fo:margin-top="0.423cm" fo:margin-bottom="0.212cm" fo:keep-with-next="always"/>
   <style:text-properties style:font-name="Liberation Sans" fo:font-size="14pt"/>
  </style:style>
  <style:style style:name="Text_20_body" style:display-name="Text body" style:family="paragraph" style:parent-style-name="Standard" style:class="text">
   <style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="0.247cm" fo:line-height="115%"/>
  </style:style>
  <style:style style:name="Heading_20_1" style:display-name="Heading 1" style:family="paragraph" style:parent-style-name="Heading" style:next-style-name="Text_20_body" style:default-outline-level="1" style:class="text">
   <style:text-properties fo:font-size="130%" fo:font-weight="bold"/>
  </style:style>
  <style:style style:name="Heading_20_2" style:display-name="Heading 2" style:family="paragraph" style:parent-style-name="Heading" style:next-style-name="Text_20_body" style:default-outline-level="2" style:class="text">
   <style:text-properties fo:font-size="115%" fo:font-weight="bold"/>
  </style:style>
 </office:styles>
 <office:automatic-styles>
  <style:page-layout style:name="Mpm1">
   <style:page-layout-properties fo:page-width="21.001cm" fo:page-height="29.7cm" style:num-format="1" style:print-orientation="portrait" fo:margin-top="2cm" fo:margin-bottom="2cm" fo:margin-left="2cm" fo:margin-right="2cm" style:writing-mode="lr-tb"/>
  </style:page-layout>
 </office:automatic-styles>
 <office:master-styles>
  <style:master-page style:name="Standard" style:page-layout-name="Mpm1"/>
 </office:master-styles>
</office:document-styles>`;

const libreOfficeContent = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content ${NS} office:version="1.3">
 <office:font-face-decls>
  <style:font-face style:name="Liberation Serif" svg:font-family="'Liberation Serif'" style:font-family-generic="roman" style:font-pitch="variable"/>
  <style:font-face style:name="Liberation Sans" svg:font-family="'Liberation Sans'" style:font-family-generic="swiss" style:font-pitch="variable"/>
 </office:font-face-decls>
 <office:automatic-styles>
  <style:style style:name="P1" style:family="paragraph" style:parent-style-name="Text_20_body">
   <style:paragraph-properties fo:text-align="center" style:justify-single-word="false" fo:margin-left="1.27cm" fo:margin-right="1.27cm" fo:text-indent="0.635cm" fo:margin-top="0.423cm" fo:margin-bottom="0.423cm"/>
  </style:style>
  <style:style style:name="T1" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style>
  <style:style style:name="T2" style:family="text"><style:text-properties fo:font-style="italic"/></style:style>
  <style:style style:name="T3" style:family="text"><style:text-properties style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"/></style:style>
  <style:style style:name="T4" style:family="text"><style:text-properties style:text-line-through-style="solid" style:text-line-through-type="single"/></style:style>
  <style:style style:name="T5" style:family="text"><style:text-properties style:font-name="Liberation Sans" fo:font-size="14pt" fo:color="#c9211e"/></style:style>
  <style:style style:name="T6" style:family="text"><style:text-properties fo:background-color="#ffff00"/></style:style>
 </office:automatic-styles>
 <office:body>
  <office:text>
   <text:h text:style-name="Heading_20_1" text:outline-level="1">Project Plan</text:h>
   <text:p text:style-name="Text_20_body">Body text with <text:span text:style-name="T1">bold</text:span> and <text:span text:style-name="T2">italic</text:span> and <text:span text:style-name="T3">underlined</text:span> and <text:span text:style-name="T4">struck</text:span> words.</text:p>
   <text:p text:style-name="Text_20_body">Wide <text:s text:c="3"/>gap and tab<text:tab/>stop.</text:p>
   <text:p text:style-name="Text_20_body">Line one<text:line-break/>Line two</text:p>
   <text:h text:style-name="Heading_20_2" text:outline-level="2">Schedule</text:h>
   <text:p text:style-name="Text_20_body"><text:span text:style-name="T5">Red Sans 14pt</text:span> and <text:span text:style-name="T6">highlighted</text:span> text.</text:p>
   <text:p text:style-name="P1">Centered indented paragraph.</text:p>
   <text:p text:style-name="Text_20_body">Nested <text:span text:style-name="T1">bold <text:span text:style-name="T2">bold-italic</text:span></text:span> spans and a <text:a xlink:type="simple" xlink:href="https://example.org/">link text</text:a>.</text:p>
  </office:text>
 </office:body>
</office:document-content>`;

writeFileSync(
  path.join(outDir, "libreoffice-writer.odt"),
  buildZip([
    { name: "mimetype", content: ODT_MIMETYPE, store: true },
    { name: "META-INF/manifest.xml", content: manifestXml() },
    { name: "meta.xml", content: metaXml("LibreOffice/7.6.4.1$Linux_X86_64 LibreOffice_project/60", "1.3") },
    { name: "content.xml", content: libreOfficeContent },
    { name: "styles.xml", content: libreOfficeStyles },
  ]),
);

// ---------------------------------------------------------------------
// Fixture B: OpenOffice.org era (ODF 1.1)
// ---------------------------------------------------------------------

const openOfficeStyles = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles ${NS} office:version="1.1">
 <office:font-face-decls>
  <style:font-face style:name="Times New Roman" svg:font-family="'Times New Roman'" style:font-family-generic="roman" style:font-pitch="variable"/>
  <style:font-face style:name="Arial" svg:font-family="Arial" style:font-family-generic="swiss" style:font-pitch="variable"/>
 </office:font-face-decls>
 <office:styles>
  <style:default-style style:family="paragraph">
   <style:paragraph-properties style:text-autospace="ideograph-alpha"/>
   <style:text-properties style:font-name="Times New Roman" fo:font-size="12pt"/>
  </style:default-style>
  <style:style style:name="Standard" style:family="paragraph" style:class="text"/>
  <style:style style:name="Heading" style:family="paragraph" style:parent-style-name="Standard" style:next-style-name="Text_20_body" style:class="text">
   <style:paragraph-properties fo:margin-top="0.423cm" fo:margin-bottom="0.212cm" fo:keep-with-next="always"/>
   <style:text-properties style:font-name="Arial" fo:font-size="14pt"/>
  </style:style>
  <style:style style:name="Text_20_body" style:display-name="Text body" style:family="paragraph" style:parent-style-name="Standard" style:class="text">
   <style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="0.212cm"/>
  </style:style>
  <style:style style:name="Heading_20_1" style:display-name="Heading 1" style:family="paragraph" style:parent-style-name="Heading" style:default-outline-level="1" style:class="text">
   <style:text-properties fo:font-size="115%" fo:font-weight="bold"/>
  </style:style>
  <style:style style:name="Heading_20_2" style:display-name="Heading 2" style:family="paragraph" style:parent-style-name="Heading" style:default-outline-level="2" style:class="text">
   <style:text-properties fo:font-size="14pt" fo:font-weight="bold" fo:font-style="italic"/>
  </style:style>
  <style:style style:name="Emphasis" style:family="text">
   <style:text-properties fo:font-style="italic"/>
  </style:style>
 </office:styles>
 <office:automatic-styles>
  <style:page-layout style:name="pm1">
   <style:page-layout-properties fo:page-width="11in" fo:page-height="8.5in" style:num-format="1" style:print-orientation="landscape" fo:margin-top="1in" fo:margin-bottom="1in" fo:margin-left="1in" fo:margin-right="1in"/>
  </style:page-layout>
 </office:automatic-styles>
 <office:master-styles>
  <style:master-page style:name="Standard" style:page-layout-name="pm1"/>
 </office:master-styles>
</office:document-styles>`;

const openOfficeContent = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content ${NS} office:version="1.1">
 <office:font-face-decls>
  <style:font-face style:name="Times New Roman" svg:font-family="'Times New Roman'" style:font-family-generic="roman" style:font-pitch="variable"/>
  <style:font-face style:name="Arial" svg:font-family="Arial" style:font-family-generic="swiss" style:font-pitch="variable"/>
 </office:font-face-decls>
 <office:automatic-styles>
  <style:style style:name="P1" style:family="paragraph" style:parent-style-name="Standard">
   <style:paragraph-properties fo:text-align="justify" style:justify-single-word="false"/>
  </style:style>
  <style:style style:name="P2" style:family="paragraph" style:parent-style-name="Standard">
   <style:paragraph-properties fo:text-align="end"/>
  </style:style>
  <style:style style:name="P3" style:family="paragraph" style:parent-style-name="Standard">
   <style:paragraph-properties fo:margin-left="0.5in" fo:margin-right="0in" fo:text-indent="-0.25in"/>
  </style:style>
  <style:style style:name="T1" style:family="text">
   <style:text-properties style:text-underline-style="dotted" style:text-underline-width="auto" style:text-underline-color="font-color"/>
  </style:style>
 </office:automatic-styles>
 <office:body>
  <office:text>
   <text:h text:style-name="Heading_20_1" text:outline-level="1">Annual Review</text:h>
   <text:p text:style-name="Standard">
     Literal whitespace   collapses to single spaces.</text:p>
   <text:h text:style-name="Heading_20_2">Financials</text:h>
   <text:p text:style-name="P1">Justified <text:span text:style-name="Emphasis">emphasised</text:span> paragraph.</text:p>
   <text:p text:style-name="P2">Right aligned via end.</text:p>
   <text:p text:style-name="P3">Hanging indent paragraph.</text:p>
   <text:p text:style-name="Standard">Dotted <text:span text:style-name="T1">underline</text:span>.</text:p>
   <text:p text:style-name="Standard"/>
  </office:text>
 </office:body>
</office:document-content>`;

writeFileSync(
  path.join(outDir, "openoffice-era.odt"),
  buildZip([
    { name: "mimetype", content: ODT_MIMETYPE, store: true },
    { name: "META-INF/manifest.xml", content: manifestXml() },
    { name: "meta.xml", content: metaXml("OpenOffice.org/3.2$Linux OpenOffice.org_project/320m12$Build-9483", "1.1") },
    { name: "content.xml", content: openOfficeContent },
    { name: "styles.xml", content: openOfficeStyles },
  ]),
);

console.log("Wrote fixtures to", outDir);
