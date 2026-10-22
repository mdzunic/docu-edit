#!/usr/bin/env node
// One-off generator for the SPEC-052 chart fixtures (run manually:
// `node scripts/build-spec052-fixtures.mjs`). Like build-spec051-fixtures.mjs
// (shapes), these are hand-built packages reproducing what Word/LibreOffice
// would emit for charts — NOT round-tripped through this project's own
// writer (src/lib/docx/chart.ts, src/lib/odf/chart.ts) — so opening them is a
// genuine test of the reader, not just "the reader understands its own
// writer's output".
//
// Both fixtures carry the same 12 charts in document order, addressable by
// their 0-based `data-testid="drawing-<n>"` index:
//   0 columnClustered   4 lineMarkers   8 scatter
//   1 columnStacked     5 area          9 columnClustered, negative values
//   2 barClustered      6 pie          10 columnClustered, one gap (empty cell)
//   3 line              7 doughnut     11 unsupported (a 3-D bar chart / ODF stock)
import { deflateRawSync } from "node:zlib";
import { writeFileSync, mkdirSync } from "node:fs";
import { fileURLToPath } from "node:url";
import path from "node:path";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function u32le(value) {
  const buf = Buffer.alloc(4);
  buf.writeUInt32LE(value >>> 0, 0);
  return buf;
}
function u16le(value) {
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

// --- minimal ZIP writer (same shape as build-spec051-fixtures.mjs) ---------
function buildZip(entries) {
  const localParts = [];
  const centralParts = [];
  let offset = 0;

  for (const entry of entries) {
    const nameBytes = Buffer.from(entry.name, "utf-8");
    const rawBytes = Buffer.isBuffer(entry.content) ? entry.content : Buffer.from(entry.content, "utf-8");
    const method = entry.store ? 0 : 8;
    const dataBytes = entry.store ? rawBytes : deflateRawSync(rawBytes);
    const crc = crc32(rawBytes);
    const version = entry.store ? 10 : 20;

    const localHeader = Buffer.concat([
      u32le(0x04034b50),
      u16le(version),
      u16le(0),
      u16le(method),
      u16le(0),
      u16le(0),
      u32le(crc),
      u32le(dataBytes.length),
      u32le(rawBytes.length),
      u16le(nameBytes.length),
      u16le(0),
      nameBytes,
    ]);
    localParts.push(localHeader, dataBytes);

    const centralHeader = Buffer.concat([
      u32le(0x02014b50),
      u16le(20),
      u16le(version),
      u16le(0),
      u16le(method),
      u16le(0),
      u16le(0),
      u32le(crc),
      u32le(dataBytes.length),
      u32le(rawBytes.length),
      u16le(nameBytes.length),
      u16le(0),
      u16le(0),
      u16le(0),
      u16le(0),
      u32le(0),
      u32le(offset),
      nameBytes,
    ]);
    centralParts.push(centralHeader);
    offset += localHeader.length + dataBytes.length;
  }

  const centralStart = offset;
  const centralBytes = Buffer.concat(centralParts);
  const end = Buffer.concat([
    u32le(0x06054b50),
    u16le(0),
    u16le(0),
    u16le(entries.length),
    u16le(entries.length),
    u32le(centralBytes.length),
    u32le(centralStart),
    u16le(0),
  ]);
  return Buffer.concat([...localParts, centralBytes, end]);
}

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

function escapeXml(text) {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

const CATEGORIES = ["Jan", "Feb", "Mar"];
const SERIES = [
  { name: "Series 1", values: [4.3, 2.5, 3.5], color: "4472C4" },
  { name: "Series 2", values: [2.4, 4.4, 1.8], color: "ED7D31" },
];
const NEGATIVE_SERIES = [{ name: "Net change", values: [5, -3, 2], color: "4472C4" }];
const GAP_SERIES = [{ name: "Attendance", values: [10, null, 15], color: "4472C4" }];

// The fixed sequence both builders walk, one entry per chart in the fixture
// (see the header comment for the resulting drawing-<n> addressing).
const CHARTS = [
  { key: "columnClustered", label: "Clustered column", categories: CATEGORIES, series: SERIES },
  { key: "columnStacked", label: "Stacked column", categories: CATEGORIES, series: SERIES },
  { key: "barClustered", label: "Clustered bar", categories: CATEGORIES, series: SERIES },
  { key: "line", label: "Line", categories: CATEGORIES, series: SERIES },
  { key: "lineMarkers", label: "Line with markers", categories: CATEGORIES, series: SERIES },
  { key: "area", label: "Area", categories: CATEGORIES, series: SERIES },
  { key: "pie", label: "Pie", categories: CATEGORIES, series: [SERIES[0]] },
  { key: "doughnut", label: "Doughnut", categories: CATEGORIES, series: [SERIES[0]] },
  { key: "scatter", label: "Scatter", categories: ["1", "2", "3"], series: SERIES },
  { key: "columnClustered", label: "Negative values", categories: CATEGORIES, series: NEGATIVE_SERIES },
  { key: "columnClustered", label: "Gap (empty cell)", categories: CATEGORIES, series: GAP_SERIES },
  { key: "unsupported", label: "Unsupported (3-D)", categories: CATEGORIES, series: SERIES },
];

// --- .docx fixture -----------------------------------------------------------
{
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
  const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
  const C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
  const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

  const cmToEmu = (cm) => Math.round(cm * 360000);

  function colLetter(i) {
    return String.fromCharCode(66 + i);
  }

  function catXml(categories) {
    const pts = categories.map((c, i) => `<c:pt idx="${i}"><c:v>${escapeXml(c)}</c:v></c:pt>`).join("");
    return (
      `<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$${categories.length + 1}</c:f>` +
      `<c:strCache><c:ptCount val="${categories.length}"/>${pts}</c:strCache></c:strRef></c:cat>`
    );
  }

  function valXml(values, col) {
    const pts = values.map((v, i) => (v === null ? "" : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`)).join("");
    return (
      `<c:val><c:numRef><c:f>Sheet1!$${col}$2:$${col}$${values.length + 1}</c:f>` +
      `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${pts}</c:numCache></c:numRef></c:val>`
    );
  }

  function serXml(series, categories, index, extra = "") {
    const col = colLetter(index);
    return (
      `<c:ser><c:idx val="${index}"/><c:order val="${index}"/>` +
      `<c:tx><c:strRef><c:f>Sheet1!$${col}$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${escapeXml(series.name)}</c:v></c:pt></c:strCache></c:strRef></c:tx>` +
      `<c:spPr><a:solidFill><a:srgbClr val="${series.color}"/></a:solidFill></c:spPr>` +
      extra +
      catXml(categories) +
      valXml(series.values, col) +
      "</c:ser>"
    );
  }

  function axesXml() {
    return (
      '<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>' +
      '<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:crossAx val="1"/></c:valAx>'
    );
  }

  function chartTypeXml(entry) {
    const { key, categories, series } = entry;
    if (key === "columnClustered" || key === "columnStacked" || key === "barClustered") {
      const dir = key === "barClustered" ? "bar" : "col";
      const grouping = key === "columnStacked" ? "stacked" : "clustered";
      const serXmlAll = series.map((s, i) => serXml(s, categories, i)).join("");
      return (
        `<c:barChart><c:barDir val="${dir}"/><c:grouping val="${grouping}"/><c:varyColors val="0"/>${serXmlAll}` +
        `<c:gapWidth val="150"/>${grouping === "stacked" ? '<c:overlap val="100"/>' : ""}` +
        '<c:axId val="1"/><c:axId val="2"/></c:barChart>' +
        axesXml()
      );
    }
    if (key === "line" || key === "lineMarkers") {
      const marker = key === "lineMarkers" ? "1" : "0";
      const serXmlAll = series
        .map((s, i) => serXml(s, categories, i, `<c:marker><c:symbol val="${key === "lineMarkers" ? "circle" : "none"}"/></c:marker>`))
        .join("");
      return (
        `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${serXmlAll}` +
        `<c:marker val="${marker}"/><c:axId val="1"/><c:axId val="2"/></c:lineChart>` +
        axesXml()
      );
    }
    if (key === "area") {
      const serXmlAll = series.map((s, i) => serXml(s, categories, i)).join("");
      return (
        `<c:areaChart><c:grouping val="standard"/><c:varyColors val="0"/>${serXmlAll}` +
        '<c:axId val="1"/><c:axId val="2"/></c:areaChart>' +
        axesXml()
      );
    }
    if (key === "pie" || key === "doughnut") {
      const serXmlAll = serXml(series[0], categories, 0);
      const holeXml = key === "doughnut" ? '<c:holeSize val="50"/>' : "";
      const tag = key === "pie" ? "pieChart" : "doughnutChart";
      return `<c:${tag}><c:varyColors val="1"/>${serXmlAll}<c:firstSliceAng val="0"/>${holeXml}</c:${tag}>`;
    }
    if (key === "scatter") {
      const serXmlAll = series
        .map((s, i) => {
          const col = colLetter(i);
          const xPts = categories.map((c, idx) => `<c:pt idx="${idx}"><c:v>${c}</c:v></c:pt>`).join("");
          const xVal =
            `<c:xVal><c:numRef><c:f>Sheet1!$A$2:$A$${categories.length + 1}</c:f>` +
            `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${categories.length}"/>${xPts}</c:numCache></c:numRef></c:xVal>`;
          return (
            `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>` +
            `<c:tx><c:strRef><c:f>Sheet1!$${col}$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${escapeXml(s.name)}</c:v></c:pt></c:strCache></c:strRef></c:tx>` +
            `<c:spPr><a:solidFill><a:srgbClr val="${s.color}"/></a:solidFill></c:spPr>` +
            xVal +
            valXml(s.values, col) +
            "<c:smooth val=\"0\"/></c:ser>"
          );
        })
        .join("");
      return (
        `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${serXmlAll}` +
        '<c:axId val="1"/><c:axId val="2"/></c:scatterChart>' +
        '<c:valAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:valAx>' +
        '<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>'
      );
    }
    // "unsupported": a real 3-D bar chart, outside the catalogue.
    const serXmlAll = series.map((s, i) => serXml(s, categories, i)).join("");
    return (
      `<c:bar3DChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${serXmlAll}` +
      '<c:shape val="box"/><c:axId val="1"/><c:axId val="2"/><c:axId val="3"/></c:bar3DChart>' +
      '<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>' +
      '<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>' +
      '<c:serAx><c:axId val="3"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:serAx>'
    );
  }

  function chartSpaceXml(entry) {
    return (
      `${XML}<c:chartSpace xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${R}">` +
      `<c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>${escapeXml(entry.label)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>` +
      `<c:plotArea><c:layout/>${chartTypeXml(entry)}</c:plotArea>` +
      '<c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend>' +
      '<c:plotVisOnly val="1"/></c:chart></c:chartSpace>'
    );
  }

  const chartParts = [];
  const chartRels = [];
  const drawings = [];
  CHARTS.forEach((entry, index) => {
    const n = index + 1;
    chartParts.push({ name: `word/charts/chart${n}.xml`, content: chartSpaceXml(entry) });
    const relId = `rId${n}`;
    chartRels.push(`<Relationship Id="${relId}" Type="${R}/chart" Target="charts/chart${n}.xml"/>`);
    const cx = cmToEmu(12);
    const cy = cmToEmu(7.5);
    drawings.push(
      `<w:p><w:r><w:drawing xmlns:wp="${WP}"><wp:inline distT="0" distB="0" distL="0" distR="0">` +
        `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
        `<wp:docPr id="${n}" name="Chart ${n}" descr="${escapeXml(entry.label)}"/><wp:cNvGraphicFramePr/>` +
        `<a:graphic xmlns:a="${A}"><a:graphicData uri="${C}"><c:chart xmlns:c="${C}" xmlns:r="${R}" r:id="${relId}"/></a:graphicData></a:graphic>` +
        "</wp:inline></w:drawing></w:r></w:p>",
    );
  });

  const documentXml =
    `${XML}<w:document xmlns:w="${W}"><w:body>` +
    drawings.join("") +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
    "</w:sectPr></w:body></w:document>";

  const stylesXml =
    `${XML}<w:styles xmlns:w="${W}">` +
    "<w:docDefaults><w:rPrDefault><w:rPr>" +
    '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/>' +
    "</w:rPr></w:rPrDefault></w:docDefaults>" +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
    "</w:styles>";

  const chartOverrides = CHARTS.map(
    (_, i) =>
      `<Override PartName="/word/charts/chart${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`,
  ).join("");
  const contentTypesXml =
    `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
    chartOverrides +
    "</Types>";

  const packageRelsXml =
    `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
    "</Relationships>";

  const documentRelsXml =
    `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
    chartRels.join("") +
    "</Relationships>";

  const corePropsXml =
    `${XML}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"` +
    ' xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"' +
    ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    "<dc:title>Chart types fixture</dc:title>" +
    '<dcterms:created xsi:type="dcterms:W3CDTF">2024-01-01T00:00:00Z</dcterms:created>' +
    '<dcterms:modified xsi:type="dcterms:W3CDTF">2024-01-01T00:00:00Z</dcterms:modified>' +
    "</cp:coreProperties>";

  const bytes = buildZip([
    { name: "[Content_Types].xml", content: contentTypesXml },
    { name: "_rels/.rels", content: packageRelsXml },
    { name: "word/document.xml", content: documentXml },
    { name: "word/_rels/document.xml.rels", content: documentRelsXml },
    { name: "word/styles.xml", content: stylesXml },
    { name: "docProps/core.xml", content: corePropsXml },
    ...chartParts,
  ]);

  for (const outDir of [
    path.join(__dirname, "..", "e2e", "fixtures"),
    path.join(__dirname, "..", "src", "lib", "docx", "fixtures"),
  ]) {
    mkdirSync(outDir, { recursive: true });
    writeFileSync(path.join(outDir, "word-chart-types.docx"), bytes);
    console.log("Wrote", path.join(outDir, "word-chart-types.docx"));
  }
}

// --- .odt fixture ------------------------------------------------------------
{
  const OFFICE = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
  const CHART_NS = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
  const TABLE_NS = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
  const TEXT_NS = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
  const SVG_NS = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
  const DRAW_NS = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
  const XLINK_NS = "http://www.w3.org/1999/xlink";

  function tableXml(categories, series) {
    const headerCells = series.map((s) => `<table:table-cell office:value-type="string"><text:p>${escapeXml(s.name)}</text:p></table:table-cell>`).join("");
    const headerRow = `<table:table-row><table:table-cell/>${headerCells}</table:table-row>`;
    const bodyRows = categories
      .map((cat, rowIndex) => {
        const catCell = `<table:table-cell office:value-type="string"><text:p>${escapeXml(cat)}</text:p></table:table-cell>`;
        const valueCells = series
          .map((s) => {
            const v = s.values[rowIndex];
            return v === null
              ? "<table:table-cell/>"
              : `<table:table-cell office:value-type="float" office:value="${v}"><text:p>${v}</text:p></table:table-cell>`;
          })
          .join("");
        return `<table:table-row>${catCell}${valueCells}</table:table-row>`;
      })
      .join("");
    const columns = series.map(() => "<table:table-column/>").join("");
    return (
      '<table:table table:name="local-table">' +
      `<table:table-header-columns><table:table-column/></table:table-header-columns><table:table-columns>${columns}</table:table-columns>` +
      `<table:table-header-rows>${headerRow}</table:table-header-rows><table:table-rows>${bodyRows}</table:table-rows>` +
      "</table:table>"
    );
  }

  function seriesXml(series, index, chartClass) {
    const col = String.fromCharCode(66 + index);
    return (
      `<chart:series chart:class="${chartClass}" chart:color="#${series.color.toLowerCase()}"` +
      ` chart:values-cell-range-address="local-table.$${col}$2:.$${col}$4" chart:label-cell-address="local-table.$${col}$1"/>`
    );
  }

  function chartClassAndAttrs(entry) {
    switch (entry.key) {
      case "columnClustered":
        return { chartClass: "chart:bar", extra: ' chart:vertical="true" chart:stacked="false"' };
      case "columnStacked":
        return { chartClass: "chart:bar", extra: ' chart:vertical="true" chart:stacked="true"' };
      case "barClustered":
        return { chartClass: "chart:bar", extra: ' chart:vertical="false" chart:stacked="false"' };
      case "line":
        return { chartClass: "chart:line", extra: ' chart:symbol-type="none"' };
      case "lineMarkers":
        return { chartClass: "chart:line", extra: ' chart:symbol-type="named-symbol"' };
      case "area":
        return { chartClass: "chart:area", extra: "" };
      case "pie":
        return { chartClass: "chart:circle", extra: "" };
      case "doughnut":
        return { chartClass: "chart:ring", extra: "" };
      case "scatter":
        return { chartClass: "chart:scatter", extra: "" };
      default:
        // "unsupported": a real chart:stock class, outside the catalogue.
        return { chartClass: "chart:stock", extra: "" };
    }
  }

  function objectContentXml(entry) {
    const { chartClass, extra } = chartClassAndAttrs(entry);
    const seriesForXml = entry.key === "pie" || entry.key === "doughnut" ? [entry.series[0]] : entry.series;
    const seriesXmlAll = seriesForXml.map((s, i) => seriesXml(s, i, chartClass)).join("");
    const hasAxes = entry.key !== "pie" && entry.key !== "doughnut" && entry.key !== "unsupported";
    const axesXml = hasAxes
      ? '<chart:axis chart:dimension="x" chart:name="primary-x"/><chart:axis chart:dimension="y" chart:name="primary-y"><chart:grid chart:class="major"/></chart:axis>'
      : "";
    return (
      '<?xml version="1.0" encoding="UTF-8"?>' +
      `<office:document-content xmlns:office="${OFFICE}" xmlns:chart="${CHART_NS}" xmlns:table="${TABLE_NS}" xmlns:text="${TEXT_NS}" xmlns:svg="${SVG_NS}" office:version="1.2">` +
      "<office:body><office:chart>" +
      `<chart:chart chart:class="${chartClass}"${extra} svg:width="12cm" svg:height="7.5cm">` +
      `<chart:title><text:p>${escapeXml(entry.label)}</text:p></chart:title>` +
      '<chart:legend chart:legend-position="end"/>' +
      `<chart:plot-area>${axesXml}${seriesXmlAll}</chart:plot-area>` +
      tableXml(entry.categories, seriesForXml) +
      "</chart:chart>" +
      "</office:chart></office:body>" +
      "</office:document-content>"
    );
  }

  const objectParts = [];
  const manifestObjectEntries = [];
  const frames = [];
  CHARTS.forEach((entry, index) => {
    const n = index + 1;
    const objectDir = `Object ${n}`;
    objectParts.push({ name: `${objectDir}/content.xml`, content: objectContentXml(entry) });
    manifestObjectEntries.push(
      `<manifest:file-entry manifest:full-path="${objectDir}/" manifest:media-type="application/vnd.oasis.opendocument.chart"/>` +
        `<manifest:file-entry manifest:full-path="${objectDir}/content.xml" manifest:media-type="text/xml"/>`,
    );
    frames.push(
      `<text:p><draw:frame xmlns:draw="${DRAW_NS}" xmlns:xlink="${XLINK_NS}" xmlns:svg="${SVG_NS}" text:anchor-type="as-char" svg:width="12cm" svg:height="7.5cm">` +
        `<draw:object xlink:href="./${objectDir}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>` +
        `<svg:desc>${escapeXml(entry.label)}</svg:desc>` +
        "</draw:frame></text:p>",
    );
  });

  const contentXml =
    '<?xml version="1.0" encoding="UTF-8"?>' +
    `<office:document-content xmlns:office="${OFFICE}" xmlns:text="${TEXT_NS}" xmlns:draw="${DRAW_NS}" xmlns:svg="${SVG_NS}" xmlns:xlink="${XLINK_NS}" xmlns:table="${TABLE_NS}" xmlns:chart="${CHART_NS}" office:version="1.2">` +
    `<office:body><office:text>${frames.join("")}</office:text></office:body>` +
    "</office:document-content>";

  const stylesXml =
    `${XML}<office:document-styles xmlns:office="${OFFICE}" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" office:version="1.2">` +
    "<office:styles>" +
    '<style:default-style style:family="paragraph"><style:text-properties fo:font-family="Liberation Sans" fo:font-size="11pt"/></style:default-style>' +
    "</office:styles>" +
    "<office:automatic-styles>" +
    '<style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm"' +
    ' fo:margin-top="2cm" fo:margin-right="2cm" fo:margin-bottom="2cm" fo:margin-left="2cm"/></style:page-layout>' +
    "</office:automatic-styles>" +
    `<office:master-styles><style:master-page style:name="Standard" style:page-layout-name="pm1"/></office:master-styles>` +
    "</office:document-styles>";

  const metaXml =
    `${XML}<office:document-meta xmlns:office="${OFFICE}" xmlns:dc="http://purl.org/dc/elements/1.1/" office:version="1.2">` +
    "<office:meta><dc:title>Chart types fixture</dc:title><dc:date>2024-01-01T00:00:00Z</dc:date></office:meta>" +
    "</office:document-meta>";

  const settingsXml =
    `<?xml version="1.0" encoding="UTF-8"?><office:document-settings xmlns:office="${OFFICE}" office:version="1.2"/>`;

  const manifestXml =
    '<?xml version="1.0" encoding="UTF-8"?>' +
    '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">' +
    '<manifest:file-entry manifest:full-path="/" manifest:version="1.2" manifest:media-type="application/vnd.oasis.opendocument.text"/>' +
    '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>' +
    '<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>' +
    '<manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>' +
    '<manifest:file-entry manifest:full-path="settings.xml" manifest:media-type="text/xml"/>' +
    manifestObjectEntries.join("") +
    "</manifest:manifest>";

  const bytes = buildZip([
    { name: "mimetype", content: "application/vnd.oasis.opendocument.text", store: true },
    { name: "content.xml", content: contentXml },
    { name: "styles.xml", content: stylesXml },
    { name: "meta.xml", content: metaXml },
    { name: "settings.xml", content: settingsXml },
    { name: "META-INF/manifest.xml", content: manifestXml },
    ...objectParts,
  ]);

  for (const outDir of [
    path.join(__dirname, "..", "e2e", "fixtures"),
    path.join(__dirname, "..", "src", "lib", "odf", "fixtures"),
  ]) {
    mkdirSync(outDir, { recursive: true });
    writeFileSync(path.join(outDir, "lo-chart-types.odt"), bytes);
    console.log("Wrote", path.join(outDir, "lo-chart-types.odt"));
  }
}
