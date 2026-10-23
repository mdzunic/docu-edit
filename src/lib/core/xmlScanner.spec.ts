// SPEC-116: the forward-only XML scanner — events in document order with
// exact source offsets, namespace resolution (default and prefixed, rebinding
// in scope), entities, CDATA, "stop", the DTD refusal, and malformed input
// reported with its offset. SPEC-137: a scan that pauses and resumes, and a
// scan of one element cut out of a larger text by offset.
import { describe, expect, it } from "vitest";
import { scanAttr, scanXml, startXmlScan, XmlScanError, type XmlScanEvent } from "./xmlScanner";

function events(text: string): XmlScanEvent[] {
  const out: XmlScanEvent[] = [];
  scanXml(text, (event) => {
    out.push(event);
  });
  return out;
}

/** The error a scan throws (failing the test if it throws nothing). */
function scanError(text: string): XmlScanError {
  try {
    scanXml(text, () => {});
  } catch (error) {
    if (error instanceof XmlScanError) {
      return error;
    }
    throw error;
  }
  throw new Error("expected an XmlScanError");
}

describe("scanXml — events and offsets", () => {
  it("reports open, text and close in document order", () => {
    const xml = `<a x="1"><b>hi</b><c/></a>`;
    const seen = events(xml).map((event) =>
      event.type === "text" ? `text:${event.text}` : `${event.type}:${event.name}${event.type === "open" && event.selfClosing ? "/" : ""}`,
    );
    expect(seen).toEqual(["open:a", "open:b", "text:hi", "close:b", "open:c/", "close:a"]);
  });

  it("gives every event the exact UTF-16 span of its source markup", () => {
    const xml = `<?xml version="1.0"?>\n<root a='q'>café &amp; \u{1F600}<leaf k="v"/></root>`;
    for (const event of events(xml)) {
      const raw = xml.slice(event.start, event.end);
      if (event.type === "open") {
        expect(raw.startsWith(`<${event.name}`)).toBe(true);
        expect(raw.endsWith(event.selfClosing ? "/>" : ">")).toBe(true);
      } else if (event.type === "close") {
        expect(raw).toBe(`</${event.name}>`);
      } else {
        expect(raw).toBe("café &amp; \u{1F600}");
        expect(event.text).toBe("café & \u{1F600}");
      }
    }
  });

  it("reports a self-closing element as one open event and no close", () => {
    const seen = events(`<r><e a="1"/></r>`);
    expect(seen.filter((event) => event.type === "close").map((event) => event.type === "close" && event.name)).toEqual(["r"]);
    const open = seen.find((event) => event.type === "open" && event.name === "e");
    expect(open?.type === "open" && open.selfClosing).toBe(true);
  });

  it("skips comments, processing instructions and the XML declaration", () => {
    const seen = events(`<?xml version="1.0" encoding="UTF-8"?><!-- lead --><r><?pi data?>a<!-- c -->b</r><!-- tail -->`);
    expect(seen.map((event) => (event.type === "text" ? event.text : event.type))).toEqual(["open", "a", "b", "close"]);
  });

  it("skips a leading byte-order mark", () => {
    expect(events("﻿<r/>")).toHaveLength(1);
  });

  it("emits whitespace-only text inside elements (offsets matter to slicers)", () => {
    const seen = events(`<r>\n  <c/>\n</r>`);
    expect(seen.filter((event) => event.type === "text").map((event) => event.type === "text" && event.text)).toEqual([
      "\n  ",
      "\n",
    ]);
  });
});

describe("scanXml — namespaces", () => {
  it("resolves the default namespace and prefixed names from in-scope declarations", () => {
    const xml =
      `<w:workbook xmlns:w="urn:main" xmlns:r="urn:rels" xmlns="urn:default">` +
      `<w:sheet r:id="rId1" name="Data"/><plain/></w:workbook>`;
    const opens = events(xml).filter((event) => event.type === "open");
    expect(opens.map((event) => event.type === "open" && [event.local, event.ns])).toEqual([
      ["workbook", "urn:main"],
      ["sheet", "urn:main"],
      ["plain", "urn:default"],
    ]);
    const sheet = opens[1]!;
    if (sheet.type !== "open") {
      throw new Error("unreachable");
    }
    // Unprefixed attributes have no namespace; prefixed ones resolve.
    expect(scanAttr(sheet.attrs, "name")).toBe("Data");
    expect(scanAttr(sheet.attrs, "id", "urn:rels")).toBe("rId1");
    expect(scanAttr(sheet.attrs, "id")).toBeUndefined();
  });

  it("rebinds a prefix (and the default) for one subtree only", () => {
    const xml = `<p:a xmlns:p="urn:one" xmlns="urn:d1"><p:b xmlns:p="urn:two" xmlns="urn:d2"><c/></p:b><p:d/><e/></p:a>`;
    const opens = events(xml).filter((event) => event.type === "open");
    expect(opens.map((event) => event.type === "open" && `${event.local}@${event.ns}`)).toEqual([
      "a@urn:one",
      "b@urn:two",
      "c@urn:d2",
      "d@urn:one",
      "e@urn:d1",
    ]);
    const closes = events(xml).filter((event) => event.type === "close");
    expect(closes.map((event) => event.type === "close" && `${event.local}@${event.ns}`)).toEqual([
      "b@urn:two",
      "a@urn:one",
    ]);
  });

  it("undeclares the default namespace with xmlns=\"\" and predeclares xml:", () => {
    const opens = events(`<a xmlns="urn:x"><b xmlns=""><c xml:space="preserve"/></b></a>`).filter(
      (event) => event.type === "open",
    );
    expect(opens.map((event) => event.type === "open" && event.ns)).toEqual(["urn:x", undefined, undefined]);
    const c = opens[2]!;
    expect(c.type === "open" && scanAttr(c.attrs, "space", "http://www.w3.org/XML/1998/namespace")).toBe("preserve");
  });

  it("leaves an unbound prefix unresolved rather than failing", () => {
    const open = events(`<x:a/>`)[0]!;
    expect(open.type === "open" && [open.local, open.ns]).toEqual(["a", undefined]);
  });
});

describe("scanXml — entities, CDATA and normalisation", () => {
  it("decodes the five predefined entities and numeric references in text and attributes", () => {
    const seen = events(`<r v="&lt;&#65;&#x42;&quot;&apos;">&lt;&gt;&amp;&quot;&apos;&#169;&#x1F600;</r>`);
    const open = seen[0]!;
    expect(open.type === "open" && scanAttr(open.attrs, "v")).toBe(`<AB"'`);
    expect(seen[1]!.type === "text" && seen[1]!.text).toBe(`<>&"'©\u{1F600}`);
  });

  it("unwraps CDATA literally, its span covering the whole section", () => {
    const xml = `<r>a<![CDATA[<b>&amp;</b>]]>c</r>`;
    const texts = events(xml).filter((event) => event.type === "text");
    expect(texts.map((event) => event.type === "text" && event.text)).toEqual(["a", "<b>&amp;</b>", "c"]);
    const cdata = texts[1]!;
    expect(xml.slice(cdata.start, cdata.end)).toBe("<![CDATA[<b>&amp;</b>]]>");
  });

  it("normalises line ends in text and whitespace in attribute values", () => {
    const seen = events(`<r a="x\ty\r\nz&#10;w">one\r\ntwo\rthree</r>`);
    const open = seen[0]!;
    expect(open.type === "open" && scanAttr(open.attrs, "a")).toBe("x y z\nw");
    expect(seen[1]!.type === "text" && seen[1]!.text).toBe("one\ntwo\nthree");
  });

  it("refuses an unknown entity, with its offset", () => {
    const xml = `<r>ok &nbsp; no</r>`;
    const error = scanError(xml);
    expect(error.offset).toBe(xml.indexOf("&nbsp;"));
    expect(error.message).toContain("&nbsp;");
  });

  it("refuses a character reference to a non-XML character", () => {
    expect(scanError(`<r>&#0;</r>`).offset).toBe(3);
  });
});

describe("scanXml — stop", () => {
  it("ends the scan as soon as the callback returns \"stop\"", () => {
    const seen: string[] = [];
    // Malformed after the third cell: never reached, so never reported.
    scanXml(`<r><c>1</c><c>2</c><c>3</c><c>4</c></bad>`, (event) => {
      if (event.type === "open" && event.name === "c") {
        seen.push("c");
        if (seen.length === 3) {
          return "stop";
        }
      }
      return undefined;
    });
    expect(seen).toEqual(["c", "c", "c"]);
  });
});

describe("scanXml — refusals and malformed input", () => {
  it("refuses a DTD rather than expanding its entities", () => {
    const xml = `<?xml version="1.0"?>\n<!DOCTYPE r [<!ENTITY boom "x">]><r>&boom;</r>`;
    const error = scanError(xml);
    expect(error.offset).toBe(xml.indexOf("<!DOCTYPE"));
    expect(error.message).toMatch(/DTD/);
  });

  it("reports a mismatched closing tag at the tag", () => {
    const xml = `<a><b></a></b>`;
    expect(scanError(xml).offset).toBe(xml.indexOf("</a>"));
  });

  it("reports an element that is never closed at its open tag", () => {
    const xml = `<a><b>text</b><c>`;
    expect(scanError(xml).offset).toBe(xml.indexOf("<c>"));
  });

  it("reports unquoted, unterminated and duplicate attributes", () => {
    expect(scanError(`<a x=1/>`).offset).toBe(5);
    expect(scanError(`<a x="1/>`).offset).toBe(5);
    const duplicate = `<a x="1" x="2"/>`;
    expect(scanError(duplicate).offset).toBe(duplicate.lastIndexOf("x="));
  });

  it("reports '<' inside an attribute value", () => {
    const xml = `<a x="1<2"/>`;
    expect(scanError(xml).offset).toBe(xml.indexOf("<2"));
  });

  it("reports text and elements outside the root", () => {
    expect(scanError(`<a/>junk`).offset).toBe(4);
    expect(scanError(`<a/><b/>`).offset).toBe(4);
    expect(scanError(`junk<a/>`).offset).toBe(0);
  });

  it("reports unterminated comments, CDATA and processing instructions", () => {
    expect(scanError(`<a><!-- open</a>`).offset).toBe(3);
    expect(scanError(`<a><![CDATA[ open</a>`).offset).toBe(3);
    expect(scanError(`<?xml version="1.0"<a/>`).offset).toBe(0);
  });

  it("refuses an empty document", () => {
    expect(scanError("   ").offset).toBe(3);
  });
});

describe("pausing and ranges (SPEC-137)", () => {
  const doc = `<?xml version="1.0"?><r xmlns="urn:a" xmlns:p="urn:p"><row n="1"><c>1</c></row><row n="2"><c>2</c></row><p:x/></r>`;

  it("resumes after a pause exactly where it left off, with the same events as one pass", () => {
    const all = events(doc);
    const seen: XmlScanEvent[] = [];
    const scan = startXmlScan(doc, (event) => {
      seen.push(event);
      return event.type === "close" && event.local === "row" ? "pause" : undefined;
    });
    let resumes = 0;
    while (!scan.resume()) {
      resumes += 1;
    }
    expect(resumes).toBe(2);
    expect(seen).toEqual(all);
    expect(scan.resume()).toBe(true);
  });

  it("ends at a stop even after pauses", () => {
    const seen: string[] = [];
    const scan = startXmlScan(doc, (event) => {
      if (event.type !== "open") {
        return undefined;
      }
      seen.push(event.local);
      return event.local === "row" ? (seen.length > 3 ? "stop" : "pause") : undefined;
    });
    expect(scan.resume()).toBe(false);
    expect(scan.resume()).toBe(true);
    expect(seen).toEqual(["r", "row", "c", "row"]);
  });

  it("scans one element of a larger text by offset, in the text's coordinates, with the scope given", () => {
    const start = doc.indexOf(`<row n="2">`);
    const end = doc.indexOf("<p:x/>");
    const out: XmlScanEvent[] = [];
    scanXml(
      doc,
      (event) => {
        out.push(event);
      },
      { from: start, to: end, scope: new Map([["", "urn:a"], ["p", "urn:p"]]) },
    );
    expect(out.map((event) => event.type)).toEqual(["open", "open", "text", "close", "close"]);
    expect(out[0]).toMatchObject({ local: "row", ns: "urn:a", start, attrs: [{ name: "n", value: "2" }] });
    expect(out.at(-1)).toMatchObject({ type: "close", local: "row", end });
    expect(doc.slice(out[2]!.start, out[2]!.end)).toBe("2");
  });

  it("refuses a range that does not hold one whole element", () => {
    const start = doc.indexOf(`<row n="2">`);
    expect(() => scanXml(doc, () => {}, { from: start, to: start + 15, scope: new Map([["", "urn:a"]]) })).toThrow(XmlScanError);
  });
});
