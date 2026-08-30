// detectMedia (SPEC-107 AC-6): every signature, plus the SVG
// false-positive cases the spec calls out — an XHTML file, an XML file whose
// COMMENT mentions svg, and an empty file. The sniff runs on bytes alone;
// no test here involves a file name.
import { describe, expect, it } from "vitest";
import { detectMedia } from "../core/packageDetect";
import {
  buildGifFixtureBytes,
  buildJpegFixtureBytes,
  buildPagesPdfBytes,
  buildPngFixtureBytes,
  buildSvgFixtureBytes,
  buildWebpFixtureBytes,
} from "./fixtures/build";

function utf8(text: string): Uint8Array {
  return new TextEncoder().encode(text);
}

describe("detectMedia", () => {
  it("recognises all six formats from their bytes", async () => {
    expect(detectMedia(buildPagesPdfBytes())).toBe("pdf");
    expect(detectMedia(await buildPngFixtureBytes())).toBe("png");
    expect(detectMedia(buildJpegFixtureBytes())).toBe("jpeg");
    expect(detectMedia(buildGifFixtureBytes())).toBe("gif");
    expect(detectMedia(buildWebpFixtureBytes())).toBe("webp");
    expect(detectMedia(buildSvgFixtureBytes())).toBe("svg");
  });

  it("finds a %PDF- header within the first 1024 bytes, not beyond", () => {
    const prefixed = new Uint8Array(100 + 20);
    prefixed.set(utf8("%PDF-1.4"), 100);
    expect(detectMedia(prefixed)).toBe("pdf");
    const buried = new Uint8Array(2000 + 20);
    buried.set(utf8("%PDF-1.4"), 1500);
    expect(detectMedia(buried)).toBeNull();
  });

  it("accepts SVG behind a BOM, XML declaration, doctype and comments", () => {
    const svg =
      "﻿<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n" +
      "<!-- a leading comment -->\n<!DOCTYPE svg>\n<!-- another -->\n" +
      "<svg xmlns=\"http://www.w3.org/2000/svg\"/>";
    expect(detectMedia(utf8(svg))).toBe("svg");
  });

  it("accepts a namespace-prefixed svg root", () => {
    expect(detectMedia(utf8('<s:svg xmlns:s="http://www.w3.org/2000/svg"/>'))).toBe("svg");
  });

  it("rejects an XHTML file", () => {
    const xhtml =
      '<?xml version="1.0"?><!DOCTYPE html><html xmlns="http://www.w3.org/1999/xhtml"><body/></html>';
    expect(detectMedia(utf8(xhtml))).toBeNull();
  });

  it("rejects an XML file whose comment mentions svg", () => {
    const xml = '<?xml version="1.0"?><!-- contains <svg> in a comment --><data/>';
    expect(detectMedia(utf8(xml))).toBeNull();
  });

  it("rejects an element merely starting with the letters svg", () => {
    expect(detectMedia(utf8("<svgfoo/>"))).toBeNull();
  });

  it("rejects an empty file and plain text", () => {
    expect(detectMedia(new Uint8Array(0))).toBeNull();
    expect(detectMedia(utf8("hello world"))).toBeNull();
  });

  it("rejects ZIP-shaped bytes (packages take the package path)", () => {
    expect(detectMedia(new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0, 0, 0, 0]))).toBeNull();
  });
});
