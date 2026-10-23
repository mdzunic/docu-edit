import { appendFileSync } from "node:fs";
const LOG = (...a: unknown[]) => appendFileSync("/tmp/rev/out.txt", a.map(String).join(" ") + "\n");
import { describe, it, expect } from "vitest";
import { parseSlideMarkdownWithReport, writeSlideMarkdown, parseSlideMarkdown } from "./slideMarkdown";
import { equalStructural, firstStructuralDifference } from "../equality";
import { PresentationBuilder } from "./presentation";
import { unzip } from "../core/zip";
import type { PresentationModel } from "../presentation/model";

function strip(deck: PresentationModel): unknown {
  return JSON.parse(JSON.stringify(deck, (k, v) => (k === "id" ? undefined : v)));
}
function rt(md: string) {
  const first = parseSlideMarkdownWithReport(md);
  const written = writeSlideMarkdown(first.deck);
  const second = parseSlideMarkdownWithReport(written.text);
  const diff = firstStructuralDifference(strip(second.deck), strip(first.deck));
  return { first, written, second, diff, eq: equalStructural(first.deck, second.deck) };
}
const PNG = "iVBORw0KGgoAAAANSUhEUgAAAAQAAAACCAAAAABawyK/AAAAC0lEQVR4nGNggAEAAAoAAX+AdF4AAAAASUVORK5CYII=";

describe("scratch", () => {
  const cases: Record<string, string> = {
    hashPara: "## T\n\n\\#",
    hashParaSlide0: "\\#\n",
    bangLink: "## T\n\n!*[a](http://x.com)*",
    leadingSpace: "## T\n\n\\ \\ a",
    wsTitle: "# ** **\n",
    listCont: "## T\n- a\n  continued",
    tabList: "## T\n\t- a",
    imgInList: `## T\n- ![](data:image/png;base64,${PNG})`,
    notesTrailingHash: "## T\n\nNotes: a\n  b",
    numberedEsc: "## T\n\n1\\. foo",
    notesEsc: "## T\n\n Notes: x",
    twoLinks: "## T\n\n[a](http://x)[b](http://y)",
    emptyBoldTitlePara: "## T\n\n** **",
    trailingBackslashEscapes: "## T\n\nC:\\\\",
    sep4: "## A\n\n---\n\n---\n\n## B",
    pipeLink: "## T\n\n[a|b](http://x)",
    codeInTitle: "## `a#`",
    titleHashes: "## a \\#",
    frontQuote: "---\ntitle: 'It''s'\n---\n# x",
    setext: "## T\n\nHello\n===",
  };
  for (const [name, md] of Object.entries(cases)) {
    it(name, () => {
      const r = rt(md);
      LOG(`=== ${name}\nIN: ${JSON.stringify(md)}\nOUT: ${JSON.stringify(r.written.text)}\ndropped1: ${JSON.stringify(r.first.dropped)}\ndropped2: ${JSON.stringify(r.second.dropped)}\ndiff: ${JSON.stringify(r.diff)} eq=${r.eq}\nslides1: ${JSON.stringify(r.first.deck.slides.map(s=>({f:s.frames.map(f=>f.kind==='text'?{role:f.role,p:f.paragraphs}:f.kind),n:s.notes})))}\nslides2: ${JSON.stringify(r.second.deck.slides.map(s=>({f:s.frames.map(f=>f.kind==='text'?{role:f.role,p:f.paragraphs}:f.kind),n:s.notes})))}`);
    });
  }
  it("html swallow notes", () => {
    const r = parseSlideMarkdownWithReport("## T\n<div>x</div>\nNotes: hello");
    LOG("HTML:", JSON.stringify(r.deck.slides[0].notes), JSON.stringify(r.dropped));
    const r2 = parseSlideMarkdownWithReport("# A\n```\ncode\n---\n## B\n---\n## C");
    LOG("FENCE:", r2.deck.slides.length, JSON.stringify(r2.dropped));
    const r3 = parseSlideMarkdownWithReport("## T\n> quote\nNotes: hi");
    LOG("BQ:", JSON.stringify(r3.deck.slides[0].notes), JSON.stringify(r3.dropped));
    const r4 = parseSlideMarkdownWithReport("## T\na | b\n--|--\nNotes: x|y");
    LOG("TABLE:", JSON.stringify(r4.deck.slides[0].notes), JSON.stringify(r4.dropped));
  });
  it("builder newline", () => {
    const deck = new PresentationBuilder().addBulletSlide("T", ["a\nb"]).setNotes(["  indented"]).toDeck();
    const w = writeSlideMarkdown(deck);
    const back = parseSlideMarkdown(w.text);
    LOG("NL:", JSON.stringify(w.text), JSON.stringify((back.slides[0].frames[1] as any).paragraphs), JSON.stringify(back.slides[0].notes));
  });
  it("cross-format media", async () => {
    const b = new PresentationBuilder().addImageSlide("T", { data: PNG, contentType: "image/png" });
    const parts = await unzip(await b.toBytes("odp"));
    LOG("ODP parts:", [...parts.keys()].join(","));
    const b2 = new PresentationBuilder({format: "odp"}).addImageSlide("T", { data: PNG, contentType: "image/png" });
    const parts2 = await unzip(await b2.toBytes("pptx"));
    LOG("PPTX parts:", [...parts2.keys()].join(","));
  });
});
describe("scratch2", () => {
  it("manifest", async () => {
    const b = new PresentationBuilder().addImageSlide("T", { data: PNG, contentType: "image/png" });
    const parts = await unzip(await b.toBytes("odp"));
    LOG("MANIFEST:", new TextDecoder().decode(parts.get("META-INF/manifest.xml")));
    const b2 = new PresentationBuilder({format: "odp"}).addImageSlide("T", { data: PNG, contentType: "image/png" });
    const parts2 = await unzip(await b2.toBytes("pptx"));
    LOG("CT:", new TextDecoder().decode(parts2.get("[Content_Types].xml")));
    LOG("RELS:", new TextDecoder().decode(parts2.get("ppt/slides/_rels/slide1.xml.rels")));
  });
});
