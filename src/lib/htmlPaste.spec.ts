import { describe, it, expect } from "vitest";
import { importHtml } from "./htmlPaste";
import { blockText } from "./textEditing";

describe("importHtml — mapping", () => {
  it("maps headings from a plain web page (AC-16)", () => {
    const { blocks } = importHtml("<h1>One</h1><h3>Three</h3><p>body</p>");
    expect(blocks.map((block) => [blockText(block), block.headingLevel])).toEqual([
      ["One", 1],
      ["Three", 3],
      ["body", undefined],
    ]);
  });

  it("maps Word Online-style markup: wrapper divs, classes, styled spans (AC-14)", () => {
    const html =
      '<div class="OutlineElement Ltr SCXW1234"><p class="Paragraph SCXW1234" ' +
      'style="margin:0;text-align:center"><span class="TextRun SCXW1234" ' +
      'style="font-weight:bold;font-size:14pt;font-family:Calibri, sans-serif">' +
      '<span class="NormalTextRun">Hello</span></span></p></div>';
    const { blocks } = importHtml(html);
    expect(blocks).toHaveLength(1);
    expect(blocks[0].formatting?.alignment).toBe("center");
    expect(blocks[0].runs).toEqual([
      {
        text: "Hello",
        formatting: { bold: true, fontSizePt: 14, fontFamily: "Calibri" },
      },
    ]);
  });

  it("maps Google Docs-style markup: the font-weight:normal <b> wrapper and CSS toggles (AC-15)", () => {
    const html =
      '<meta charset="utf-8"><b style="font-weight:normal" id="docs-internal-guid-1">' +
      '<h2 dir="ltr"><span style="font-size:16pt;font-weight:400">Head</span></h2>' +
      '<p dir="ltr"><span style="font-weight:700">bold</span>' +
      '<span style="font-style:italic">italic</span>' +
      '<span style="text-decoration:underline">under</span>' +
      '<span style="text-decoration:line-through">struck</span></p></b>';
    const { blocks } = importHtml(html);
    expect(blocks).toHaveLength(2);
    expect(blocks[0].headingLevel).toBe(2);
    // The <b> wrapper's font-weight:normal must not bold everything.
    expect(blocks[0].runs[0].formatting.bold).toBeUndefined();
    expect(blocks[1].runs).toEqual([
      { text: "bold", formatting: { bold: true } },
      { text: "italic", formatting: { italic: true } },
      { text: "under", formatting: { underline: true } },
      { text: "struck", formatting: { strikethrough: true } },
    ]);
  });

  it("maps bold, italic and underline tags (AC-17..19)", () => {
    const { blocks } = importHtml(
      "<p><b>b</b><strong>s</strong><i>i</i><em>e</em><u>u</u></p>",
    );
    expect(blocks[0].runs).toEqual([
      { text: "bs", formatting: { bold: true } },
      { text: "ie", formatting: { italic: true } },
      { text: "u", formatting: { underline: true } },
    ]);
  });

  it("maps strikethrough, sub/superscript, colour, size and highlight", () => {
    const { blocks } = importHtml(
      '<p><s>gone</s><sub>lo</sub><sup>hi</sup>' +
        '<span style="color:rgb(255, 0, 0)">red</span>' +
        '<span style="background-color:#ffff00">marked</span>' +
        '<font color="#0000ff" face="Georgia" size="5">legacy</font></p>',
    );
    expect(blocks[0].runs).toEqual([
      { text: "gone", formatting: { strikethrough: true } },
      { text: "lo", formatting: { vertAlign: "subscript" } },
      { text: "hi", formatting: { vertAlign: "superscript" } },
      { text: "red", formatting: { color: "#ff0000" } },
      { text: "marked", formatting: { highlight: "yellow" } },
      { text: "legacy", formatting: { color: "#0000ff", fontFamily: "Georgia", fontSizePt: 18 } },
    ]);
  });

  it("converts px font sizes to half-point-snapped points", () => {
    const { blocks } = importHtml('<p><span style="font-size:16px">x</span></p>');
    expect(blocks[0].runs[0].formatting.fontSizePt).toBe(12);
  });

  it("imports list items as plain paragraphs until SPEC-018 (AC-20)", () => {
    const { blocks } = importHtml(
      "<ul><li>one</li><li><b>two</b></li><li>outer<ul><li>nested</li></ul></li></ul>" +
        "<ol><li>first</li></ol>",
    );
    expect(blocks.map(blockText)).toEqual(["one", "two", "outer", "nested", "first"]);
    expect(blocks[1].runs[0].formatting).toEqual({ bold: true });
  });

  it("keeps anchor runs and their formatting but drops the URL (AC-21)", () => {
    const { blocks } = importHtml('<p>see <a href="https://example.com"><b>the docs</b></a></p>');
    expect(blocks[0].runs).toEqual([
      { text: "see ", formatting: {} },
      { text: "the docs", formatting: { bold: true } },
    ]);
    expect(JSON.stringify(blocks)).not.toContain("example.com");
  });

  it("flattens tables to one tab-separated paragraph per row until SPEC-022 (AC-22)", () => {
    const { blocks } = importHtml(
      "<table><tbody><tr><td>a1</td><td><b>b1</b></td></tr>" +
        "<tr><td><p>a2</p></td><td>b2</td></tr></tbody></table>",
    );
    expect(blocks.map(blockText)).toEqual(["a1\tb1", "a2\tb2"]);
    // The tab separator coalesces with the unformatted cell; the bold cell
    // keeps its own run.
    expect(blocks[0].runs).toEqual([
      { text: "a1\t", formatting: {} },
      { text: "b1", formatting: { bold: true } },
    ]);
  });

  it("maps <br> to a forced line break inside the paragraph", () => {
    const { blocks } = importHtml("<p>line one<br>line two</p>");
    expect(blocks).toHaveLength(1);
    expect(blockText(blocks[0])).toBe("line one\nline two");
  });

  it("maps paragraph CSS to paragraph formatting per the table", () => {
    const { blocks } = importHtml(
      '<p style="text-align:right;margin-top:12pt;margin-bottom:8px;' +
        'line-height:1.5;text-indent:1.27cm;padding-left:36pt">x</p>',
    );
    expect(blocks[0].formatting).toEqual({
      alignment: "right",
      spaceBeforePt: 12,
      spaceAfterPt: 6,
      lineSpacing: 1.5,
      indentFirstLineCm: 1.27,
      indentStartCm: 1.27,
    });
  });

  it("keeps an empty <p> as an empty paragraph but not wrapper whitespace", () => {
    const { blocks } = importHtml("<div>\n  <p>a</p>\n  <p></p>\n  <p>b</p>\n</div>");
    expect(blocks.map(blockText)).toEqual(["a", "", "b"]);
  });

  it("collapses inter-tag whitespace without gluing words together", () => {
    const { blocks } = importHtml("<p>foo <b>bar</b>\n baz</p>");
    expect(blockText(blocks[0])).toBe("foo bar baz");
  });

  it("imports a bare inline fragment as a single paragraph", () => {
    const { blocks } = importHtml('plain <b>and bold</b>');
    expect(blocks).toHaveLength(1);
    expect(blocks[0].runs).toEqual([
      { text: "plain ", formatting: {} },
      { text: "and bold", formatting: { bold: true } },
    ]);
  });

  it("maps dir=rtl onto the paragraph direction", () => {
    const { blocks } = importHtml('<p dir="rtl">שלום</p>');
    expect(blocks[0].formatting).toEqual({ direction: "rtl" });
  });
});

describe("importHtml — sanitisation (AC-23..26, AC-39..44)", () => {
  it("survives the spec's hostile payload with only the safe text (the security test)", () => {
    const html =
      "<script>window.__xss = 1</script>" +
      '<style onclick="window.__xss = 2">p { color: red }</style>' +
      '<iframe src="https://evil.example/frame"></iframe>' +
      '<img src="https://evil.example/tracker.png">' +
      '<p onmouseover="window.__xss = 3">safe ' +
      '<a href="javascript:window.__xss = 4">link text</a></p>';
    const { blocks, imagesDropped } = importHtml(html);
    expect(blocks.map(blockText)).toEqual(["safe link text"]);
    expect(imagesDropped).toBe(true);
    const serialised = JSON.stringify(blocks);
    expect(serialised).not.toContain("script");
    expect(serialised).not.toContain("__xss");
    expect(serialised).not.toContain("iframe");
    expect(serialised).not.toContain("evil.example");
    expect(serialised).not.toContain("javascript:");
    expect(serialised).not.toContain("color: red");
  });

  it("drops <script> and <style> elements wholesale (AC-41, AC-42)", () => {
    const { blocks } = importHtml(
      "<p>a</p><script>alert(1)</script><style>p{}</style><p>b</p>",
    );
    expect(blocks.map(blockText)).toEqual(["a", "b"]);
  });

  it("drops embedded remote content carriers (AC-43)", () => {
    const { blocks } = importHtml(
      "<p>keep</p><iframe src=x></iframe><object data=x></object>" +
        "<embed src=x><video src=x></video><audio src=x></audio>" +
        '<form action=x><input value=v><button>go</button></form>' +
        "<svg><circle onload=alert(1)></circle></svg><link rel=stylesheet href=x>",
    );
    expect(blocks.map(blockText)).toEqual(["keep"]);
  });

  it("drops event handlers and javascript: URLs (AC-39, AC-40)", () => {
    const { blocks } = importHtml(
      '<p onclick="alert(1)" onmouseover="alert(2)">' +
        '<span onfocus="alert(3)">text</span>' +
        '<a href="JAVASCRIPT:alert(4)">go</a></p>',
    );
    expect(JSON.stringify(blocks)).not.toContain("alert");
    expect(blockText(blocks[0])).toBe("textgo");
  });

  it("drops classes and positioning styles (AC-24, AC-25)", () => {
    const { blocks } = importHtml(
      '<p class="MsoNormal" style="position:absolute;left:10px;top:20px;float:right;z-index:9">' +
        '<span class="fancy" style="position:fixed">x</span></p>',
    );
    expect(blocks[0].formatting).toBeUndefined();
    expect(blocks[0].runs).toEqual([{ text: "x", formatting: {} }]);
    expect(JSON.stringify(blocks)).not.toContain("MsoNormal");
  });

  it("never interprets style values that could smuggle a fetch", () => {
    const { blocks } = importHtml(
      '<p><span style="background-color:url(https://evil.example/x)">a</span>' +
        '<span style="color:expression(alert(1))">b</span>' +
        "<span style=\"font-family:url('x')\">c</span></p>",
    );
    // All three spans import unformatted, so they coalesce into one run.
    expect(blocks[0].runs).toEqual([{ text: "abc", formatting: {} }]);
  });

  it("drops unknown markup while keeping its characters (AC-26)", () => {
    const { blocks } = importHtml(
      "<p><abbr title='x'>HTML</abbr> and <custom-widget data-x='1'>widget text</custom-widget></p>",
    );
    expect(blocks[0].runs).toEqual([
      { text: "HTML and widget text", formatting: {} },
    ]);
  });

  it("flags images (remote and data:) for the unsupported-toast without importing them (AC-46)", () => {
    const remote = importHtml('<p>text</p><img src="https://example.com/x.png">');
    expect(remote.imagesDropped).toBe(true);
    expect(remote.blocks.map(blockText)).toEqual(["text"]);
    const inline = importHtml('<img src="data:image/png;base64,AAAA">');
    expect(inline.imagesDropped).toBe(true);
    expect(inline.blocks).toEqual([]);
  });
});
