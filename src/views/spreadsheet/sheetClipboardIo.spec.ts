// SPEC-131 "Paste": a clipboard image inserts as a drawing only when the
// text/html flavour (if any) holds no table and no internal payload — a
// screenshot inserts, a range copied from Excel (HTML table + PNG) pastes
// cells.
import { describe, expect, it } from "vitest";
import { PAYLOAD_ATTRIBUTE } from "../../lib/spreadsheet/clipboard";
import { pastedImage } from "./sheetClipboardIo";

function clipboard(html: string, files: File[]) {
  return {
    getData: (type: string) => (type === "text/html" ? html : ""),
    items: files.map((file) => ({ kind: "file", type: file.type, getAsFile: () => file })) as unknown as DataTransferItemList,
    files: files as unknown as FileList,
  };
}

describe("SPEC-131 pastedImage", () => {
  const png = new File([new Uint8Array([1, 2, 3])], "shot.png", { type: "image/png" });

  it("answers an accepted image when there is no HTML, or HTML without a table", () => {
    expect(pastedImage(clipboard("", [png]))).toBe(png);
    expect(pastedImage(clipboard('<img src="x.png">', [png]))).toBe(png);
  });

  it("answers null for a copied range (an HTML table, or this view's own payload), an unsupported type or no file", () => {
    expect(pastedImage(clipboard("<table><tr><td>1</td></tr></table>", [png]))).toBeNull();
    expect(pastedImage(clipboard(`<div ${PAYLOAD_ATTRIBUTE}="{}"></div>`, [png]))).toBeNull();
    expect(pastedImage(clipboard("", [new File([""], "a.tiff", { type: "image/tiff" })]))).toBeNull();
    expect(pastedImage(clipboard("", []))).toBeNull();
  });
});
