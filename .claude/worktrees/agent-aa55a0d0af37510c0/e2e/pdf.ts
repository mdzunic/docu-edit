// Reading an exported PDF in an e2e assertion.
//
// The writer deflates its page content streams and its embedded font
// programs (see buildPdf.ts), so the text-showing operators, the vector fill
// and stroke operators and the sfnt bytes are all inside /FlateDecode
// streams. A test that wants to assert on any of those has to expand them
// first; one asserting only on object dictionaries (/Count, /Type /Page,
// /BaseFont) can read the file as latin-1 directly.
import { readFileSync } from "node:fs";
import { inflateSync } from "node:zlib";

/** The file as latin-1 text with every /FlateDecode stream expanded in
 * place, so substring assertions see the operators rather than zlib bytes.
 * A stream that does not actually inflate (an image passed through by a
 * different filter chain) is left exactly as it is. */
export function pdfTextOf(bytes: Buffer): string {
  const raw = bytes.toString("latin1");
  const header = /<<([^>]*)\/Length (\d+) >>\nstream\n/g;
  let out = "";
  let cursor = 0;
  for (let match = header.exec(raw); match !== null; match = header.exec(raw)) {
    const start = match.index + match[0].length;
    const length = Number(match[2]);
    out += raw.slice(cursor, start);
    const body = bytes.subarray(start, start + length);
    let expanded: string | null = null;
    if (match[1].includes("/Filter /FlateDecode")) {
      try {
        expanded = inflateSync(body).toString("latin1");
      } catch {
        expanded = null;
      }
    }
    out += expanded ?? body.toString("latin1");
    cursor = start + length;
  }
  return out + raw.slice(cursor);
}

/** `pdfTextOf` for a file on disk — what a Playwright download hands back. */
export function readPdfText(path: string): string {
  return pdfTextOf(readFileSync(path));
}
