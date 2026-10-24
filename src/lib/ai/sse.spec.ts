// SPEC-113 — the incremental event-stream parser the gateway's stream() uses:
// events assembled from arbitrary chunk boundaries per the WHATWG rules.
import { describe, expect, it } from "vitest";
import { createSseParser } from "./sse";

function parse(chunks: string[]): string[] {
  const events: string[] = [];
  const parser = createSseParser((data) => events.push(data));
  for (const chunk of chunks) {
    parser.push(chunk);
  }
  parser.end();
  return events;
}

describe("createSseParser", () => {
  it("dispatches one event per blank line", () => {
    expect(parse(["data: a\n\ndata: b\n\n"])).toEqual(["a", "b"]);
  });

  it("assembles an event split anywhere — mid-field, mid-value, between line ends", () => {
    const stream = "data: first\n\ndata: second\n\n";
    for (let cut = 1; cut < stream.length; cut++) {
      expect(parse([stream.slice(0, cut), stream.slice(cut)])).toEqual(["first", "second"]);
    }
    expect(parse([...stream])).toEqual(["first", "second"]);
  });

  it("treats CRLF, LF and CR alike, including a CRLF split across two chunks", () => {
    expect(parse(["data: a\r\n\r\ndata: b\r\rdata: c\n\n"])).toEqual(["a", "b", "c"]);
    expect(parse(["data: x\r", "\ndata: y\r", "\n\r", "\n"])).toEqual(["x\ny"]);
  });

  it("joins several data lines with a newline", () => {
    expect(parse(["data: one\ndata: two\ndata:three\n\n"])).toEqual(["one\ntwo\nthree"]);
  });

  it("skips comments and ignores event, id and retry fields", () => {
    expect(parse([": ping\n\nevent: delta\nid: 7\nretry: 100\ndata: z\n\n"])).toEqual(["z"]);
  });

  it("strips exactly one space after the colon", () => {
    expect(parse(["data:  two spaces\n\n"])).toEqual([" two spaces"]);
  });

  it("does not dispatch an event with no data, or one left without its blank line", () => {
    expect(parse(["event: nothing\n\n", "data: unfinished"])).toEqual([]);
  });
});
