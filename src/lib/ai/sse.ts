// An incremental Server-Sent Events parser (SPEC-113) for the gateway's
// streaming form. Text arrives in whatever chunks the network delivers — an
// event split mid-line, a CR at the end of one chunk and its LF at the start
// of the next, several events in one chunk — and this assembles them into
// complete events per the WHATWG event-stream rules: lines end at CRLF, LF or
// CR; a blank line dispatches; a line starting with ":" is a comment; a
// "data" field's value (after one optional space) is joined to the event's
// earlier data lines with "\n". Only `data` matters to the gateway; `event`,
// `id` and `retry` are parsed past and ignored.
//
// It parses the stream's framing, never its content: each event's data is
// handed over as the string it is.

export interface SseParser {
  /** Feeds the next decoded chunk of the stream. */
  push(chunk: string): void;
  /** Ends the stream. A final event left without its blank line is not
   * dispatched — per the rules, an incomplete event is discarded. */
  end(): void;
}

export function createSseParser(onData: (data: string) => void): SseParser {
  let line = "";
  let data: string[] = [];
  /** A CR ended the last chunk; an LF opening the next one is its pair. */
  let pendingCr = false;

  function dispatch(): void {
    if (data.length > 0) {
      const joined = data.join("\n");
      data = [];
      onData(joined);
    }
  }

  function processLine(text: string): void {
    if (text === "") {
      dispatch();
      return;
    }
    if (text.startsWith(":")) {
      return;
    }
    const colon = text.indexOf(":");
    const field = colon === -1 ? text : text.slice(0, colon);
    let value = colon === -1 ? "" : text.slice(colon + 1);
    if (value.startsWith(" ")) {
      value = value.slice(1);
    }
    if (field === "data") {
      data.push(value);
    }
  }

  return {
    push(chunk: string): void {
      let start = 0;
      for (let i = 0; i < chunk.length; i++) {
        const ch = chunk[i];
        if (ch === "\n" && pendingCr && i === 0) {
          // The LF half of a CRLF split across chunks: already ended.
          pendingCr = false;
          start = 1;
          continue;
        }
        pendingCr = false;
        if (ch === "\r" || ch === "\n") {
          processLine(line + chunk.slice(start, i));
          line = "";
          if (ch === "\r") {
            if (chunk[i + 1] === "\n") {
              i++;
            } else if (i === chunk.length - 1) {
              pendingCr = true;
            }
          }
          start = i + 1;
        }
      }
      line += chunk.slice(start);
    },
    end(): void {
      line = "";
      data = [];
      pendingCr = false;
    },
  };
}
