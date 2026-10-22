// A minimal SVG path-data reader for the PDF exporter (SPEC-034).
//
// The PDF builder draws real vector paths for anything the page paints as
// SVG — today the connectors (SPEC-070). Their `d` strings come from this
// app's own writers (`routedPathToSvg`, `pointsToSvg`), so only the commands
// those emit are supported: move, line, horizontal/vertical line, cubic and
// close, in both absolute and relative spellings. Anything else makes the
// whole path unreadable and `parseSvgPath` returns null — the exporter then
// skips that path rather than drawing a mangled one.
//
// Pure and framework-free: no DOM, so it is unit-tested directly.

/** One point in the path's own user units. */
export interface SvgPoint {
  x: number;
  y: number;
}

/** One drawn segment: a straight line to `to`, or a cubic when the two
 * control points are present. */
export interface SvgSegment {
  to: SvgPoint;
  c1?: SvgPoint;
  c2?: SvgPoint;
}

/** One sub-path: where the pen went down, what it drew, and whether it was
 * closed back to the start. */
export interface SvgSubpath {
  start: SvgPoint;
  segments: SvgSegment[];
  closed: boolean;
}

/** Splits path data into `[command, ...numbers]` tokens. Numbers may be
 * separated by whitespace, commas, or nothing at all (`M0 0L1-1`). */
const TOKEN = /([MmLlHhVvCcZz])|(-?(?:\d*\.\d+|\d+)(?:[eE][-+]?\d+)?)/g;

/** Parses SVG path data into sub-paths, or null when it contains a command
 * this reader does not model (arcs, quadratics, smooth curves) or is
 * malformed. An empty or whitespace-only `d` parses as no sub-paths. */
export function parseSvgPath(d: string): SvgSubpath[] | null {
  const tokens: Array<string | number> = [];
  let consumed = 0;
  TOKEN.lastIndex = 0;
  for (let match = TOKEN.exec(d); match; match = TOKEN.exec(d)) {
    tokens.push(match[1] !== undefined ? match[1] : Number(match[2]));
    consumed += match[0].length;
  }
  // Anything the tokenizer had to skip (a stray letter, "1,,2") means the
  // data is not what this reader models.
  if (consumed !== d.replace(/[\s,]/g, "").length) {
    return null;
  }

  const subpaths: SvgSubpath[] = [];
  let current: SvgSubpath | null = null;
  let cursor: SvgPoint = { x: 0, y: 0 };
  let subpathStart: SvgPoint = { x: 0, y: 0 };
  let index = 0;
  let command = "";

  /** Reads `count` numbers, or returns null when they are not all there. */
  const numbers = (count: number): number[] | null => {
    const values: number[] = [];
    for (let i = 0; i < count; i += 1) {
      const value = tokens[index];
      if (typeof value !== "number") {
        return null;
      }
      values.push(value);
      index += 1;
    }
    return values;
  };

  while (index < tokens.length) {
    const token = tokens[index];
    if (typeof token === "string") {
      command = token;
      index += 1;
      if (command === "Z" || command === "z") {
        if (current) {
          current.closed = true;
          cursor = { ...subpathStart };
        }
        continue;
      }
    } else if (command === "") {
      return null; // numbers before any command
    } else if (command === "M") {
      command = "L"; // repeated move coordinates are implicit line-tos
    } else if (command === "m") {
      command = "l";
    }

    const relative = command === command.toLowerCase();
    const base = relative ? cursor : { x: 0, y: 0 };
    if (command === "M" || command === "m") {
      const values = numbers(2);
      if (!values) {
        return null;
      }
      cursor = { x: base.x + values[0], y: base.y + values[1] };
      subpathStart = { ...cursor };
      current = { start: { ...cursor }, segments: [], closed: false };
      subpaths.push(current);
      continue;
    }
    if (!current) {
      return null; // a draw command before any move
    }
    if (command === "L" || command === "l") {
      const values = numbers(2);
      if (!values) {
        return null;
      }
      cursor = { x: base.x + values[0], y: base.y + values[1] };
      current.segments.push({ to: { ...cursor } });
      continue;
    }
    if (command === "H" || command === "h") {
      const values = numbers(1);
      if (!values) {
        return null;
      }
      cursor = { x: base.x + values[0], y: cursor.y };
      current.segments.push({ to: { ...cursor } });
      continue;
    }
    if (command === "V" || command === "v") {
      const values = numbers(1);
      if (!values) {
        return null;
      }
      cursor = { x: cursor.x, y: base.y + values[0] };
      current.segments.push({ to: { ...cursor } });
      continue;
    }
    if (command === "C" || command === "c") {
      const values = numbers(6);
      if (!values) {
        return null;
      }
      const c1 = { x: base.x + values[0], y: base.y + values[1] };
      const c2 = { x: base.x + values[2], y: base.y + values[3] };
      cursor = { x: base.x + values[4], y: base.y + values[5] };
      current.segments.push({ to: { ...cursor }, c1, c2 });
      continue;
    }
    return null;
  }

  return subpaths.filter((subpath) => subpath.segments.length > 0 || subpath.closed);
}
