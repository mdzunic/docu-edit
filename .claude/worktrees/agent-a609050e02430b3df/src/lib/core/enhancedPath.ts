// SPEC-071: the ODF `draw:enhanced-path` command string <-> `ShapePath`
// bridge. Product-neutral and framework-free: pure string in, pure model
// out. The model carries only move/line/cubic/close, so the reader converts
// ODF's quadratic (`Q`) and arc (`A`/`T`) commands to cubics here at the
// boundary (SPEC-071 "Primitive reduction"); a command we do not model — or
// any non-numeric token, which is what a formula-driven (`draw:equation`)
// path contains — makes the parse return null so the caller keeps the shape
// on SPEC-051's verbatim preservation path (AC-31).
import { arcToCubics, PATH_GRID, quadraticToCubic, type PathStep, type ShapeSubpath, type ShapePath } from "./shapePath";

/** Splits a command string into command letters and numeric tokens. A
 * formula-driven path carries `?f0`/`$0` references, which surface here as
 * tokens that are neither a known command nor a finite number. */
function tokenize(d: string): string[] | null {
  const tokens = d.match(/[A-Za-z]|-?\d*\.?\d+(?:[eE][-+]?\d+)?/g);
  return tokens && tokens.length > 0 ? tokens : null;
}

/** Parses an ODF `draw:enhanced-path` (absolute M/L/C/Z/N plus Q/A/T reduced
 * to cubics) against its `svg:viewBox` extent into a `ShapePath` in 0..1
 * space, or null when it is formula-driven or uses an unmodelled command.
 * `converted` is true when any quadratic/arc was reduced, so the caller can
 * preserve the source for a lossless same-format save. */
export function parseEnhancedPath(
  d: string,
  viewBoxW: number,
  viewBoxH: number,
): { path: ShapePath; converted: boolean } | null {
  const tokens = tokenize(d);
  if (!tokens || viewBoxW <= 0 || viewBoxH <= 0) {
    return null;
  }
  let i = 0;
  const subpaths: ShapeSubpath[] = [];
  let steps: PathStep[] = [];
  let closed = false;
  let converted = false;
  let cur = { x: 0, y: 0 };

  const flush = (): void => {
    if (steps.length > 0) {
      subpaths.push({ steps, closed });
    }
    steps = [];
    closed = false;
  };

  /** Reads `count` numbers as viewBox-fraction points (x/y interleaved), or
   * null when a token is missing or non-numeric (formula-driven). */
  const readPoints = (count: number): { x: number; y: number }[] | null => {
    const out: { x: number; y: number }[] = [];
    for (let p = 0; p < count; p++) {
      const xt = tokens[i++];
      const yt = tokens[i++];
      const x = Number(xt);
      const y = Number(yt);
      if (xt === undefined || yt === undefined || !Number.isFinite(x) || !Number.isFinite(y)) {
        return null;
      }
      out.push({ x: x / viewBoxW, y: y / viewBoxH });
    }
    return out;
  };
  const readNums = (count: number): number[] | null => {
    const out: number[] = [];
    for (let p = 0; p < count; p++) {
      const t = tokens[i++];
      const n = Number(t);
      if (t === undefined || !Number.isFinite(n)) {
        return null;
      }
      out.push(n);
    }
    return out;
  };

  while (i < tokens.length) {
    const cmd = tokens[i++];
    if (cmd === "M") {
      flush();
      const p = readPoints(1);
      if (!p) {
        return null;
      }
      steps.push({ op: "move", x: p[0].x, y: p[0].y });
      cur = p[0];
    } else if (cmd === "L") {
      // L accepts one or more implicit points.
      do {
        const p = readPoints(1);
        if (!p) {
          return null;
        }
        steps.push({ op: "line", x: p[0].x, y: p[0].y });
        cur = p[0];
      } while (i < tokens.length && Number.isFinite(Number(tokens[i])));
    } else if (cmd === "C") {
      do {
        const p = readPoints(3);
        if (!p) {
          return null;
        }
        steps.push({ op: "cubic", x1: p[0].x, y1: p[0].y, x2: p[1].x, y2: p[1].y, x: p[2].x, y: p[2].y });
        cur = p[2];
      } while (i < tokens.length && Number.isFinite(Number(tokens[i])));
    } else if (cmd === "Q") {
      const p = readPoints(2);
      if (!p) {
        return null;
      }
      steps.push({ op: "cubic", ...quadraticToCubic(cur, p[0], p[1]) });
      cur = p[1];
      converted = true;
    } else if (cmd === "A") {
      // arcto: (x1 y1) (x2 y2) bounding box, (x3 y3) start, (x4 y4) end.
      const p = readPoints(4);
      if (!p) {
        return null;
      }
      const cx = (p[0].x + p[1].x) / 2;
      const cy = (p[0].y + p[1].y) / 2;
      const rx = Math.abs(p[1].x - p[0].x) / 2;
      const ry = Math.abs(p[1].y - p[0].y) / 2;
      if (rx === 0 || ry === 0) {
        return null;
      }
      const start = Math.atan2((p[2].y - cy) / ry, (p[2].x - cx) / rx);
      let sweep = Math.atan2((p[3].y - cy) / ry, (p[3].x - cx) / rx) - start;
      if (sweep <= 0) {
        sweep += 2 * Math.PI;
      }
      for (const seg of arcToCubics(cx, cy, rx, ry, start, sweep)) {
        steps.push({ op: "cubic", ...seg });
        cur = { x: seg.x, y: seg.y };
      }
      converted = true;
    } else if (cmd === "T") {
      // angle-ellipseto: (cx cy rx ry startDeg endDeg) — a line to the arc's
      // start, then the arc itself.
      const n = readNums(6);
      if (!n) {
        return null;
      }
      const cx = n[0] / viewBoxW;
      const cy = n[1] / viewBoxH;
      const rx = n[2] / viewBoxW;
      const ry = n[3] / viewBoxH;
      if (rx === 0 || ry === 0) {
        return null;
      }
      const start = (n[4] * Math.PI) / 180;
      let sweep = ((n[5] - n[4]) * Math.PI) / 180;
      if (sweep === 0) {
        sweep = 2 * Math.PI;
      }
      const startPt = { x: cx + rx * Math.cos(start), y: cy + ry * Math.sin(start) };
      steps.push({ op: "line", x: startPt.x, y: startPt.y });
      for (const seg of arcToCubics(cx, cy, rx, ry, start, sweep)) {
        steps.push({ op: "cubic", ...seg });
        cur = { x: seg.x, y: seg.y };
      }
      converted = true;
    } else if (cmd === "Z") {
      steps.push({ op: "close" });
      closed = true;
    } else if (cmd === "N") {
      flush();
    } else if (cmd === "F" || cmd === "S") {
      // no-fill / no-stroke hints: not modelled, ignored (no parameters).
      continue;
    } else {
      // An unmodelled command (W/V/U/X/Y/B, lowercase relatives, …) or a
      // stray token — keep the shape on the preservation path.
      return null;
    }
  }
  flush();
  if (subpaths.length === 0) {
    return null;
  }
  return { path: { subpaths }, converted };
}

/** Serialises a `ShapePath` as an ODF `draw:enhanced-path` command string on
 * the shared 21600 grid, terminated by `N` as ODF requires (AC-27). */
export function enhancedPathString(path: ShapePath): string {
  const g = PATH_GRID;
  const n = (v: number): string => String(Math.round(v * g));
  const parts: string[] = [];
  for (const subpath of path.subpaths) {
    for (const step of subpath.steps) {
      if (step.op === "move") {
        parts.push(`M ${n(step.x)} ${n(step.y)}`);
      } else if (step.op === "line") {
        parts.push(`L ${n(step.x)} ${n(step.y)}`);
      } else if (step.op === "cubic") {
        parts.push(`C ${n(step.x1)} ${n(step.y1)} ${n(step.x2)} ${n(step.y2)} ${n(step.x)} ${n(step.y)}`);
      } else {
        parts.push("Z");
      }
    }
  }
  parts.push("N");
  return parts.join(" ");
}

/** The `svg:viewBox` a path writes into — the shared grid in both axes. */
export function enhancedPathViewBox(): string {
  return `0 0 ${PATH_GRID} ${PATH_GRID}`;
}
