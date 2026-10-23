// Motion-path geometry (SPEC-105) — the format-neutral, framework-free maths
// behind the `motionPath` effect: how a `MotionPath` preset resolves to path
// commands, how a path string parses (and when it is out of subset), how a
// read path is recognised as one of the presets again, and the points the
// players, the stage overlay and the endpoint drag work from.
//
// Coordinates are SLIDE FRACTIONS — 1.0 is the full slide width on x and the
// full slide height on y — with the origin at the frame's centre, in both
// wire formats, so a path is deck-size independent. Screen axes: +x right,
// +y down.
//
// Layering rule (presentation/boundary.spec.ts): imports only presentation
// siblings and src/lib/core.
import type { MotionPath, TransitionDirection } from "./model";
import { DEFAULT_PATH_DISTANCE, MAX_PATH_DISTANCE, MIN_PATH_DISTANCE } from "./animations";

/** One path command, its arguments absolute (x, y pairs). */
export interface PathCommand {
  op: "M" | "L" | "C" | "Z";
  args: number[];
}

export interface PathPoint {
  x: number;
  y: number;
}

const UNIT: Record<TransitionDirection, PathPoint> = {
  right: { x: 1, y: 0 },
  left: { x: -1, y: 0 },
  up: { x: 0, y: -1 },
  down: { x: 0, y: 1 },
};

/** The bow side of an arc: the travel direction turned a quarter
 * anticlockwise on screen (a rightward arc bows up). */
function leftOf(u: PathPoint): PathPoint {
  return { x: u.y, y: -u.x };
}

/** A turn's second leg: the travel direction turned a quarter clockwise on
 * screen (a rightward turn continues down). */
function rightOf(u: PathPoint): PathPoint {
  return { x: -u.y, y: u.x };
}

/** The cubic-bezier handle ratio for a quarter circle. */
const KAPPA = 0.5523;

/** A path's direction, defaulted ("right"). */
export function motionPathDirection(path: MotionPath): TransitionDirection {
  return path.direction ?? "right";
}

/** A path's distance (slide fraction), defaulted and clamped to 0.05–1. */
export function motionPathDistance(path: MotionPath): number {
  const d = path.distanceFraction ?? DEFAULT_PATH_DISTANCE;
  return Math.min(MAX_PATH_DISTANCE, Math.max(MIN_PATH_DISTANCE, d));
}

/** The commands a path resolves to: `line` is a straight `direction` ×
 * distance move, `arc` reaches the same end bowing perpendicular by half the
 * distance, `turn` is an L (distance along `direction`, then distance a
 * quarter clockwise), `loop` a closed circle of radius distance / 2 to the
 * right of the frame, and `custom` its own `d`. Undefined for a custom path
 * that is out of subset. */
export function resolveMotionPath(path: MotionPath): PathCommand[] | undefined {
  if (path.preset === "custom") {
    return path.d === undefined ? undefined : parseMotionPathD(path.d);
  }
  const d = motionPathDistance(path);
  const u = UNIT[motionPathDirection(path)];
  const end = { x: u.x * d, y: u.y * d };
  switch (path.preset) {
    case "line":
      return [move(0, 0), { op: "L", args: [end.x, end.y] }];
    case "arc": {
      // A cubic whose midpoint sits d/2 off the chord: handles at 4/3 of that.
      const p = leftOf(u);
      const k = (2 * d) / 3;
      return [
        move(0, 0),
        { op: "C", args: [p.x * k, p.y * k, end.x + p.x * k, end.y + p.y * k, end.x, end.y] },
      ];
    }
    case "turn": {
      const q = rightOf(u);
      return [move(0, 0), { op: "L", args: [end.x, end.y] }, { op: "L", args: [end.x + q.x * d, end.y + q.y * d] }];
    }
    case "loop": {
      const r = d / 2;
      const k = KAPPA * r;
      return [
        move(0, 0),
        { op: "C", args: [0, -k, r - k, -r, r, -r] },
        { op: "C", args: [r + k, -r, 2 * r, -k, 2 * r, 0] },
        { op: "C", args: [2 * r, k, r + k, r, r, r] },
        { op: "C", args: [r - k, r, 0, k, 0, 0] },
        { op: "Z", args: [] },
      ];
    }
    default:
      return undefined;
  }
}

function move(x: number, y: number): PathCommand {
  return { op: "M", args: [x, y] };
}

/** A coordinate as the wire writes it: at most four decimals, no `-0`. */
export function formatPathNumber(n: number): string {
  const rounded = Number(n.toFixed(4));
  return Object.is(rounded, -0) ? "0" : String(rounded);
}

/** Path commands as a canonical `M x y L x y C … Z` string. */
export function formatPathD(commands: readonly PathCommand[]): string {
  return commands
    .map((c) => (c.args.length > 0 ? `${c.op} ${c.args.map(formatPathNumber).join(" ")}` : c.op))
    .join(" ");
}

const ARG_COUNT: Record<PathCommand["op"], number> = { M: 2, L: 2, C: 6, Z: 0 };
const TOKEN = /([MLCZEmlcze])|([-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?)|([\s,]+)|(.)/g;

/** Parses a motion-path string to absolute commands, or undefined when it is
 * out of subset: only `M`/`L`/`C`/`Z` with numeric arguments (lower-case
 * relative forms are made absolute; a trailing PowerPoint `E` end marker is
 * tolerated), starting with a move. */
export function parseMotionPathD(d: string): PathCommand[] | undefined {
  const tokens: (string | number)[] = [];
  let bad = false;
  for (const m of d.matchAll(TOKEN)) {
    if (m[1] !== undefined) {
      tokens.push(m[1]);
    } else if (m[2] !== undefined) {
      tokens.push(Number(m[2]));
    } else if (m[4] !== undefined) {
      bad = true;
    }
  }
  if (bad) {
    return undefined;
  }
  // A trailing PowerPoint end marker.
  if (tokens.length > 0 && (tokens[tokens.length - 1] === "E" || tokens[tokens.length - 1] === "e")) {
    tokens.pop();
  }
  const commands: PathCommand[] = [];
  let cursor = { x: 0, y: 0 };
  let start = { x: 0, y: 0 };
  let i = 0;
  while (i < tokens.length) {
    const token = tokens[i];
    if (typeof token !== "string") {
      return undefined;
    }
    i++;
    const upper = token.toUpperCase() as PathCommand["op"] | "E";
    if (upper === "E" || !(upper in ARG_COUNT)) {
      return undefined;
    }
    const op = upper as PathCommand["op"];
    const relative = token !== upper;
    if (commands.length === 0 && op !== "M") {
      return undefined;
    }
    if (op === "Z") {
      commands.push({ op: "Z", args: [] });
      cursor = start;
      continue;
    }
    // One or more argument sets (SVG's implicit repetition).
    let sets = 0;
    while (i < tokens.length && typeof tokens[i] === "number") {
      const args: number[] = [];
      for (let k = 0; k < ARG_COUNT[op]; k++) {
        const v = tokens[i++];
        if (typeof v !== "number" || !Number.isFinite(v)) {
          return undefined;
        }
        args.push(v);
      }
      const abs = args.map((v, k) => (relative ? v + (k % 2 === 0 ? cursor.x : cursor.y) : v));
      // A repeated M argument set is an implicit L.
      const thisOp = op === "M" && sets > 0 ? "L" : op;
      commands.push({ op: thisOp, args: abs });
      cursor = { x: abs[abs.length - 2], y: abs[abs.length - 1] };
      if (thisOp === "M") {
        start = cursor;
      }
      sets++;
    }
    if (sets === 0) {
      return undefined;
    }
  }
  return commands.length > 0 ? commands : undefined;
}

/** Whether two command lists match within `epsilon` per coordinate. */
export function pathsMatch(a: readonly PathCommand[], b: readonly PathCommand[], epsilon = 1e-3): boolean {
  return (
    a.length === b.length &&
    a.every(
      (c, i) =>
        c.op === b[i].op &&
        c.args.length === b[i].args.length &&
        c.args.every((v, k) => Math.abs(v - b[i].args[k]) <= epsilon),
    )
  );
}

/** The path's start and end points (a closing `Z` returns to the start). */
export function pathEndpoints(commands: readonly PathCommand[]): { start: PathPoint; end: PathPoint } {
  let start = { x: 0, y: 0 };
  let cursor = { x: 0, y: 0 };
  let first: PathPoint | undefined;
  for (const c of commands) {
    if (c.op === "Z") {
      cursor = start;
      continue;
    }
    cursor = { x: c.args[c.args.length - 2], y: c.args[c.args.length - 1] };
    if (c.op === "M") {
      start = cursor;
    }
    first ??= cursor;
  }
  return { start: first ?? { x: 0, y: 0 }, end: cursor };
}

/** The point an edit handle sits on: a path's end, except a loop (which
 * closes on its start) — there the far side of the circle. */
export function motionPathHandle(path: MotionPath): PathPoint {
  const commands = resolveMotionPath(path);
  if (!commands) {
    return { x: 0, y: 0 };
  }
  if (path.preset === "loop") {
    return { x: motionPathDistance(path), y: 0 };
  }
  return pathEndpoints(commands).end;
}

/** The frame's net offset once the path has played: its end point. */
export function motionPathEnd(path: MotionPath): PathPoint {
  const commands = resolveMotionPath(path);
  return commands ? pathEndpoints(commands).end : { x: 0, y: 0 };
}

const DIRECTIONS: TransitionDirection[] = ["right", "left", "up", "down"];

function round4(n: number): number {
  return Number(n.toFixed(4));
}

/** A read path recognised as `preset` again: the direction and distance whose
 * resolved commands match `commands`, with defaults omitted — or undefined
 * when no preset parameters reproduce it (a path this spec cannot resolve,
 * which keeps the slide's timing preserved). `custom` keeps the path itself,
 * in canonical form. */
export function inferMotionPath(preset: MotionPath["preset"], commands: readonly PathCommand[]): MotionPath | undefined {
  if (preset === "custom") {
    return { preset, d: formatPathD(commands) };
  }
  const { end } = pathEndpoints(commands);
  const directions = preset === "loop" ? (["right"] as TransitionDirection[]) : DIRECTIONS;
  for (const direction of directions) {
    const u = UNIT[direction];
    const reach =
      preset === "loop" ? Math.max(...commands.flatMap((c) => c.args.filter((_, k) => k % 2 === 0))) : end.x * u.x + end.y * u.y;
    const distance = round4(reach);
    if (distance < MIN_PATH_DISTANCE - 1e-9 || distance > MAX_PATH_DISTANCE + 1e-9) {
      continue;
    }
    const candidate: MotionPath = { preset };
    if (direction !== "right" && preset !== "loop") {
      candidate.direction = direction;
    }
    if (distance !== DEFAULT_PATH_DISTANCE) {
      candidate.distanceFraction = distance;
    }
    const resolved = resolveMotionPath(candidate);
    if (resolved && pathsMatch(resolved, commands)) {
      return candidate;
    }
  }
  return undefined;
}

/** `count + 1` points along the path at even parameter steps per segment
 * (the CSS keyframes the players animate through). */
export function samplePath(commands: readonly PathCommand[], count: number): PathPoint[] {
  const segments: ((t: number) => PathPoint)[] = [];
  let cursor = { x: 0, y: 0 };
  let start = { x: 0, y: 0 };
  for (const c of commands) {
    const from = cursor;
    if (c.op === "M") {
      cursor = { x: c.args[0], y: c.args[1] };
      start = cursor;
      continue;
    }
    const to = c.op === "Z" ? start : { x: c.args[c.args.length - 2], y: c.args[c.args.length - 1] };
    if (c.op === "C") {
      const [x1, y1, x2, y2] = c.args;
      segments.push((t) => {
        const s = 1 - t;
        return {
          x: s * s * s * from.x + 3 * s * s * t * x1 + 3 * s * t * t * x2 + t * t * t * to.x,
          y: s * s * s * from.y + 3 * s * s * t * y1 + 3 * s * t * t * y2 + t * t * t * to.y,
        };
      });
    } else {
      segments.push((t) => ({ x: from.x + (to.x - from.x) * t, y: from.y + (to.y - from.y) * t }));
    }
    cursor = to;
  }
  if (segments.length === 0) {
    return Array.from({ length: count + 1 }, () => ({ ...cursor }));
  }
  return Array.from({ length: count + 1 }, (_, i) => {
    const position = (i / count) * segments.length;
    const index = Math.min(segments.length - 1, Math.floor(position));
    return segments[index](position - index);
  });
}

/** The path as an SVG `d` in surface pixels: fractions scaled by the slide's
 * pixel size and offset to the frame centre `origin` (px). */
export function pathDInPixels(
  commands: readonly PathCommand[],
  origin: PathPoint,
  slideWidthPx: number,
  slideHeightPx: number,
): string {
  return commands
    .map((c) =>
      c.args.length > 0
        ? `${c.op} ${c.args
            .map((v, k) => (k % 2 === 0 ? origin.x + v * slideWidthPx : origin.y + v * slideHeightPx).toFixed(2))
            .join(" ")}`
        : c.op,
    )
    .join(" ");
}

/** The path a dragged handle at `point` (slide fractions from the frame's
 * centre) asks for, keeping `path`'s preset: the four directions the model
 * can express, chosen by where the handle sits, and the distance it reaches,
 * clamped to 0.05–1 and rounded to whole percent. A line or arc takes the
 * dominant axis; a turn's end lies in one quadrant per direction; a loop
 * keeps its side and takes the reach. Never called for a custom path. */
export function motionPathFromHandle(path: MotionPath, point: PathPoint): MotionPath {
  let direction: TransitionDirection = motionPathDirection(path);
  let reach: number;
  if (path.preset === "turn") {
    direction = point.x >= 0 ? (point.y >= 0 ? "right" : "up") : point.y >= 0 ? "down" : "left";
    reach = (Math.abs(point.x) + Math.abs(point.y)) / 2;
  } else if (path.preset === "loop") {
    reach = Math.max(Math.abs(point.x), Math.abs(point.y));
  } else if (Math.abs(point.x) >= Math.abs(point.y)) {
    direction = point.x >= 0 ? "right" : "left";
    reach = Math.abs(point.x);
  } else {
    direction = point.y >= 0 ? "down" : "up";
    reach = Math.abs(point.y);
  }
  const distance = Math.min(MAX_PATH_DISTANCE, Math.max(MIN_PATH_DISTANCE, Math.round(reach * 100) / 100));
  const next: MotionPath = { preset: path.preset };
  if (direction !== "right" && path.preset !== "loop") {
    next.direction = direction;
  }
  if (distance !== DEFAULT_PATH_DISTANCE) {
    next.distanceFraction = distance;
  }
  return next;
}
