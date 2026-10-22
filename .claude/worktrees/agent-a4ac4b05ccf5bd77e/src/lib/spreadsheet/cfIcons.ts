// Conditional-format icon artwork (SPEC-129 "Icon artwork") — THIS FILE is
// the normative source for every icon set's glyphs, colours and labels.
// The artwork is this repo's own, authored here from geometry (the
// core/chartIcons.ts precedent): block arrows rotated about the viewBox
// centre, circles, triangles, flags, ticks, crosses, a star from its ten
// polar points, signal bars, quarter pies, boxes and faces — never copied
// from Office's or LibreOffice's image resources. Pure path-string maths,
// no DOM, so it unit-tests like every other lib module.
//
// Every glyph is drawn in a 16-unit square viewBox as a list of paths, each
// painted "icon" (filled with the icon's colour), "white" (filled white —
// the tick inside a circle, a face's eyes) or "none" (unfilled, stroked in
// the icon's colour — an empty star, an unlit bar).
import type { IconSetName } from "./conditionalFormats";

export type CfGlyph =
  | "arrowUp"
  | "arrowUpRight"
  | "arrowRight"
  | "arrowDownRight"
  | "arrowDown"
  | "circle"
  | "circleRimmed"
  | "triangleUp"
  | "triangleDown"
  | "dash"
  | "diamond"
  | "flag"
  | "check"
  | "cross"
  | "exclaim"
  | "checkCircle"
  | "crossCircle"
  | "exclaimCircle"
  | "star"
  | "starHalf"
  | "starEmpty"
  | "bars0"
  | "bars1"
  | "bars2"
  | "bars3"
  | "bars4"
  | "quarter0"
  | "quarter1"
  | "quarter2"
  | "quarter3"
  | "quarter4"
  | "box0"
  | "box1"
  | "box2"
  | "box3"
  | "box4"
  | "smileHappy"
  | "smileNeutral"
  | "smileSad";

export interface CfIconDef {
  glyph: CfGlyph;
  color: string;
  label: string;
}

export interface CfGlyphPath {
  d: string;
  paint: "icon" | "white" | "none";
}

export const CF_ICON_COLORS = {
  green: "#2E9E4F",
  amber: "#E8A317",
  red: "#D64541",
  grey: "#808080",
  black: "#333333",
  pink: "#E99A9A",
  blue: "#3B78C3",
  gold: "#F2B01E",
} as const;

/** The side of the square viewBox every glyph is authored in. */
export const CF_GLYPH_SIZE = 16;

/** The stroke width a "none" path is drawn with, in viewBox units. */
export const CF_GLYPH_STROKE = 1.2;

const C = CF_ICON_COLORS;
const CENTRE = CF_GLYPH_SIZE / 2;

type Point = readonly [number, number];

function num(value: number): string {
  return String(Math.round(value * 100) / 100);
}

/** A closed polygon through `points`. */
function polygon(points: readonly Point[]): string {
  return `${points.map(([x, y], i) => `${i === 0 ? "M" : "L"}${num(x)} ${num(y)}`).join("")}Z`;
}

/** A full circle as two arcs. */
function circle(cx: number, cy: number, r: number): string {
  return `M${num(cx - r)} ${num(cy)}A${num(r)} ${num(r)} 0 1 0 ${num(cx + r)} ${num(cy)}A${num(r)} ${num(r)} 0 1 0 ${num(cx - r)} ${num(cy)}Z`;
}

/** `points` rotated by `degrees` (clockwise on screen) about the centre. */
function rotate(points: readonly Point[], degrees: number): Point[] {
  const a = (degrees * Math.PI) / 180;
  const cos = Math.cos(a);
  const sin = Math.sin(a);
  return points.map(([x, y]) => {
    const dx = x - CENTRE;
    const dy = y - CENTRE;
    return [CENTRE + dx * cos - dy * sin, CENTRE + dx * sin + dy * cos] as const;
  });
}

/** `points` scaled by `factor` about the centre. */
function scale(points: readonly Point[], factor: number): Point[] {
  return points.map(([x, y]) => [CENTRE + (x - CENTRE) * factor, CENTRE + (y - CENTRE) * factor] as const);
}

/** The right-pointing block arrow every arrow glyph is a rotation of: a
 * head reaching the right edge, a shaft from the left edge. */
const ARROW_RIGHT: readonly Point[] = [
  [15, 8],
  [8.5, 1.5],
  [8.5, 5.4],
  [1.2, 5.4],
  [1.2, 10.6],
  [8.5, 10.6],
  [8.5, 14.5],
];

/** A tick: a short left stroke meeting a long right one. */
const CHECK: readonly Point[] = [
  [1.5, 8.6],
  [3.6, 6.5],
  [6.4, 9.3],
  [12.4, 3.3],
  [14.5, 5.4],
  [6.4, 13.5],
];

/** A saltire cross of two bars 2.1 units wide. */
const CROSS: readonly Point[] = [
  [3, 1.5],
  [8, 6.5],
  [13, 1.5],
  [14.5, 3],
  [9.5, 8],
  [14.5, 13],
  [13, 14.5],
  [8, 9.5],
  [3, 14.5],
  [1.5, 13],
  [6.5, 8],
  [1.5, 3],
];

/** An exclamation mark: a tapering bar and its dot. */
const EXCLAIM_BAR: readonly Point[] = [
  [6.4, 1.5],
  [9.6, 1.5],
  [8.9, 10],
  [7.1, 10],
];
const EXCLAIM_DOT: readonly Point[] = [
  [6.6, 11.6],
  [9.4, 11.6],
  [9.4, 14.4],
  [6.6, 14.4],
];

/** The five-pointed star's ten points, clockwise from the top: outer and
 * inner radii about a centre a little below the middle, so it sits level. */
function starPoints(): Point[] {
  const cx = CENTRE;
  const cy = 8.6;
  const outer = 7.2;
  const inner = 2.95;
  const points: Point[] = [];
  for (let i = 0; i < 10; i += 1) {
    const r = i % 2 === 0 ? outer : inner;
    const a = (-90 + i * 36) * (Math.PI / 180);
    points.push([cx + r * Math.cos(a), cy + r * Math.sin(a)]);
  }
  return points;
}

/** The left half of the star: the bottom inner point up the left arms to
 * the top point, closed down the vertical axis. */
function halfStarPoints(): Point[] {
  const star = starPoints();
  return [star[5]!, star[6]!, star[7]!, star[8]!, star[9]!, star[0]!];
}

/** Signal bars: four bars of rising height, the first `lit` filled. */
function bars(lit: number): CfGlyphPath[] {
  const paths: CfGlyphPath[] = [];
  for (let i = 0; i < 4; i += 1) {
    const left = 1.5 + i * 3.4;
    const top = 11.5 - i * 3.2;
    const d = polygon([
      [left, top],
      [left + 2.4, top],
      [left + 2.4, 14.5],
      [left, 14.5],
    ]);
    paths.push({ d, paint: i < lit ? "icon" : "none" });
  }
  return paths;
}

/** A pie: the outline, and `quarters` quarters filled clockwise from 12
 * o'clock (all four: the whole disc). */
function quarters(count: number): CfGlyphPath[] {
  const r = 6.6;
  const outline: CfGlyphPath = { d: circle(CENTRE, CENTRE, r), paint: "none" };
  if (count === 0) {
    return [outline];
  }
  if (count === 4) {
    return [{ d: circle(CENTRE, CENTRE, r), paint: "icon" }];
  }
  const end = (count * 90 - 90) * (Math.PI / 180);
  const x = CENTRE + r * Math.cos(end);
  const y = CENTRE + r * Math.sin(end);
  const sector = `M${num(CENTRE)} ${num(CENTRE)}L${num(CENTRE)} ${num(CENTRE - r)}A${num(r)} ${num(r)} 0 ${count > 2 ? 1 : 0} 1 ${num(x)} ${num(y)}Z`;
  return [outline, { d: sector, paint: "icon" }];
}

/** A 2 × 2 grid of boxes, the first `lit` filled (reading order). */
function boxes(lit: number): CfGlyphPath[] {
  const cells: Point[] = [
    [1.5, 1.5],
    [8.75, 1.5],
    [1.5, 8.75],
    [8.75, 8.75],
  ];
  return cells.map(([x, y], i) => ({
    d: polygon([
      [x, y],
      [x + 5.75, y],
      [x + 5.75, y + 5.75],
      [x, y + 5.75],
    ]),
    paint: i < lit ? "icon" : "none",
  }));
}

/** A face: the disc, two eyes and a mouth curving up, flat or down. */
function face(mood: "happy" | "neutral" | "sad"): CfGlyphPath[] {
  const mouth =
    mood === "happy"
      ? "M4.4 9.4Q8 13.9 11.6 9.4L10.3 9.4Q8 11.9 5.7 9.4Z"
      : mood === "sad"
        ? "M4.4 12.9Q8 8.4 11.6 12.9L10.3 12.9Q8 10.4 5.7 12.9Z"
        : polygon([
            [4.8, 10.3],
            [11.2, 10.3],
            [11.2, 11.7],
            [4.8, 11.7],
          ]);
  return [
    { d: circle(CENTRE, CENTRE, 7), paint: "icon" },
    { d: circle(5.6, 6.2, 1.15), paint: "white" },
    { d: circle(10.4, 6.2, 1.15), paint: "white" },
    { d: mouth, paint: "white" },
  ];
}

/** A symbol inside a disc: the disc in the icon colour, the symbol white. */
function circled(symbol: readonly (readonly Point[])[]): CfGlyphPath[] {
  return [{ d: circle(CENTRE, CENTRE, 7.2), paint: "icon" }, ...symbol.map((points) => ({ d: polygon(scale(points, 0.62)), paint: "white" as const }))];
}

/** The paths of `glyph`, in a 16-unit viewBox (see the module header). */
export function cfGlyphPaths(glyph: CfGlyph): CfGlyphPath[] {
  switch (glyph) {
    case "arrowRight":
      return [{ d: polygon(ARROW_RIGHT), paint: "icon" }];
    case "arrowUp":
      return [{ d: polygon(rotate(ARROW_RIGHT, -90)), paint: "icon" }];
    case "arrowDown":
      return [{ d: polygon(rotate(ARROW_RIGHT, 90)), paint: "icon" }];
    case "arrowUpRight":
      return [{ d: polygon(scale(rotate(ARROW_RIGHT, -45), 0.92)), paint: "icon" }];
    case "arrowDownRight":
      return [{ d: polygon(scale(rotate(ARROW_RIGHT, 45), 0.92)), paint: "icon" }];
    case "circle":
      return [{ d: circle(CENTRE, CENTRE, 6.8), paint: "icon" }];
    case "circleRimmed":
      return [
        { d: circle(CENTRE, CENTRE, 7), paint: "none" },
        { d: circle(CENTRE, CENTRE, 5), paint: "icon" },
      ];
    case "triangleUp":
      return [
        {
          d: polygon([
            [8, 2],
            [14.8, 13.6],
            [1.2, 13.6],
          ]),
          paint: "icon",
        },
      ];
    case "triangleDown":
      return [
        {
          d: polygon([
            [8, 14],
            [14.8, 2.4],
            [1.2, 2.4],
          ]),
          paint: "icon",
        },
      ];
    case "dash":
      return [
        {
          d: polygon([
            [2, 6.4],
            [14, 6.4],
            [14, 9.6],
            [2, 9.6],
          ]),
          paint: "icon",
        },
      ];
    case "diamond":
      return [
        {
          d: polygon([
            [8, 1],
            [15, 8],
            [8, 15],
            [1, 8],
          ]),
          paint: "icon",
        },
      ];
    case "flag":
      return [
        {
          d: polygon([
            [2.6, 1.2],
            [4.2, 1.2],
            [4.2, 15],
            [2.6, 15],
          ]),
          paint: "icon",
        },
        {
          d: polygon([
            [4.2, 1.6],
            [13.8, 1.6],
            [11.2, 5.2],
            [13.8, 8.8],
            [4.2, 8.8],
          ]),
          paint: "icon",
        },
      ];
    case "check":
      return [{ d: polygon(CHECK), paint: "icon" }];
    case "cross":
      return [{ d: polygon(CROSS), paint: "icon" }];
    case "exclaim":
      return [
        { d: polygon(EXCLAIM_BAR), paint: "icon" },
        { d: polygon(EXCLAIM_DOT), paint: "icon" },
      ];
    case "checkCircle":
      return circled([CHECK]);
    case "crossCircle":
      return circled([CROSS]);
    case "exclaimCircle":
      return circled([EXCLAIM_BAR, EXCLAIM_DOT]);
    case "star":
      return [{ d: polygon(starPoints()), paint: "icon" }];
    case "starHalf":
      return [
        { d: polygon(halfStarPoints()), paint: "icon" },
        { d: polygon(starPoints()), paint: "none" },
      ];
    case "starEmpty":
      return [{ d: polygon(starPoints()), paint: "none" }];
    case "bars0":
      return bars(0);
    case "bars1":
      return bars(1);
    case "bars2":
      return bars(2);
    case "bars3":
      return bars(3);
    case "bars4":
      return bars(4);
    case "quarter0":
      return quarters(0);
    case "quarter1":
      return quarters(1);
    case "quarter2":
      return quarters(2);
    case "quarter3":
      return quarters(3);
    case "quarter4":
      return quarters(4);
    case "box0":
      return boxes(0);
    case "box1":
      return boxes(1);
    case "box2":
      return boxes(2);
    case "box3":
      return boxes(3);
    case "box4":
      return boxes(4);
    case "smileHappy":
      return face("happy");
    case "smileNeutral":
      return face("neutral");
    case "smileSad":
      return face("sad");
  }
}

function icon(glyph: CfGlyph, color: string, label: string): CfIconDef {
  return { glyph, color, label };
}

const ARROWS_3 = (red: string, amber: string, green: string, prefix: [string, string, string]): CfIconDef[] => [
  icon("arrowDown", red, `${prefix[0]} down arrow`),
  icon("arrowRight", amber, `${prefix[1]} right arrow`),
  icon("arrowUp", green, `${prefix[2]} up arrow`),
];

const ARROWS_4 = (red: string, amber: string, green: string, prefix: [string, string, string]): CfIconDef[] => [
  icon("arrowDown", red, `${prefix[0]} down arrow`),
  icon("arrowDownRight", amber, `${prefix[1]} down-right arrow`),
  icon("arrowUpRight", amber, `${prefix[1]} up-right arrow`),
  icon("arrowUp", green, `${prefix[2]} up arrow`),
];

const ARROWS_5 = (red: string, amber: string, green: string, prefix: [string, string, string]): CfIconDef[] => [
  icon("arrowDown", red, `${prefix[0]} down arrow`),
  icon("arrowDownRight", amber, `${prefix[1]} down-right arrow`),
  icon("arrowRight", amber, `${prefix[1]} right arrow`),
  icon("arrowUpRight", amber, `${prefix[1]} up-right arrow`),
  icon("arrowUp", green, `${prefix[2]} up arrow`),
];

const COLOURED: [string, string, string] = ["Red", "Amber", "Green"];
const GREY: [string, string, string] = ["Grey", "Grey", "Grey"];

/** Every set's icons, lowest → highest (SPEC-129 "Icon artwork"). */
export const CF_ICON_SETS: Record<IconSetName, CfIconDef[]> = {
  "3Arrows": ARROWS_3(C.red, C.amber, C.green, COLOURED),
  "3ArrowsGray": ARROWS_3(C.grey, C.grey, C.grey, GREY),
  "3Flags": [icon("flag", C.red, "Red flag"), icon("flag", C.amber, "Amber flag"), icon("flag", C.green, "Green flag")],
  "3TrafficLights1": [icon("circle", C.red, "Red circle"), icon("circle", C.amber, "Amber circle"), icon("circle", C.green, "Green circle")],
  "3TrafficLights2": [
    icon("circleRimmed", C.red, "Red rimmed circle"),
    icon("circleRimmed", C.amber, "Amber rimmed circle"),
    icon("circleRimmed", C.green, "Green rimmed circle"),
  ],
  "3Signs": [icon("diamond", C.red, "Red diamond"), icon("triangleUp", C.amber, "Amber triangle"), icon("circle", C.green, "Green circle")],
  "3Symbols": [
    icon("crossCircle", C.red, "Red circled cross"),
    icon("exclaimCircle", C.amber, "Amber circled exclamation mark"),
    icon("checkCircle", C.green, "Green circled check mark"),
  ],
  "3Symbols2": [icon("cross", C.red, "Red cross"), icon("exclaim", C.amber, "Amber exclamation mark"), icon("check", C.green, "Green check mark")],
  "3Stars": [icon("starEmpty", C.gold, "Empty star"), icon("starHalf", C.gold, "Half star"), icon("star", C.gold, "Full star")],
  "3Triangles": [icon("triangleDown", C.red, "Red down triangle"), icon("dash", C.amber, "Amber dash"), icon("triangleUp", C.green, "Green up triangle")],
  "3Smilies": [icon("smileSad", C.black, "Black sad face"), icon("smileNeutral", C.black, "Black neutral face"), icon("smileHappy", C.black, "Black happy face")],
  "3ColorSmilies": [icon("smileSad", C.red, "Red sad face"), icon("smileNeutral", C.amber, "Amber neutral face"), icon("smileHappy", C.green, "Green happy face")],
  "4Arrows": ARROWS_4(C.red, C.amber, C.green, COLOURED),
  "4ArrowsGray": ARROWS_4(C.grey, C.grey, C.grey, GREY),
  "4RedToBlack": [icon("circle", C.black, "Black circle"), icon("circle", C.grey, "Grey circle"), icon("circle", C.pink, "Pink circle"), icon("circle", C.red, "Red circle")],
  "4Rating": [icon("bars1", C.blue, "1 of 4 bars"), icon("bars2", C.blue, "2 of 4 bars"), icon("bars3", C.blue, "3 of 4 bars"), icon("bars4", C.blue, "4 of 4 bars")],
  "4TrafficLights": [
    icon("circle", C.black, "Black circle"),
    icon("circle", C.red, "Red circle"),
    icon("circle", C.amber, "Amber circle"),
    icon("circle", C.green, "Green circle"),
  ],
  "5Arrows": ARROWS_5(C.red, C.amber, C.green, COLOURED),
  "5ArrowsGray": ARROWS_5(C.grey, C.grey, C.grey, GREY),
  "5Rating": [
    icon("bars0", C.blue, "0 of 4 bars"),
    icon("bars1", C.blue, "1 of 4 bars"),
    icon("bars2", C.blue, "2 of 4 bars"),
    icon("bars3", C.blue, "3 of 4 bars"),
    icon("bars4", C.blue, "4 of 4 bars"),
  ],
  "5Quarters": [
    icon("quarter0", C.black, "Empty circle"),
    icon("quarter1", C.black, "Quarter circle"),
    icon("quarter2", C.black, "Half circle"),
    icon("quarter3", C.black, "Three-quarter circle"),
    icon("quarter4", C.black, "Full circle"),
  ],
  "5Boxes": [
    icon("box0", C.blue, "0 of 4 boxes"),
    icon("box1", C.blue, "1 of 4 boxes"),
    icon("box2", C.blue, "2 of 4 boxes"),
    icon("box3", C.blue, "3 of 4 boxes"),
    icon("box4", C.blue, "4 of 4 boxes"),
  ],
  NoIcons: [],
};

/** The icon at `index` of `set`, or undefined (NoIcons, an unknown set
 * name, or out of range) — `set` as a component prop carries it, a string. */
export function cfIconAt(set: IconSetName | string, index: number): CfIconDef | undefined {
  return Object.hasOwn(CF_ICON_SETS, set) ? CF_ICON_SETS[set as IconSetName][index] : undefined;
}

/** The fill a glyph path takes: the icon colour, white, or none (stroked). */
export function cfPaintFill(paint: CfGlyphPath["paint"], color: string): string {
  return paint === "icon" ? color : paint === "white" ? "#FFFFFF" : "none";
}

/** A set's name as menus show it ("3 Arrows", "3 Traffic Lights (Rimmed)"). */
export const ICON_SET_LABELS: Record<IconSetName, string> = {
  "3Arrows": "3 Arrows",
  "3ArrowsGray": "3 Arrows (Grey)",
  "3Flags": "3 Flags",
  "3TrafficLights1": "3 Traffic Lights",
  "3TrafficLights2": "3 Traffic Lights (Rimmed)",
  "3Signs": "3 Signs",
  "3Symbols": "3 Symbols (Circled)",
  "3Symbols2": "3 Symbols",
  "3Stars": "3 Stars",
  "3Triangles": "3 Triangles",
  "3Smilies": "3 Smilies",
  "3ColorSmilies": "3 Coloured Smilies",
  "4Arrows": "4 Arrows",
  "4ArrowsGray": "4 Arrows (Grey)",
  "4RedToBlack": "Red To Black",
  "4Rating": "4 Ratings",
  "4TrafficLights": "4 Traffic Lights",
  "5Arrows": "5 Arrows",
  "5ArrowsGray": "5 Arrows (Grey)",
  "5Rating": "5 Ratings",
  "5Quarters": "5 Quarters",
  "5Boxes": "5 Boxes",
  NoIcons: "No Cell Icon",
};
