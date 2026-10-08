// The slide clipboard (SPEC-099, normative for the payload): copy/cut/paste
// of whole slides within a deck and across decks IN THIS APP. A copy writes
// two flavours through `ClipboardEvent.clipboardData`, exactly as SPEC-016's
// DOC_MODEL_MIME does — `text/plain` (the slides' outline, which is all
// PowerPoint, Keynote or a text editor will ever get: no rich slide flavour is
// authorable from a browser) and the private `SLIDES_MIME` JSON — and keeps
// the payload in a module-level in-tab store keyed by its `nonce`. Above
// `MAX_SLIDE_CLIPBOARD_BYTES` the private flavour is omitted, and the in-tab
// store (recognised by the outline text) still pastes those slides intact.
//
// The clipboard is an untrusted boundary — anything can write any string
// under any key — so `parseSlideClipboardPayload` rebuilds the payload field
// by field from a whitelist (core/clipboardPayload.ts's envelope + guards),
// never splicing parsed JSON into a model. Every pasted slide then goes
// through `importSlides`' repair rules.
import { buildClipboardEnvelope, isFiniteNumber, isSafeColor, parseClipboardEnvelope } from "../core/clipboardPayload";
import { THEME_COLOR_SLOTS } from "../core/theme";
import { flushPendingMedia, type MediaPart } from "../core/media";
import type { DeckTheme, PresentationFormat, PresentationModel, Slide, SlideFrame, SlideLayout } from "./model";
import { slideOutline } from "./slideModel";

/** The private MIME type the slide payload travels under. */
export const SLIDES_MIME = "application/x-docu-edit-slides+json";

/** The largest private flavour written to the system clipboard; a bigger copy
 * pastes from the in-tab store only. */
export const MAX_SLIDE_CLIPBOARD_BYTES = 8 * 1024 * 1024;

export interface SlideClipboardPayload {
  mimeType: typeof SLIDES_MIME;
  version: 1;
  nonce: string;
  format: PresentationFormat;
  slideWidthCm: number;
  slideHeightCm: number;
  /** Ids present but ignored — always re-minted on paste. */
  slides: Slide[];
  /** Only the parts the copied slides reference. */
  media: MediaPart[];
  /** SPEC-094, for "keep source formatting". */
  theme?: DeckTheme;
  /** The layouts the copied slides use, so a paste can map them by kind. */
  layouts?: SlideLayout[];
}

// --- the in-tab store --------------------------------------------------------

interface StoredCopy {
  nonce: string;
  text: string;
  payload: SlideClipboardPayload;
}

/** The most copies the in-tab store remembers (oldest dropped). */
const STORE_LIMIT = 8;
const store: StoredCopy[] = [];

function remember(entry: StoredCopy): void {
  store.push(entry);
  if (store.length > STORE_LIMIT) {
    store.shift();
  }
}

/** The in-tab copy with this nonce, if this tab made it. */
export function storedSlideCopy(nonce: string): SlideClipboardPayload | undefined {
  return store.find((entry) => entry.nonce === nonce)?.payload;
}

/** The latest in-tab copy whose outline text is exactly `text` — how a paste
 * that carries no private flavour (over the cap, or read through the async
 * text API) is recognised as this tab's own copy. */
export function storedSlideCopyByText(text: string): SlideClipboardPayload | undefined {
  if (text === "") {
    return undefined;
  }
  for (let i = store.length - 1; i >= 0; i--) {
    if (store[i].text === text) {
      return store[i].payload;
    }
  }
  return undefined;
}

/** The latest in-tab copy (the button paste's fallback when the clipboard
 * cannot be read at all). */
export function latestSlideCopy(): SlideClipboardPayload | undefined {
  return store[store.length - 1]?.payload;
}

/** Forgets every in-tab copy (tests). */
export function clearSlideClipboardStore(): void {
  store.length = 0;
}

// --- copy --------------------------------------------------------------------

function mediaIdsOf(frames: SlideFrame[], into: Set<string>): void {
  for (const each of frames) {
    const paragraphs =
      each.kind === "text"
        ? each.paragraphs
        : each.kind === "shape"
          ? (each.paragraphs ?? [])
          : each.kind === "table"
            ? each.rows.flatMap((row) => row.cells.flatMap((cell) => cell.paragraphs))
            : [];
    for (const paragraph of paragraphs) {
      if (paragraph.bulletImageId) {
        into.add(paragraph.bulletImageId);
      }
    }
    if (each.kind === "image") {
      into.add(each.mediaId);
    } else if (each.kind === "media") {
      if (each.mediaId) into.add(each.mediaId);
      if (each.posterMediaId) into.add(each.posterMediaId);
    } else if (each.kind === "group") {
      mediaIdsOf(each.children, into);
    }
  }
}

/** The outline text of the given slides: each slide's title then its lines,
 * slides separated by a blank line — the `text/plain` flavour. */
export function slidesOutlineText(slides: Slide[]): string {
  return slides
    .map((slide) => {
      const outline = slideOutline(slide);
      return [...(outline.title ? [outline.title] : []), ...outline.lines].join("\n");
    })
    .join("\n\n");
}

/** The payload for `model.slides[indexes]` (in deck order), with a fresh nonce. */
export function buildSlideClipboardPayload(model: PresentationModel, indexes: number[]): SlideClipboardPayload {
  const slides = [...new Set(indexes)]
    .filter((index) => index >= 0 && index < model.slides.length)
    .toSorted((a, b) => a - b)
    .map((index) => structuredClone(model.slides[index]));
  const ids = new Set<string>();
  const layoutKeys = new Set<string>();
  for (const slide of slides) {
    mediaIdsOf(slide.frames, ids);
    if (slide.layoutKey) {
      layoutKeys.add(slide.layoutKey);
    }
  }
  flushPendingMedia(model.media);
  const payload: SlideClipboardPayload = {
    mimeType: SLIDES_MIME,
    version: 1,
    nonce: crypto.randomUUID(),
    format: model.format,
    slideWidthCm: model.slideWidthCm,
    slideHeightCm: model.slideHeightCm,
    slides,
    media: (model.media ?? []).filter((part) => ids.has(part.mediaId)).map((part) => ({ ...part })),
  };
  if (model.theme) {
    payload.theme = structuredClone(model.theme);
  }
  const layouts = (model.layouts ?? []).filter((layout) => layoutKeys.has(layout.key));
  if (layouts.length > 0) {
    payload.layouts = structuredClone(layouts);
  }
  return payload;
}

/** Both flavours of a copy of `model.slides[indexes]`, ready for `setData`:
 * `text` (the outline) and `payload` (the private JSON, or null above
 * `MAX_SLIDE_CLIPBOARD_BYTES`). The payload is also kept in the in-tab store,
 * so a paste in this tab is intact either way. */
export function slideClipboardFlavours(
  model: PresentationModel,
  indexes: number[],
): { text: string; payload: string | null; nonce: string } {
  const payload = buildSlideClipboardPayload(model, indexes);
  const text = slidesOutlineText(payload.slides);
  const { mimeType, version, ...body } = payload;
  const json = buildClipboardEnvelope(mimeType, version, body);
  remember({ nonce: payload.nonce, text, payload });
  const tooLarge = new TextEncoder().encode(json).byteLength > MAX_SLIDE_CLIPBOARD_BYTES;
  return { text, payload: tooLarge ? null : json, nonce: payload.nonce };
}

/** What a paste resolves to: the private flavour (the in-tab copy with the
 * same nonce taking precedence over the clipboard's rebuilt one), else this
 * tab's own copy recognised by its outline text; null when the clipboard holds
 * no slides. */
export function resolveSlidePaste(flavours: { slides?: string; text?: string }): SlideClipboardPayload | null {
  if (flavours.slides) {
    const parsed = parseSlideClipboardPayload(flavours.slides);
    if (parsed) {
      return storedSlideCopy(parsed.nonce) ?? parsed;
    }
  }
  return storedSlideCopyByText(flavours.text ?? "") ?? null;
}

// --- the whitelist parser ------------------------------------------------------

type Check = (value: unknown, depth: number) => unknown;

const MAX_DEPTH = 24;
const UNSAFE_KEYS = new Set(["__proto__", "constructor", "prototype"]);

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

const str: Check = (value) => (typeof value === "string" ? value : undefined);
const text: Check = (value) => (typeof value === "string" && value.length > 0 ? value : undefined);
const num: Check = (value) => (isFiniteNumber(value) ? value : undefined);
const int: Check = (value) => (isFiniteNumber(value) && Number.isInteger(value) ? value : undefined);
const positive: Check = (value) => (isFiniteNumber(value) && value > 0 ? value : undefined);
const yes: Check = (value) => (value === true ? true : undefined);
const no: Check = (value) => (value === false ? false : undefined);
const oneOf =
  (...allowed: string[]): Check =>
  (value) =>
    typeof value === "string" && allowed.includes(value) ? value : undefined;

/** A link target safe to carry: http(s), mailto, or an in-document fragment. */
const href: Check = (value) =>
  typeof value === "string" && /^(?:https?:|mailto:|#)/i.test(value.trim()) ? value : undefined;

/** A package-relative path (media ids): no scheme, no traversal. */
const partPath: Check = (value) =>
  typeof value === "string" && /^[A-Za-z0-9_\-./ ]{1,200}$/.test(value) && !value.includes("..") ? value : undefined;

/** A `ModelColor`: a safe literal or a SPEC-094 `ThemeColorRef`. */
const color: Check = (value) => {
  if (isSafeColor(value)) {
    return value;
  }
  if (!isRecord(value) || typeof value.theme !== "string" || !(THEME_COLOR_SLOTS as string[]).includes(value.theme)) {
    return undefined;
  }
  const ref: Record<string, unknown> = { theme: value.theme };
  for (const key of ["lumMod", "lumOff", "shade", "tint"]) {
    if (isFiniteNumber(value[key])) {
      ref[key] = value[key];
    }
  }
  return ref;
};

/** A `ModelFont`: a family name or a `{ theme: "major" | "minor" }` ref. */
const font: Check = (value) => {
  if (typeof value === "string" && value.length > 0 && value.length <= 200) {
    return value;
  }
  return isRecord(value) && (value.theme === "major" || value.theme === "minor") ? { theme: value.theme } : undefined;
};

/** Small leaf structures (chart series/options, cell borders, column style
 * names) rebuilt as plain data: strings, finite numbers, booleans and null,
 * nested arrays/objects with safe keys only. */
const plain: Check = (value, depth) => {
  if (depth > MAX_DEPTH) {
    return undefined;
  }
  if (typeof value === "string" || typeof value === "boolean" || value === null || isFiniteNumber(value)) {
    return value;
  }
  if (Array.isArray(value)) {
    return value.map((item) => plain(item, depth + 1) ?? null);
  }
  if (isRecord(value)) {
    const out: Record<string, unknown> = {};
    for (const [key, child] of Object.entries(value)) {
      const clean = UNSAFE_KEYS.has(key) ? undefined : plain(child, depth + 1);
      if (clean !== undefined) {
        out[key] = clean;
      }
    }
    return out;
  }
  return undefined;
};

const arrayOf =
  (check: Check): Check =>
  (value, depth) => {
    if (!Array.isArray(value) || depth > MAX_DEPTH) {
      return undefined;
    }
    const out: unknown[] = [];
    for (const item of value) {
      const clean = check(item, depth + 1);
      if (clean !== undefined) {
        out.push(clean);
      }
    }
    return out;
  };

/** An object rebuilt from `spec`'s keys only; undefined when a `required`
 * key fails its check. */
const shape =
  (spec: Record<string, Check>, required: string[] = []): Check =>
  (value, depth) => {
    if (!isRecord(value) || depth > MAX_DEPTH) {
      return undefined;
    }
    const out: Record<string, unknown> = {};
    for (const [key, check] of Object.entries(spec)) {
      const clean = check(value[key], depth + 1);
      if (clean !== undefined) {
        out[key] = clean;
      } else if (required.includes(key)) {
        return undefined;
      }
    }
    return out;
  };

const fill: Check = (value, depth) => {
  if (!isRecord(value)) return undefined;
  if (value.type === "none") return { type: "none" };
  return value.type === "solid" ? shape({ type: oneOf("solid"), color }, ["color"])(value, depth) : undefined;
};
const stroke = shape({ color, widthPt: num }, ["color", "widthPt"]);
const run = shape(
  {
    text: str,
    bold: yes,
    italic: yes,
    underline: yes,
    fontFamily: font,
    sizePt: positive,
    color,
    letterSpacingPt: num,
    linkHref: href,
    linkSlide: int,
    linkTooltip: str,
  },
  ["text"],
);
const paragraph = shape(
  {
    runs: arrayOf(run),
    align: oneOf("left", "center", "right", "justify"),
    bullet: oneOf("disc", "number"),
    level: int,
    bulletStyle: oneOf("disc", "circle", "square", "decimal", "lowerAlpha", "upperAlpha", "lowerRoman", "upperRoman"),
    bulletChar: str,
    bulletImageId: partPath,
    startAt: int,
    indentCm: num,
    hangingCm: num,
    lineSpacing: num,
    lineSpacingPt: num,
    lineSpacingRule: oneOf("exact"),
    spaceBeforePt: num,
    spaceAfterPt: num,
    pptxPPr: str,
    odpParaProps: str,
  },
  ["runs"],
);
const body = shape({
  autofit: oneOf("none", "shrink", "resizeShape"),
  fontScalePct: num,
  lineScalePct: num,
  anchor: oneOf("top", "middle", "bottom"),
  insetsCm: shape({ left: num, top: num, right: num, bottom: num }, ["left", "top", "right", "bottom"]),
  columns: shape({ count: int, gapCm: num }, ["count", "gapCm"]),
  wrap: no,
});

const BASE: Record<string, Check> = {
  id: str,
  xCm: num,
  yCm: num,
  widthCm: num,
  heightCm: num,
  rotationDeg: num,
  linkHref: href,
  linkSlide: int,
  linkTooltip: str,
};
const BASE_REQUIRED = ["kind", "xCm", "yCm", "widthCm", "heightCm"];

const cell = shape(
  {
    paragraphs: arrayOf(paragraph),
    colSpan: int,
    rowSpan: int,
    covered: yes,
    fill,
    borders: plain,
    paddingCm: num,
    verticalAlign: oneOf("top", "middle", "bottom"),
    pptxTcPr: str,
    odpCellStyleName: str,
  },
  ["paragraphs"],
);
const row = shape({ cells: arrayOf(cell), heightCm: num, isHeader: yes, odpRowStyleName: str }, ["cells"]);

const FRAME_SPECS: Record<string, [Record<string, Check>, string[]]> = {
  text: [
    {
      role: oneOf("title", "subtitle", "body", "other"),
      phType: str,
      phIdx: int,
      paragraphs: arrayOf(paragraph),
      fill,
      stroke,
      body,
      pptxBodyPr: str,
      odpBodyProps: str,
    },
    ["paragraphs"],
  ],
  image: [{ mediaId: partPath, stroke }, ["mediaId"]],
  shape: [
    {
      geometry: text,
      adjustments: arrayOf(num),
      fill,
      stroke,
      paragraphs: arrayOf(paragraph),
      body,
      pptxBodyPr: str,
      odpBodyProps: str,
    },
    ["geometry"],
  ],
  media: [
    {
      media: oneOf("video", "audio"),
      mediaId: partPath,
      linkTarget: href,
      posterMediaId: partPath,
      contentType: str,
      autoStart: yes,
      loop: yes,
      muted: yes,
      volume: num,
      fullScreen: yes,
      oversize: yes,
      stroke,
      sourceXml: str,
    },
    ["media"],
  ],
  group: [{ children: (value, depth) => arrayOf(frame)(value, depth), xml: str }, ["children"]],
  foreign: [{ xml: str, label: str, noBounds: yes }, ["xml"]],
  chart: [
    {
      chartType: text,
      categories: arrayOf(str),
      series: plain,
      options: plain,
      unsupportedChartXml: str,
      sourcePart: str,
      edited: yes,
      xml: str,
    },
    ["chartType", "categories", "series", "options"],
  ],
  table: [
    {
      rows: arrayOf(row),
      columnWidthsCm: arrayOf(num),
      firstRow: yes,
      bandRow: yes,
      styleId: str,
      pptxTblPr: str,
      pptxGraphicFramePr: str,
      odpTableStyleName: str,
      odpColumnStyleNames: plain,
    },
    ["rows", "columnWidthsCm"],
  ],
};

function frame(value: unknown, depth: number): unknown {
  if (!isRecord(value) || typeof value.kind !== "string" || !Object.hasOwn(FRAME_SPECS, value.kind)) {
    return undefined;
  }
  const [spec, required] = FRAME_SPECS[value.kind];
  return shape({ kind: oneOf(value.kind), ...BASE, ...spec }, [...BASE_REQUIRED, ...required])(value, depth);
}

const DIRECTIONS = ["left", "right", "up", "down"];
const slide = shape(
  {
    id: str,
    frames: arrayOf(frame),
    background: fill,
    notes: arrayOf(str),
    transition: shape(
      {
        type: oneOf("fade", "push", "wipe", "cover", "uncover", "split", "none"),
        direction: oneOf(...DIRECTIONS),
        durationMs: num,
        advanceOnClick: no,
        advanceAfterMs: num,
      },
      ["type"],
    ),
    animations: arrayOf(
      shape(
        {
          frameId: text,
          effect: oneOf("appear", "fadeIn", "flyIn", "wipeIn", "zoomIn", "fadeOut", "flyOut"),
          direction: oneOf(...DIRECTIONS),
          trigger: oneOf("click", "withPrevious", "afterPrevious"),
          durationMs: num,
          delayMs: num,
        },
        ["frameId", "effect", "trigger"],
      ),
    ),
    layoutKey: str,
    // Only the preserved keys the import report reads (what stripping costs).
    preserved: shape({ "p:transition": str, "p:timing": str, "draw:transition": str, anim: str }),
    fields: arrayOf(
      shape(
        {
          kind: oneOf("footer", "slideNumber", "date"),
          xCm: num,
          yCm: num,
          widthCm: num,
          heightCm: num,
          align: oneOf("left", "center", "right"),
          sizePt: num,
          color: (value) => (isSafeColor(value) ? value : undefined),
          fontFamily: str,
          text: str,
          wireType: str,
          wireStyleName: str,
        },
        ["kind", "xCm", "yCm", "widthCm", "heightCm"],
      ),
    ),
  },
  ["frames"],
);

const mediaPart = shape(
  {
    mediaId: partPath,
    contentType: (value) => (typeof value === "string" && /^(?:image|video|audio)\/[a-z0-9.+-]+$/i.test(value) ? value : undefined),
    data: (value) => (typeof value === "string" && /^[A-Za-z0-9+/=\s]*$/.test(value) ? value : undefined),
  },
  ["mediaId", "contentType", "data"],
);

const themeColors: Check = (value) => {
  if (!isRecord(value)) return undefined;
  const out: Record<string, string> = {};
  for (const slot of THEME_COLOR_SLOTS) {
    const hex = value[slot];
    if (typeof hex !== "string" || !/^#[0-9a-f]{6}$/i.test(hex)) {
      return undefined;
    }
    out[slot] = hex.toLowerCase();
  }
  return out;
};
const theme = shape(
  {
    name: str,
    colors: themeColors,
    fonts: shape({ major: str, minor: str }, ["major", "minor"]),
    colorMap: plain,
  },
  ["name", "colors", "fonts"],
);
const layout = shape(
  {
    key: text,
    kind: oneOf("title", "titleAndBody", "sectionHeader", "twoContent", "titleOnly", "blank", "other"),
    name: str,
    slots: arrayOf(
      shape(
        {
          role: oneOf("title", "subtitle", "body", "other"),
          phType: str,
          phIdx: int,
          xCm: num,
          yCm: num,
          widthCm: num,
          heightCm: num,
        },
        ["role", "xCm", "yCm", "widthCm", "heightCm"],
      ),
    ),
    masterName: str,
  },
  ["key", "kind", "name", "slots"],
);

/** Parses (and validates) the private slide flavour: the envelope label is
 * checked and every field rebuilt from the whitelist above. Null when the
 * payload is not a v1 SLIDES_MIME payload or holds no usable slide. */
export function parseSlideClipboardPayload(raw: string): SlideClipboardPayload | null {
  const envelope = parseClipboardEnvelope(raw, SLIDES_MIME, 1);
  if (!envelope) {
    return null;
  }
  const nonce = typeof envelope.nonce === "string" && /^[A-Za-z0-9-]{1,100}$/.test(envelope.nonce) ? envelope.nonce : undefined;
  const format = envelope.format === "pptx" || envelope.format === "odp" ? envelope.format : undefined;
  const width = positive(envelope.slideWidthCm, 0) as number | undefined;
  const height = positive(envelope.slideHeightCm, 0) as number | undefined;
  const slides = (arrayOf(slide)(envelope.slides, 0) as Slide[] | undefined) ?? [];
  if (nonce === undefined || format === undefined || width === undefined || height === undefined || slides.length === 0) {
    return null;
  }
  const payload: SlideClipboardPayload = {
    mimeType: SLIDES_MIME,
    version: 1,
    nonce,
    format,
    slideWidthCm: width,
    slideHeightCm: height,
    slides,
    media: (arrayOf(mediaPart)(envelope.media, 0) as MediaPart[] | undefined) ?? [],
  };
  const parsedTheme = theme(envelope.theme, 0) as DeckTheme | undefined;
  if (parsedTheme) {
    payload.theme = parsedTheme;
  }
  const layouts = arrayOf(layout)(envelope.layouts, 0) as SlideLayout[] | undefined;
  if (layouts && layouts.length > 0) {
    payload.layouts = layouts;
  }
  return payload;
}
