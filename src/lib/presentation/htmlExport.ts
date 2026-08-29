// Self-contained HTML export of a presentation (SPEC-088) — a PURE, DOM-free
// serialiser over the model, so it lives in src/lib and is unit-testable in
// Node. It turns a `PresentationModel` into a single `.html` document that
// opens from `file://` with no external requests: inline CSS + JS, images as
// `data:` URIs, system font stacks (the model's family first). The output
// reproduces the SPEC-079 stage render for every frame kind and embeds a static
// viewer script that reproduces SPEC-084 navigation, SPEC-085 transitions and
// SPEC-086 builds by reading per-slide `data-*` attributes — the script carries
// no per-deck code, only per-deck DATA in the DOM (so all untrusted content is
// HTML-escaped exactly once, at emission, through `esc`).
//
// Layering rule (presentation/boundary.spec.ts): imports only presentation
// siblings and src/lib/core — never a Vue component. The stage's render
// decisions (pt→px, 0.95cm/level bullet indent, marker glyphs) are reproduced
// here rather than imported from components/presentation/slideRender.ts.
import { cmToPx } from "../core/units";
import { mediaDataUri, type MediaPart } from "../core/media";
import { shapeDefinition, shapeSvgPath } from "../core/shapeCatalogue";
import {
  DEFAULT_PRESENTATION_TITLE,
  type GroupFrame,
  type PresentationModel,
  type ShapeFrame,
  type Slide,
  type SlideFrame,
  type SlideParagraph,
  type SlideRun,
  type TableFrame,
  type TextFrame,
} from "./model";
import {
  buildGroups,
  effectiveAnimationDelayMs,
  effectiveAnimationDirection,
  effectiveAnimationDurationMs,
} from "./animations";
import { effectiveDirection, effectiveDurationMs } from "./transitions";
import { advanceKeys, backKeys, firstKeys, lastKeys } from "./playbackKeys";

/** pt → px at 96/72 — the ratio the stage renderer uses. */
const PT_TO_PX = 96 / 72;
/** Bullet indent per level, in cm (SPEC-079). */
const LEVEL_INDENT_CM = 0.95;
/** A run with no explicit size renders at this point size. */
const DEFAULT_FONT_SIZE_PT = 18;

/** The system font stack the export falls back to — no web fonts are embedded
 * (AC-7), so a run's own family (if any) sits in front of these. */
const SYSTEM_STACK = "system-ui, -apple-system, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif";

// --- escaping / validation --------------------------------------------------

/** HTML-escapes text and attribute values — the ONE place untrusted content is
 * neutralised (AC-48/49). Quotes are escaped so the same helper serves attribute
 * values; `'` too so a value can sit in a single-quoted CSS string. */
export function esc(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

/** A `#rrggbb` colour re-validated (AC-48): the lower-cased hex when it matches,
 * else `undefined`. Untrusted colour strings never reach the CSS otherwise. */
export function validColor(color: string | undefined): string | undefined {
  return color && /^#[0-9a-fA-F]{6}$/.test(color) ? color.toLowerCase() : undefined;
}

/** A font-family name reduced to a safe token set (letters, digits, spaces,
 * hyphens) so it cannot break out of the CSS string it lands in (AC-48). */
function safeFamily(family: string): string {
  return family.replace(/[^A-Za-z0-9 -]/g, "").trim();
}

/** A number rounded to 3 decimals for a compact, deterministic style string. */
function num(value: number): string {
  const rounded = Math.round(value * 1000) / 1000;
  return String(Object.is(rounded, -0) ? 0 : rounded);
}

const px = (cm: number): string => `${num(cmToPx(cm))}px`;

// --- text rendering (mirrors slideRender.ts's decisions) --------------------

/** A frame's paragraphs annotated with their bullet marker glyph — numbered
 * bullets restart per contiguous same-level run (reproduced from the stage's
 * `markedParagraphs`). */
function markedParagraphs(paragraphs: SlideParagraph[]): { para: SlideParagraph; marker: string }[] {
  const counts: Record<number, number> = {};
  let prev: SlideParagraph | undefined;
  return paragraphs.map((para) => {
    let marker = "";
    if (para.bullet === "number") {
      const level = para.level ?? 0;
      if (prev && prev.bullet === "number" && (prev.level ?? 0) === level) {
        counts[level] = (counts[level] ?? 0) + 1;
      } else {
        counts[level] = 1;
      }
      marker = `${counts[level]}.`;
    } else if (para.bullet === "disc") {
      marker = "•";
    }
    prev = para;
    return { para, marker };
  });
}

function runStyle(run: SlideRun): string {
  const parts: string[] = [];
  if (run.bold) parts.push("font-weight:700");
  if (run.italic) parts.push("font-style:italic");
  if (run.underline) parts.push("text-decoration:underline");
  const family = run.fontFamily ? safeFamily(run.fontFamily) : "";
  if (family) parts.push(`font-family:'${family}', ${SYSTEM_STACK}`);
  parts.push(`font-size:${num((run.sizePt ?? DEFAULT_FONT_SIZE_PT) * PT_TO_PX)}px`);
  parts.push(`color:${validColor(run.color) ?? "#000000"}`);
  return parts.join(";");
}

const ALIGNMENTS = new Set(["left", "center", "right", "justify"]);

function renderParagraph(entry: { para: SlideParagraph; marker: string }, centered: boolean): string {
  const align = centered ? "center" : entry.para.align && ALIGNMENTS.has(entry.para.align) ? entry.para.align : "left";
  const styleParts = [`text-align:${align}`];
  if (!centered && entry.para.level) {
    styleParts.push(`margin-left:${px(entry.para.level * LEVEL_INDENT_CM)}`);
  }
  const marker = entry.marker && !centered ? `<span class="marker" aria-hidden="true">${esc(entry.marker)} </span>` : "";
  const runs = entry.para.runs.map((run) => `<span style="${esc(runStyle(run))}">${esc(run.text)}</span>`).join("");
  return `<div class="para" style="${esc(styleParts.join(";"))}">${marker}${runs}</div>`;
}

// --- frame positioning ------------------------------------------------------

function positionStyle(frame: SlideFrame): string[] {
  const parts = [`left:${px(frame.xCm)}`, `top:${px(frame.yCm)}`, `width:${px(frame.widthCm)}`, `height:${px(frame.heightCm)}`];
  if (frame.rotationDeg) {
    parts.push(`transform:rotate(${num(frame.rotationDeg)}deg)`, "transform-origin:center");
  }
  return parts;
}

// --- shapes -----------------------------------------------------------------

function renderShape(frame: ShapeFrame, fid: string): string {
  const style = positionStyle(frame);
  const w = cmToPx(frame.widthCm);
  const h = cmToPx(frame.heightCm);
  const def = shapeDefinition(frame.geometry);
  const d = shapeSvgPath(frame.geometry, w, h, frame.adjustments);
  const fill = frame.fill && frame.fill.type === "solid" ? validColor(frame.fill.color) : undefined;
  const stroke = frame.stroke ? validColor(frame.stroke.color) : undefined;
  const pathAttrs: string[] = [`d="${esc(d)}"`, `fill="${fill ?? "none"}"`, 'fill-rule="evenodd"'];
  if (stroke) {
    pathAttrs.push(`stroke="${stroke}"`, `stroke-width="${num((frame.stroke?.widthPt ?? 1) * PT_TO_PX)}"`);
  } else if (def?.isLine) {
    pathAttrs.push('stroke="#000000"', 'stroke-width="1"');
  }
  const svg =
    `<svg class="shape-svg" width="${num(w)}" height="${num(h)}" viewBox="0 0 ${num(w)} ${num(h)}" ` +
    `preserveAspectRatio="none" aria-hidden="true"><path ${pathAttrs.join(" ")}/></svg>`;
  const paragraphs = frame.paragraphs && frame.paragraphs.length > 0
    ? `<div class="shape-text">${markedParagraphs(frame.paragraphs).map((e) => renderParagraph(e, true)).join("")}</div>`
    : "";
  return `<div class="frame shape" data-fid="${fid}" style="${esc(style.join(";"))}">${svg}${paragraphs}</div>`;
}

// --- per-slide frame emission -----------------------------------------------

interface EmitContext {
  media: Map<string, MediaPart>;
  fidOf: Map<string, string>;
  counter: { n: number };
}

function nextFid(ctx: EmitContext, frame: SlideFrame): string {
  const fid = `f${ctx.counter.n++}`;
  ctx.fidOf.set(frame.id, fid);
  return fid;
}

function renderTextFrame(frame: TextFrame, fid: string): string {
  const style = positionStyle(frame);
  const fill = frame.fill && frame.fill.type === "solid" ? validColor(frame.fill.color) : undefined;
  if (fill) style.push(`background:${fill}`);
  const body = markedParagraphs(frame.paragraphs).map((e) => renderParagraph(e, false)).join("");
  return `<div class="frame text" data-fid="${fid}" style="${esc(style.join(";"))}">${body}</div>`;
}

function renderImageFrame(frame: Extract<SlideFrame, { kind: "image" }>, fid: string, ctx: EmitContext): string {
  const style = positionStyle(frame);
  const src = mediaDataUri(ctx.media.get(frame.mediaId));
  return `<div class="frame image" data-fid="${fid}" style="${esc(style.join(";"))}"><img alt="" src="${esc(src)}"></div>`;
}

function renderGroupFrame(frame: GroupFrame, fid: string, ctx: EmitContext): string {
  const children = frame.children.map((child) => renderFrame(child, ctx)).join("");
  return `<div class="frame group" data-fid="${fid}" style="position:absolute">${children}</div>`;
}

function renderForeignFrame(frame: Extract<SlideFrame, { kind: "foreign" }>, fid: string): string {
  if (frame.noBounds) {
    return "";
  }
  const style = positionStyle(frame);
  return `<div class="frame foreign" data-fid="${fid}" style="${esc(style.join(";"))}"><span class="foreign-label">${esc(frame.label)}</span></div>`;
}

function renderTableFrame(frame: TableFrame, fid: string): string {
  const style = positionStyle(frame);
  const cols = frame.columnWidthsCm
    .map((cm) => `<col style="width:${px(cm)}">`)
    .join("");
  const rows = frame.rows
    .map((row) => {
      const cells = row.cells
        .map((cell) => {
          if (cell.covered) return "";
          const spanAttrs = [
            cell.colSpan && cell.colSpan > 1 ? ` colspan="${cell.colSpan}"` : "",
            cell.rowSpan && cell.rowSpan > 1 ? ` rowspan="${cell.rowSpan}"` : "",
          ].join("");
          const cellStyle: string[] = [];
          const fill = cell.fill && cell.fill.type === "solid" ? validColor(cell.fill.color) : undefined;
          if (fill) cellStyle.push(`background:${fill}`);
          cellStyle.push(`vertical-align:${cell.verticalAlign ?? "top"}`);
          const body = markedParagraphs(cell.paragraphs).map((e) => renderParagraph(e, false)).join("");
          return `<td${spanAttrs} style="${esc(cellStyle.join(";"))}">${body}</td>`;
        })
        .join("");
      return `<tr>${cells}</tr>`;
    })
    .join("");
  return (
    `<div class="frame table" data-fid="${fid}" style="${esc(style.join(";"))}">` +
    `<table class="slide-table" style="border-collapse:collapse;width:100%;height:100%">` +
    `<colgroup>${cols}</colgroup><tbody>${rows}</tbody></table></div>`
  );
}

function renderFrame(frame: SlideFrame, ctx: EmitContext): string {
  const fid = nextFid(ctx, frame);
  switch (frame.kind) {
    case "text":
      return renderTextFrame(frame, fid);
    case "image":
      return renderImageFrame(frame, fid, ctx);
    case "shape":
      return renderShape(frame, fid);
    case "group":
      return renderGroupFrame(frame, fid, ctx);
    case "foreign":
      return renderForeignFrame(frame, fid);
    case "table":
      return renderTableFrame(frame, fid);
    case "chart":
      return renderChartFrame(frame, fid);
  }
}

/** A chart on an exported slide (SPEC-092). This exporter has no canvas, so a
 * chart renders as a positioned placeholder box for now; SPEC-088 replaces it
 * with a `core/chartCanvas.ts`-painted image. */
function renderChartFrame(frame: Extract<SlideFrame, { kind: "chart" }>, fid: string): string {
  const style = positionStyle(frame);
  return `<div class="frame chart" data-fid="${fid}" style="${esc(style.join(";"))}"><span class="foreign-label">Chart</span></div>`;
}

// --- animation data ---------------------------------------------------------

interface AnimData {
  fid: string;
  effect: string;
  dir: string;
  ms: number;
  delay: number;
  trigger: string;
  build: number;
  order: number;
}

/** The slide's animations flattened to the viewer's data shape: each mapped to
 * its frame's emitted `data-fid`, tagged with its build-group index and global
 * play order (SPEC-086's `buildGroups`). Animations whose frame did not emit
 * (e.g. a `noBounds` foreign target) are dropped. */
function slideAnimData(slide: Slide, fidOf: Map<string, string>): AnimData[] {
  const groups = buildGroups(slide.animations ?? []);
  const out: AnimData[] = [];
  let order = 0;
  groups.forEach((group, build) => {
    for (const animation of group) {
      const fid = fidOf.get(animation.frameId);
      const index = order++;
      if (!fid) {
        continue;
      }
      out.push({
        fid,
        effect: animation.effect,
        dir: effectiveAnimationDirection(animation) ?? "left",
        ms: effectiveAnimationDurationMs(animation),
        delay: effectiveAnimationDelayMs(animation),
        trigger: index === 0 ? "click" : animation.trigger,
        build,
        order: index,
      });
    }
  });
  return out;
}

// --- slide sections ---------------------------------------------------------

function surfaceBackground(slide: Slide): string {
  const bg = slide.background;
  return (bg && bg.type === "solid" ? validColor(bg.color) : undefined) ?? "#ffffff";
}

function renderSlide(slide: Slide, index: number, widthPx: number, heightPx: number, ctx: EmitContext): string {
  ctx.fidOf = new Map();
  const frames = slide.frames.map((frame) => renderFrame(frame, ctx)).join("");
  const anims = slideAnimData(slide, ctx.fidOf);
  const builds = buildGroups(slide.animations ?? []).length;

  const transition = slide.transition;
  const type = transition && transition.type ? transition.type : "none";
  const dir = transition && transition.type !== "none" ? effectiveDirection(transition) ?? "" : "";
  const ms = transition && transition.type !== "none" ? effectiveDurationMs(transition) : 0;
  const advanceAfter = transition && transition.advanceAfterMs && transition.advanceAfterMs > 0 ? Math.round(transition.advanceAfterMs) : 0;

  const attrs = [
    `class="slide"`,
    `id="s${index + 1}"`,
    `data-index="${index}"`,
    `data-builds="${builds}"`,
    `data-transition="${type}"`,
    `data-transition-dir="${dir}"`,
    `data-transition-ms="${ms}"`,
    `data-advance-after="${advanceAfter}"`,
    `data-anim="${esc(JSON.stringify(anims))}"`,
    `style="width:${num(widthPx)}px;height:${num(heightPx)}px;background:${surfaceBackground(slide)}"`,
    `aria-hidden="true"`,
    `inert`,
  ];
  return `<section ${attrs.join(" ")}>${frames}</section>`;
}

// --- notes ------------------------------------------------------------------

function renderNotes(slides: Slide[]): string {
  const blocks = slides
    .map((slide, index) => {
      const paras = (slide.notes ?? []).map((line) => `<p>${esc(line)}</p>`).join("");
      return `<div class="note" data-index="${index}">${paras}</div>`;
    })
    .join("");
  return `<aside id="notes" hidden aria-hidden="true"><div id="notes-body">${blocks}</div></aside>`;
}

// --- key table (shared with playbackKeys.ts, asserted equal by the spec) -----

/** The navigation key tables embedded into the viewer, byte-for-byte the
 * `playbackKeys.ts` arrays — the export's single source, asserted equal by the
 * unit test so the two surfaces cannot drift (SPEC-088 Reference). */
export function keyTableJson(): string {
  return JSON.stringify({ advance: advanceKeys, back: backKeys, first: firstKeys, last: lastKeys });
}

// --- the download file name (twin of pptxFileName) --------------------------

/** `<title>.html`, sanitised the way `pptxFileName` sanitises: a blank title
 * falls back to the default presentation title. */
export function htmlFileName(title: string): string {
  const trimmed = title.trim();
  return `${trimmed.length > 0 ? trimmed : DEFAULT_PRESENTATION_TITLE}.html`;
}

// --- the document -----------------------------------------------------------

export interface HtmlExportOptions {
  includeNotes: boolean;
}

/** Serialises a presentation to a single self-contained HTML document (see the
 * module header). Pure — never mutates the model. Throws on a deck with no
 * slides (the menu item is disabled; this guard is for API callers). */
export function exportPresentationHtml(model: PresentationModel, options: HtmlExportOptions): string {
  if (model.slides.length === 0) {
    throw new Error("Cannot export a presentation with no slides");
  }
  const widthPx = cmToPx(model.slideWidthCm);
  const heightPx = cmToPx(model.slideHeightCm);
  const ctx: EmitContext = {
    media: new Map((model.media ?? []).map((part) => [part.mediaId, part])),
    fidOf: new Map(),
    counter: { n: 0 },
  };
  const sections = model.slides.map((slide, index) => renderSlide(slide, index, widthPx, heightPx, ctx)).join("");
  const title = esc(model.title || DEFAULT_PRESENTATION_TITLE);
  const notes = options.includeNotes ? renderNotes(model.slides) : "";

  return (
    `<!doctype html><html lang="en"><head><meta charset="utf-8">` +
    `<meta name="viewport" content="width=device-width, initial-scale=1">` +
    `<title>${title}</title><style>${STYLESHEET}</style></head>` +
    `<body data-docu-edit-export="1">` +
    `<main id="stage" role="region" aria-roledescription="slide show" aria-label="${title}" ` +
    `data-slide-w="${num(widthPx)}" data-slide-h="${num(heightPx)}" data-notes="${options.includeNotes ? "1" : "0"}">` +
    sections +
    `</main>` +
    `<div id="hud" aria-hidden="true"><span id="hud-counter"></span></div>` +
    `<div id="end" hidden role="alert">End of slide show — click or press any key to exit</div>` +
    notes +
    `<div id="live" class="sr-only" aria-live="polite"></div>` +
    `<script>${viewerScript()}</script>` +
    `</body></html>`
  );
}

// The static stylesheet (readable, unminified — a feature per the spec). It
// carries only per-TYPE knowledge; every per-deck value is an inline style or a
// data-* attribute the viewer reads.
const STYLESHEET = `
*{box-sizing:border-box}
html,body{margin:0;padding:0;height:100%;background:#000;overflow:hidden;font-family:${SYSTEM_STACK}}
#stage{position:absolute;top:50%;left:50%;transform-origin:center center}
.slide{position:absolute;top:0;left:0;overflow:hidden;display:none}
.slide.current{display:block}
.frame{position:absolute;box-sizing:border-box;overflow:visible}
.frame.text{display:flex;flex-direction:column;justify-content:flex-start}
.para{margin:0;line-height:1.2;white-space:pre-wrap;overflow-wrap:anywhere}
.marker{color:inherit}
.frame.image img{width:100%;height:100%;object-fit:fill;display:block}
.shape-svg{position:absolute;top:0;left:0;overflow:visible}
.shape-text{position:absolute;inset:0;display:flex;flex-direction:column;align-items:center;justify-content:center;pointer-events:none}
.frame.group{left:0!important;top:0!important;width:0!important;height:0!important;overflow:visible}
.frame.foreign{border:1px dashed #b0b0b0;display:flex;align-items:center;justify-content:center}
.foreign-label{color:#666;font-size:.75rem}
#hud{position:fixed;bottom:12px;right:16px;color:#fff;opacity:.55;font-size:14px;pointer-events:none}
#end{position:fixed;inset:0;display:flex;align-items:center;justify-content:center;background:#000;color:#fff;font-size:20px;text-align:center;padding:24px}
#end[hidden]{display:none}
#notes{position:fixed;left:0;right:0;bottom:0;max-height:40%;overflow:auto;background:rgba(0,0,0,.85);color:#fff;padding:12px 20px;font-size:16px;line-height:1.4}
#notes p{margin:.25em 0}
#notes[hidden]{display:none}
.sr-only{position:absolute;width:1px;height:1px;padding:0;margin:-1px;overflow:hidden;clip:rect(0 0 0 0);white-space:nowrap;border:0}
`;

// The viewer script — one static template for every deck (no per-deck code). It
// reads the sections' data-* attributes and reproduces SPEC-084 navigation,
// SPEC-085 transitions and SPEC-086 builds. Written with string concatenation
// (no backticks / ${}) so it embeds cleanly inside this module's own template.
function viewerScript(): string {
  return `(function(){
"use strict";
var KEYS=${keyTableJson()};
var SYSTEM_STACK=${JSON.stringify(SYSTEM_STACK)};
var stage=document.getElementById("stage");
var sections=[].slice.call(stage.querySelectorAll("section.slide"));
var N=sections.length;
var hud=document.getElementById("hud-counter");
var live=document.getElementById("live");
var endEl=document.getElementById("end");
var notesEl=document.getElementById("notes");
var notesBody=document.getElementById("notes-body");
var hasNotes=stage.getAttribute("data-notes")==="1";
var surfaceW=parseFloat(stage.getAttribute("data-slide-w"))||1;
var surfaceH=parseFloat(stage.getAttribute("data-slide-h"))||1;
stage.style.width=surfaceW+"px";stage.style.height=surfaceH+"px";
var mq=window.matchMedia?window.matchMedia("(prefers-reduced-motion: reduce)"):{matches:false,addEventListener:function(){}};
function reduced(){return !!mq.matches;}

// ---- geometry helpers (ported from transitionStyles.ts / animationStyles.ts)
function enterTranslate(d){return d==="left"?"translateX(100%)":d==="right"?"translateX(-100%)":d==="up"?"translateY(100%)":"translateY(-100%)";}
function exitTranslate(d){return d==="left"?"translateX(-100%)":d==="right"?"translateX(100%)":d==="up"?"translateY(-100%)":"translateY(100%)";}
function wipeEnterClip(d){return d==="left"?"inset(0 0 0 100%)":d==="right"?"inset(0 100% 0 0)":d==="up"?"inset(100% 0 0 0)":"inset(0 0 100% 0)";}
function stagePair(type,dir,phase){var enter=phase==="enter";
 if(type==="fade")return{o:{},i:{opacity:enter?"0":"1"},top:true};
 if(type==="push")return{o:{transform:enter?"translateX(0)":exitTranslate(dir)},i:{transform:enter?enterTranslate(dir):"translateX(0)"},top:true};
 if(type==="cover")return{o:{},i:{transform:enter?enterTranslate(dir):"translateX(0)"},top:true};
 if(type==="uncover")return{o:{transform:enter?"translateX(0)":exitTranslate(dir)},i:{},top:false};
 if(type==="wipe")return{o:{},i:{clipPath:enter?wipeEnterClip(dir):"inset(0 0 0 0)"},top:true};
 if(type==="split")return{o:{},i:{clipPath:enter?"inset(0 50% 0 50%)":"inset(0 0 0 0)"},top:true};
 return{o:{},i:{},top:true};}
var FLY={left:"translateX(-120%)",right:"translateX(120%)",up:"translateY(-120%)",down:"translateY(120%)"};
var WIPEH={left:"inset(0 100% 0 0)",right:"inset(0 0 0 100%)",up:"inset(0 0 100% 0)",down:"inset(100% 0 0 0)"};
function fromStyle(a){var d=a.dir||"left";switch(a.effect){
 case"appear":return{visibility:"visible"};
 case"fadeIn":return{visibility:"visible",opacity:"0"};
 case"fadeOut":return{visibility:"visible",opacity:"1"};
 case"flyIn":return{visibility:"visible",opacity:"1",transform:FLY[d]};
 case"flyOut":return{visibility:"visible",opacity:"1",transform:"none"};
 case"wipeIn":return{visibility:"visible",opacity:"1",clipPath:WIPEH[d]};
 case"zoomIn":return{visibility:"visible",opacity:"0",transform:"scale(0.2)"};
 default:return{visibility:"visible"};}}
function toStyle(a){var d=a.dir||"left";switch(a.effect){
 case"appear":return{visibility:"visible"};
 case"fadeIn":return{visibility:"visible",opacity:"1"};
 case"fadeOut":return{visibility:"visible",opacity:"0"};
 case"flyIn":return{visibility:"visible",opacity:"1",transform:"none"};
 case"flyOut":return{visibility:"visible",opacity:"0",transform:FLY[d]};
 case"wipeIn":return{visibility:"visible",opacity:"1",clipPath:"inset(0)"};
 case"zoomIn":return{visibility:"visible",opacity:"1",transform:"none"};
 default:return{visibility:"visible"};}}
function transitionFor(a){var props=(a.effect==="fadeIn"||a.effect==="fadeOut")?["opacity"]:a.effect==="wipeIn"?["clip-path"]:a.effect==="zoomIn"?["opacity","transform"]:["transform","opacity"];
 return props.map(function(p){return p+" "+a.ms+"ms ease";}).join(", ");}
function isEntrance(e){return e==="appear"||e==="fadeIn"||e==="flyIn"||e==="wipeIn"||e==="zoomIn";}
function playbackDuration(a){return a.effect==="appear"?0:a.ms;}

// ---- per-section state
function anims(sec){try{return JSON.parse(sec.getAttribute("data-anim")||"[]");}catch(e){return[];}}
function buildCount(i){return parseInt(sections[i].getAttribute("data-builds"),10)||0;}
function frameEl(sec,fid){return sec.querySelector('[data-fid="'+fid+'"]');}
function clearFrameStyle(el){el.style.opacity="";el.style.visibility="";el.style.transform="";el.style.clipPath="";el.style.transition="";}
function applyStyle(el,s){el.style.transition=s.transition!==undefined?s.transition:"";
 el.style.opacity=s.opacity!==undefined?s.opacity:"";
 el.style.visibility=s.visibility!==undefined?s.visibility:"";
 el.style.transform=s.transform!==undefined?s.transform:"";
 el.style.clipPath=s.clipPath!==undefined?s.clipPath:"";}
function isHiddenAt(list,fid,groupsPlayed){var hidden=false,i;
 for(i=0;i<list.length;i++){if(list[i].fid===fid&&isEntrance(list[i].effect)){hidden=true;break;}}
 for(i=0;i<list.length;i++){var a=list[i];if(a.build<groupsPlayed&&a.fid===fid){hidden=!isEntrance(a.effect);}}
 return hidden;}
function settle(sec,groupsPlayed){var list=anims(sec);
 var frames=[].slice.call(sec.querySelectorAll("[data-fid]"));
 frames.forEach(function(el){clearFrameStyle(el);
  var fid=el.getAttribute("data-fid");
  if(isHiddenAt(list,fid,groupsPlayed)){el.style.opacity="0";el.style.visibility="hidden";}});}

// ---- build engine (per current slide)
var buildTimers=[];var pendingBuild=-1;var buildLive=false;
function clearBuildTimers(){buildTimers.forEach(function(t){clearTimeout(t);});buildTimers=[];}
function raf(cb){if(window.requestAnimationFrame)window.requestAnimationFrame(function(){window.requestAnimationFrame(cb);});else setTimeout(cb,16);}
function playBuild(sec,groupIndex,done){clearBuildTimers();
 if(reduced()){settle(sec,groupIndex+1);if(done)done();return;}
 var list=anims(sec).filter(function(a){return a.build===groupIndex;}).sort(function(a,b){return a.order-b.order;});
 settle(sec,groupIndex);
 // schedule via groupTimeline maths
 var prevStart=0,prevDur=0,total=0;var i;
 var sched=[];
 for(i=0;i<list.length;i++){var a=list[i];var dur=playbackDuration(a);var delay=a.delay;var start;
  if(i===0)start=delay;else if(a.trigger==="withPrevious")start=prevStart+delay;else start=prevStart+prevDur+delay;
  sched.push({a:a,start:start,dur:dur});prevStart=start;prevDur=dur;total=Math.max(total,start+dur);}
 buildLive=true;pendingBuild=groupIndex+1;
 sched.forEach(function(entry){var el=frameEl(sec,entry.a.fid);if(!el)return;
  buildTimers.push(setTimeout(function(){applyStyle(el,Object.assign({transition:"none"},fromStyle(entry.a)));
   raf(function(){var t=toStyle(entry.a);t.transition=entry.a.effect==="appear"?"none":transitionFor(entry.a);applyStyle(el,t);});
  },entry.start));});
 buildTimers.push(setTimeout(function(){settle(sec,groupIndex+1);buildLive=false;if(done)done();},total+40));}
function completeBuild(){if(!buildLive)return false;clearBuildTimers();settle(sections[state.slide],pendingBuild);buildLive=false;return true;}

// ---- transition engine
var trans=null;
function clearTransitionStyle(sec){sec.style.transform="";sec.style.opacity="";sec.style.clipPath="";sec.style.transition="";sec.style.zIndex="";}
function runTransition(from,to,type,dir,ms,done){
 var pair=stagePair(type,dir,"enter");
 var outEl=sections[from],inEl=sections[to];
 outEl.classList.add("current");inEl.classList.add("current");
 outEl.style.zIndex=pair.top?"0":"1";inEl.style.zIndex=pair.top?"1":"0";
 var tr="transform "+ms+"ms ease, opacity "+ms+"ms ease, clip-path "+ms+"ms ease";
 applyStyle(outEl,Object.assign({transition:"none"},pair.o));applyStyle(inEl,Object.assign({transition:"none"},pair.i));
 void inEl.offsetWidth;
 var fired=false;var timer=null;
 function finish(){if(fired)return;fired=true;if(timer)clearTimeout(timer);
  inEl.removeEventListener("transitionend",finish);
  clearTransitionStyle(outEl);clearTransitionStyle(inEl);
  outEl.classList.remove("current");trans=null;done();}
 raf(function(){var act=stagePair(type,dir,"active");
  applyStyle(outEl,Object.assign({transition:tr},act.o));applyStyle(inEl,Object.assign({transition:tr},act.i));
  inEl.addEventListener("transitionend",finish);timer=setTimeout(finish,ms+100);});
 trans={finish:finish};}
function finishTransition(){if(trans)trans.finish();}

// ---- navigation state machine
var state={slide:0,build:0};
var mode="playing"; // playing | endscreen | black
var autoTimer=null;
function clearAuto(){if(autoTimer){clearTimeout(autoTimer);autoTimer=null;}}
function scheduleAuto(){clearAuto();if(mode!=="playing")return;
 var ms=parseInt(sections[state.slide].getAttribute("data-advance-after"),10)||0;
 if(ms>0)autoTimer=setTimeout(function(){advance();},ms);}
function showOnly(index){for(var i=0;i<N;i++){var on=i===index;sections[i].classList.toggle("current",on);
  sections[i].setAttribute("aria-hidden",on?"false":"true");if(on)sections[i].removeAttribute("inert");else sections[i].setAttribute("inert","");}}
function updateUI(){hud.textContent=(state.slide+1)+" / "+N;live.textContent="Slide "+(state.slide+1)+" of "+N;
 if(notesEl&&!notesEl.hasAttribute("hidden"))renderNote();}
function renderNote(){if(!hasNotes)return;var blocks=notesBody?notesBody.children:[];
 for(var i=0;i<blocks.length;i++){blocks[i].style.display=(parseInt(blocks[i].getAttribute("data-index"),10)===state.slide)?"block":"none";}}
function land(index,build){showOnly(index);settle(sections[index],build);state.slide=index;state.build=build;mode="playing";
 endEl.setAttribute("hidden","");updateUI();syncHash();scheduleAuto();}

function enterEnd(){mode="endscreen";clearAuto();endEl.removeAttribute("hidden");live.textContent="";}
function enterBlack(){mode="black";endEl.setAttribute("hidden","");showOnly(-1);live.textContent="";}
function restart(){land(0,0);}

function advance(){clearAuto();
 if(trans){finishTransition();return;}
 if(completeBuild())return;
 if(mode==="endscreen"){enterBlack();return;}
 if(mode==="black"){restart();return;}
 if(state.build<buildCount(state.slide)){playBuild(sections[state.slide],state.build,null);state.build++;updateUI();return;}
 if(state.slide<N-1){var from=state.slide;var to=from+1;state.build=0;
  var sec=sections[to];var type=sec.getAttribute("data-transition");var dir=sec.getAttribute("data-transition-dir")||"left";var ms=parseInt(sec.getAttribute("data-transition-ms"),10)||0;
  state.slide=to;
  if(type&&type!=="none"&&!reduced()&&ms>0){settle(sec,0);sec.setAttribute("aria-hidden","false");sec.removeAttribute("inert");
   runTransition(from,to,type,dir,ms,function(){showOnly(to);updateUI();syncHash();scheduleAuto();});
   updateUI();syncHash();}
  else{land(to,0);}
  return;}
 enterEnd();}
function back(){clearAuto();if(trans)finishTransition();
 if(mode==="black"){land(N-1,buildCount(N-1));return;}
 if(mode==="endscreen"){land(N-1,buildCount(N-1));return;}
 if(state.slide<=0){land(0,buildCount(0)>0?state.build:0);return;}
 var prev=state.slide-1;land(prev,buildCount(prev));}
function jumpTo(oneBased){clearAuto();if(trans)finishTransition();var idx=oneBased-1;if(idx<0)idx=0;if(idx>N-1)idx=N-1;land(idx,0);}

// ---- scaling
function rescale(){var w=window.innerWidth,h=window.innerHeight;if(w<=0||h<=0)return;
 var s=Math.min(w/surfaceW,h/surfaceH);stage.style.transform="translate(-50%,-50%) scale("+s+")";}

// ---- hash
function syncHash(){if(window.history&&window.history.replaceState){try{window.history.replaceState(null,"","#"+(state.slide+1));}catch(e){}}}
function startSlideFromHash(){var h=(location.hash||"").replace(/^#/,"");var n=parseInt(h,10);if(isFinite(n)&&n>=1){return Math.min(N,Math.max(1,n))-1;}return 0;}

// ---- input
function normaliseKey(k){return k.length===1?k.toLowerCase():k;}
function actionForKey(k){k=normaliseKey(k);if(KEYS.advance.indexOf(k)>=0)return"advance";if(KEYS.back.indexOf(k)>=0)return"back";if(KEYS.first.indexOf(k)>=0)return"first";if(KEYS.last.indexOf(k)>=0)return"last";return null;}
var digitBuffer="";var digitTimer=null;
function clearDigits(){digitBuffer="";if(digitTimer){clearTimeout(digitTimer);digitTimer=null;}}
document.addEventListener("keydown",function(e){
 if(e.metaKey||e.ctrlKey||e.altKey)return;
 var t=e.target;if(t&&(t.tagName==="INPUT"||t.tagName==="TEXTAREA"||t.isContentEditable))return;
 if(hasNotes&&(e.key==="v"||e.key==="V")){e.preventDefault();toggleNotes();return;}
 if(e.key>="0"&&e.key<="9"){e.preventDefault();digitBuffer+=e.key;if(digitTimer)clearTimeout(digitTimer);digitTimer=setTimeout(clearDigits,3000);return;}
 if(e.key==="Enter"&&digitBuffer.length>0){e.preventDefault();var target=parseInt(digitBuffer,10);clearDigits();jumpTo(target);return;}
 var action=actionForKey(e.key);
 if(mode!=="playing"){e.preventDefault();if(action==="back")back();else advance();clearDigits();return;}
 if(!action)return;e.preventDefault();clearDigits();
 if(action==="advance")advance();else if(action==="back")back();else if(action==="first")jumpTo(1);else if(action==="last")jumpTo(N);
});
// pointer / touch
var SWIPE=48;var startX=0,startY=0,swiped=false;
stage.addEventListener("pointerdown",function(e){startX=e.clientX;startY=e.clientY;swiped=false;});
window.addEventListener("pointerup",function(e){var dx=e.clientX-startX,dy=e.clientY-startY;
 if(Math.abs(dx)>=SWIPE&&Math.abs(dx)>Math.abs(dy)){swiped=true;if(dx<0)advance();else back();}});
document.addEventListener("click",function(){if(swiped){swiped=false;return;}if(mode!=="playing"){advance();return;}advance();});
document.addEventListener("visibilitychange",function(){if(document.hidden)clearAuto();else scheduleAuto();});
window.addEventListener("resize",rescale);
mq.addEventListener&&mq.addEventListener("change",function(){});

// ---- notes drawer
function toggleNotes(){if(!hasNotes)return;if(notesEl.hasAttribute("hidden")){notesEl.removeAttribute("hidden");notesEl.setAttribute("aria-hidden","false");renderNote();}else{notesEl.setAttribute("hidden","");notesEl.setAttribute("aria-hidden","true");}}

// ---- boot
rescale();
var startIndex=startSlideFromHash();
land(startIndex,0);
})();`;
}
