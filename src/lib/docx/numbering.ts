// WordprocessingML numbering (SPEC-018): the numbering.xml catalog parser
// the reader resolves w:numPr references against, and the writer-side plan
// that turns the model's adjacency-defined lists back into abstractNum/num
// definitions.
//
// Mapping choices that keep the round trip exact:
//  - One w:num instance per contiguous run of list paragraphs. OOXML counts
//    numbering per num instance, which is exactly the model's restart
//    semantics (AC-11): a run interrupted by a plain paragraph is two runs,
//    hence two nums, hence a restart.
//  - abstractNum definitions are shared between runs with identical level
//    definitions, so identical lists don't duplicate them.
//  - Levels carry the model's values (marker style -> w:numFmt/w:lvlText,
//    startAt -> w:start, indentCm -> w:ind/@w:left), and the reader stores
//    a field only where it differs from the editor default, so an
//    export -> import cycle reproduces the model bit-for-bit.
import type { ListFormatting, ListMarkerStyle, ListType } from "../documentModel";
import {
  bulletStyleForChar,
  clampListLevel,
  defaultMarkerStyle,
  MAX_LIST_LEVEL,
  planListLevels,
  type PlannedListLevel,
} from "../lists";
import type { DocumentModel } from "../documentModel";
import { firstChildElement, parseXml, serializeElement, wordAttr, WORD_NAMESPACE } from "./xml";
import { cmToTwips, twipsToCm } from "./units";

export const NUMBERING_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";

export const NUMBERING_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";

/** The marker box's hanging indent in twips (0.635 cm). */
const HANGING_TWIPS = 360;

// w:numFmt values for the model's number styles, and back.
const NUM_FMT_BY_STYLE: Record<string, string> = {
  decimal: "decimal",
  lowerAlpha: "lowerLetter",
  upperAlpha: "upperLetter",
  lowerRoman: "lowerRoman",
  upperRoman: "upperRoman",
};

const STYLE_BY_NUM_FMT: Record<string, ListMarkerStyle> = {
  decimal: "decimal",
  lowerLetter: "lowerAlpha",
  upperLetter: "upperAlpha",
  lowerRoman: "lowerRoman",
  upperRoman: "upperRoman",
};

const BULLET_CHAR_BY_STYLE: Record<string, string> = {
  disc: "•",
  circle: "◦",
  square: "▪",
};

// --- reading ---------------------------------------------------------------

interface NumberingLevelDef {
  type: ListType;
  marker: ListMarkerStyle;
  startAt: number;
  /** w:ind/@w:left in twips; undefined when the level declares none. */
  leftTwips?: number;
}

interface NumInstance {
  abstractNumId: string;
  /** w:lvlOverride/w:startOverride values by ilvl. */
  startOverrides: Map<number, number>;
}

export interface NumberingCatalog {
  /** Resolves a paragraph's w:numPr into the model's list formatting, or
   * undefined when the numId is unknown (or "0", OOXML's "no numbering"). */
  listFormattingFor(numId: string, ilvl: number): ListFormatting | undefined;
}

export const EMPTY_NUMBERING_CATALOG: NumberingCatalog = {
  listFormattingFor: () => undefined,
};

function parseLevel(lvlEl: Element): NumberingLevelDef {
  const numFmtEl = firstChildElement(lvlEl, WORD_NAMESPACE, "numFmt");
  const numFmt = numFmtEl ? (wordAttr(numFmtEl, "val") ?? "decimal") : "decimal";
  const lvlTextEl = firstChildElement(lvlEl, WORD_NAMESPACE, "lvlText");
  const lvlText = lvlTextEl ? (wordAttr(lvlTextEl, "val") ?? "") : "";
  const startEl = firstChildElement(lvlEl, WORD_NAMESPACE, "start");
  const startVal = startEl ? Number(wordAttr(startEl, "val")) : NaN;
  const ilvl = Number(wordAttr(lvlEl, "ilvl") ?? "0");
  const level = clampListLevel(ilvl + 1);

  let type: ListType;
  let marker: ListMarkerStyle;
  if (numFmt === "bullet") {
    type = "bullet";
    marker = bulletStyleForChar(lvlText) ?? defaultMarkerStyle("bullet", level);
  } else {
    // Any numeric format outside the supported set (ordinal, cardinalText,
    // …) degrades to the closest supported marker rather than being
    // dropped (see intake ambiguity on unsupported markers).
    type = "number";
    marker = STYLE_BY_NUM_FMT[numFmt] ?? "decimal";
  }

  const def: NumberingLevelDef = {
    type,
    marker,
    startAt: Number.isFinite(startVal) && startVal >= 0 ? startVal : 1,
  };

  const pPr = firstChildElement(lvlEl, WORD_NAMESPACE, "pPr");
  const ind = pPr ? firstChildElement(pPr, WORD_NAMESPACE, "ind") : undefined;
  const left = ind ? (wordAttr(ind, "left") ?? wordAttr(ind, "start")) : null;
  if (left !== null && left !== "") {
    const twips = Number(left);
    if (Number.isFinite(twips)) {
      def.leftTwips = twips;
    }
  }
  return def;
}

/** Parses a numbering part into a catalog. `undefined` (no numbering part)
 * yields the empty catalog. */
export function parseNumberingCatalog(numberingDoc: Document | undefined): NumberingCatalog {
  if (!numberingDoc) {
    return EMPTY_NUMBERING_CATALOG;
  }
  const root = numberingDoc.documentElement;
  if (root.namespaceURI !== WORD_NAMESPACE || root.localName !== "numbering") {
    return EMPTY_NUMBERING_CATALOG;
  }

  const abstractNums = new Map<string, Map<number, NumberingLevelDef>>();
  const nums = new Map<string, NumInstance>();
  for (const child of Array.from(root.children)) {
    if (child.namespaceURI !== WORD_NAMESPACE) {
      continue;
    }
    if (child.localName === "abstractNum") {
      const id = wordAttr(child, "abstractNumId");
      if (id === null) {
        continue;
      }
      const levels = new Map<number, NumberingLevelDef>();
      for (const lvlEl of Array.from(child.children)) {
        if (lvlEl.namespaceURI === WORD_NAMESPACE && lvlEl.localName === "lvl") {
          const ilvl = Number(wordAttr(lvlEl, "ilvl") ?? "");
          if (Number.isInteger(ilvl) && ilvl >= 0) {
            levels.set(ilvl, parseLevel(lvlEl));
          }
        }
      }
      abstractNums.set(id, levels);
    } else if (child.localName === "num") {
      const numId = wordAttr(child, "numId");
      const abstractRef = firstChildElement(child, WORD_NAMESPACE, "abstractNumId");
      const abstractNumId = abstractRef ? wordAttr(abstractRef, "val") : null;
      if (numId === null || abstractNumId === null) {
        continue;
      }
      const startOverrides = new Map<number, number>();
      for (const overrideEl of Array.from(child.children)) {
        if (overrideEl.namespaceURI !== WORD_NAMESPACE || overrideEl.localName !== "lvlOverride") {
          continue;
        }
        const ilvl = Number(wordAttr(overrideEl, "ilvl") ?? "");
        const startEl = firstChildElement(overrideEl, WORD_NAMESPACE, "startOverride");
        const start = startEl ? Number(wordAttr(startEl, "val")) : NaN;
        if (Number.isInteger(ilvl) && Number.isFinite(start)) {
          startOverrides.set(ilvl, start);
        }
      }
      nums.set(numId, { abstractNumId, startOverrides });
    }
  }

  return {
    listFormattingFor(numId: string, ilvl: number): ListFormatting | undefined {
      if (numId === "0") {
        return undefined; // OOXML's explicit "no numbering".
      }
      const num = nums.get(numId);
      const levels = num ? abstractNums.get(num.abstractNumId) : undefined;
      if (!num || !levels) {
        return undefined;
      }
      // Levels beyond 9 clamp onto level 9 (see the intake ambiguity on
      // unsupported depth) and fall back to a bare level-9 definition when
      // the file defines none for the original ilvl.
      const def = levels.get(ilvl) ?? levels.get(MAX_LIST_LEVEL - 1);
      const level = clampListLevel(ilvl + 1);
      if (!def) {
        return { type: "number", level };
      }
      const list: ListFormatting = { type: def.type, level };
      if (def.marker !== defaultMarkerStyle(def.type, level)) {
        list.marker = def.marker;
      }
      const startAt = num.startOverrides.get(ilvl) ?? def.startAt;
      if (def.type === "number" && startAt !== 1) {
        list.startAt = startAt;
      }
      // Stored only when it differs from the editor default for the level;
      // compared in twips so float noise cannot leak into the model.
      if (def.leftTwips !== undefined && def.leftTwips !== level * 720) {
        list.indentCm = twipsToCm(def.leftTwips);
      }
      return list;
    },
  };
}

// --- writing ---------------------------------------------------------------

/** A planned reference from one list paragraph to its num instance. */
export interface NumberingRef {
  numId: number;
  ilvl: number;
}

/** The writer-side numbering plan: per-paragraph w:numPr values plus the
 * abstractNum/num definitions the numbering part needs. */
export interface NumberingPlan {
  /** Keyed "section:block". */
  refs: Map<string, NumberingRef>;
  abstractNumsXml: string;
  numsXml: string;
  isEmpty: boolean;
}

export function numberingRefFor(
  plan: NumberingPlan,
  section: number,
  block: number,
): NumberingRef | undefined {
  return plan.refs.get(`${section}:${block}`);
}

function escapeAttr(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function levelXml(ilvl: number, def: PlannedListLevel): string {
  const numFmt = def.type === "bullet" ? "bullet" : NUM_FMT_BY_STYLE[def.style];
  const lvlText = def.type === "bullet" ? BULLET_CHAR_BY_STYLE[def.style] : `%${ilvl + 1}.`;
  return (
    `<w:lvl w:ilvl="${ilvl}">` +
    `<w:start w:val="${def.startAt}"/>` +
    `<w:numFmt w:val="${numFmt}"/>` +
    `<w:lvlText w:val="${escapeAttr(lvlText)}"/>` +
    '<w:lvlJc w:val="left"/>' +
    `<w:pPr><w:ind w:left="${cmToTwips(def.indentCm)}" w:hanging="${HANGING_TWIPS}"/></w:pPr>` +
    "</w:lvl>"
  );
}

/** Plans the numbering for every list paragraph in the model. `firstIds`
 * lets a same-format save allocate ids above whatever the preserved
 * numbering part already uses. */
export function planNumbering(
  model: DocumentModel,
  firstIds: { abstractNum: number; num: number } = { abstractNum: 1, num: 1 },
): NumberingPlan {
  const refs = new Map<string, NumberingRef>();
  const abstractByKey = new Map<string, number>();
  const abstractXml: string[] = [];
  const numXml: string[] = [];
  let nextAbstract = firstIds.abstractNum;
  let nextNum = firstIds.num;

  model.sections.forEach((section, sectionIndex) => {
    let runStart = -1;
    const flushRun = (end: number) => {
      if (runStart === -1) {
        return;
      }
      const blocks = section.blocks.slice(runStart, end);
      const items = blocks.map((block) => block.list!);
      const levels = planListLevels(items);
      const key = JSON.stringify(levels);
      let abstractId = abstractByKey.get(key);
      if (abstractId === undefined) {
        abstractId = nextAbstract++;
        abstractByKey.set(key, abstractId);
        abstractXml.push(
          `<w:abstractNum w:abstractNumId="${abstractId}">` +
            '<w:multiLevelType w:val="hybridMultilevel"/>' +
            levels.map((def, ilvl) => levelXml(ilvl, def)).join("") +
            "</w:abstractNum>",
        );
      }
      const numId = nextNum++;
      numXml.push(`<w:num w:numId="${numId}"><w:abstractNumId w:val="${abstractId}"/></w:num>`);
      items.forEach((item, offset) => {
        refs.set(`${sectionIndex}:${runStart + offset}`, {
          numId,
          ilvl: clampListLevel(item.level) - 1,
        });
      });
      runStart = -1;
    };
    section.blocks.forEach((block, blockIndex) => {
      if (block.list) {
        if (runStart === -1) {
          runStart = blockIndex;
        }
      } else {
        flushRun(blockIndex);
      }
    });
    flushRun(section.blocks.length);
  });

  return {
    refs,
    abstractNumsXml: abstractXml.join(""),
    numsXml: numXml.join(""),
    isEmpty: refs.size === 0,
  };
}

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/** A fresh numbering part from a plan (AC-18). */
export function buildNumberingXml(plan: NumberingPlan): string {
  return (
    `${XML_DECLARATION}<w:numbering xmlns:w="${WORD_NAMESPACE}">` +
    plan.abstractNumsXml +
    plan.numsXml +
    "</w:numbering>"
  );
}

/** The first abstractNumId/numId above everything a preserved numbering
 * part already defines, so merged-in definitions can never collide. */
export function nextNumberingIds(numberingXml: string): { abstractNum: number; num: number } {
  let maxAbstract = 0;
  let maxNum = 0;
  try {
    const root = parseXml(numberingXml).documentElement;
    for (const child of Array.from(root.children)) {
      if (child.namespaceURI !== WORD_NAMESPACE) {
        continue;
      }
      if (child.localName === "abstractNum") {
        maxAbstract = Math.max(maxAbstract, Number(wordAttr(child, "abstractNumId")) || 0);
      } else if (child.localName === "num") {
        maxNum = Math.max(maxNum, Number(wordAttr(child, "numId")) || 0);
      }
    }
  } catch {
    // Unparseable preserved part: start clear of typical id ranges.
    return { abstractNum: 1000, num: 1000 };
  }
  return { abstractNum: maxAbstract + 1, num: maxNum + 1 };
}

/** Merges a plan's definitions into a preserved numbering part: the
 * original definitions stay verbatim (opaque tables and inlines may
 * reference them), our abstractNums are inserted before the first w:num
 * (the schema orders all abstractNum before all num) and our nums are
 * appended. DOM-based, so producer prefix choices don't matter. */
export function mergeNumberingXml(originalXml: string, plan: NumberingPlan): string {
  const doc = parseXml(originalXml);
  const root = doc.documentElement;
  const fragments = parseXml(
    `<w:numbering xmlns:w="${WORD_NAMESPACE}">${plan.abstractNumsXml}${plan.numsXml}</w:numbering>`,
  ).documentElement;
  const firstNum =
    Array.from(root.children).find(
      (child) => child.namespaceURI === WORD_NAMESPACE && child.localName === "num",
    ) ?? null;
  for (const child of Array.from(fragments.children)) {
    const imported = doc.importNode(child, true);
    if (child.localName === "abstractNum") {
      root.insertBefore(imported, firstNum);
    } else {
      root.appendChild(imported);
    }
  }
  return `${XML_DECLARATION}${serializeElement(root)}`;
}
