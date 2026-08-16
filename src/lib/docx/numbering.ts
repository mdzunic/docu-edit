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
import type { Block } from "../documentModel";
import type {
  ListDefinition,
  ListFormatting,
  ListLevelDefinition,
  ListMarkerStyle,
  ListType,
} from "../core/lists";
import {
  bulletStyleForChar,
  clampListLevel,
  defaultMarkerStyle,
  isBulletStyle,
  LIST_INDENT_STEP_CM,
  LIST_MARKER_WIDTH_CM,
  MAX_LIST_LEVEL,
  planListLevels,
  type PlannedListLevel,
} from "../core/lists";
import type { DocumentModel } from "../documentModel";
import { wordAttr, WORD_NAMESPACE } from "./xml";
import { escapeXmlAttribute, firstChildElement, parseXml, serializeElement } from "../core/xml";
import { cmToTwips, twipsToCm } from "../core/units";

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
  /** The raw w:lvlText value — what composite (SPEC-032) reconstruction
   * parses prefixes/suffixes/components back out of. */
  lvlText: string;
  /** w:lvlJc, lowercased; undefined when absent. */
  jc?: string;
  /** w:ind/@w:hanging (positive) or -@w:firstLine, in twips. */
  hangingTwips?: number;
  /** w:pStyle inside the level — a paragraph-style link (SPEC-032 AC-13). */
  pStyle?: string;
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
  /** The multilevel definitions (SPEC-032) the resolved references actually
   * used, in first-use order — what the reader stores as the model's
   * listDefinitions. */
  definitionsUsed(): ListDefinition[];
}

export const EMPTY_NUMBERING_CATALOG: NumberingCatalog = {
  listFormattingFor: () => undefined,
  definitionsUsed: () => [],
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
    lvlText,
  };

  const jcEl = firstChildElement(lvlEl, WORD_NAMESPACE, "lvlJc");
  const jc = jcEl ? wordAttr(jcEl, "val") : null;
  if (jc) {
    def.jc = jc.toLowerCase();
  }
  const pStyleEl = firstChildElement(lvlEl, WORD_NAMESPACE, "pStyle");
  const pStyle = pStyleEl ? wordAttr(pStyleEl, "val") : null;
  if (pStyle) {
    def.pStyle = pStyle;
  }

  const pPr = firstChildElement(lvlEl, WORD_NAMESPACE, "pPr");
  const ind = pPr ? firstChildElement(pPr, WORD_NAMESPACE, "ind") : undefined;
  const left = ind ? (wordAttr(ind, "left") ?? wordAttr(ind, "start")) : null;
  if (left !== null && left !== "") {
    const twips = Number(left);
    if (Number.isFinite(twips)) {
      def.leftTwips = twips;
    }
  }
  const hanging = ind ? wordAttr(ind, "hanging") : null;
  const firstLine = ind ? wordAttr(ind, "firstLine") : null;
  if (hanging !== null && hanging !== "" && Number.isFinite(Number(hanging))) {
    def.hangingTwips = Number(hanging);
  } else if (firstLine !== null && firstLine !== "" && Number.isFinite(Number(firstLine))) {
    def.hangingTwips = -Number(firstLine);
  }
  return def;
}

// --- composite (SPEC-032) definitions --------------------------------------

/** Splits a numeric lvlText template into prefix / component count / suffix
 * — the inverse of `levelTemplate` below. "%1.%2.%3" has three components;
 * the text before the first `%N` token is the prefix, after the last the
 * suffix (separators between tokens are normalised to "."). */
export function parseLevelTemplate(lvlText: string): {
  prefix: string;
  suffix: string;
  includeLevels: number;
} {
  const tokens = [...lvlText.matchAll(/%\d/g)];
  if (tokens.length === 0) {
    return { prefix: "", suffix: lvlText, includeLevels: 1 };
  }
  const first = tokens[0];
  const last = tokens[tokens.length - 1];
  return {
    prefix: lvlText.slice(0, first.index),
    suffix: lvlText.slice((last.index ?? 0) + last[0].length),
    includeLevels: tokens.length,
  };
}

/** Whether an abstractNum needs the full SPEC-032 definition treatment.
 * SPEC-018's own writer marks its simple per-run definitions
 * `hybridMultilevel` and only ever emits `%N.` numeric templates, plain
 * bullet glyphs, left justification and no style links — anything else
 * (notably our own `multilevel`-marked definitions, and Word's real
 * multilevel presets) reconstructs as a ListDefinition so instances and
 * restarts survive re-opening. */
function levelsAreComposite(
  multiLevelType: string | undefined,
  levels: Map<number, NumberingLevelDef>,
): boolean {
  if (multiLevelType === "multilevel") {
    return true;
  }
  for (const [ilvl, def] of levels) {
    if (def.pStyle !== undefined) {
      return true;
    }
    if (def.jc !== undefined && def.jc !== "left") {
      return true;
    }
    if (def.type === "bullet") {
      if (bulletStyleForChar(def.lvlText) === undefined) {
        return true;
      }
    } else if (def.lvlText !== `%${ilvl + 1}.`) {
      return true;
    }
  }
  return false;
}

/** Reconstructs a ListDefinition from a composite abstractNum's levels —
 * the inverse of `definitionAbstractXml`, with the same default elision as
 * commands.normalizeDefinitionLevels so write -> read is a fixed point. */
export function definitionFromLevels(
  defId: string,
  levels: Map<number, NumberingLevelDef>,
): ListDefinition {
  const out: ListLevelDefinition[] = [];
  let styleLinks: Record<string, number> | undefined;
  for (let ilvl = 0; ilvl < MAX_LIST_LEVEL; ilvl++) {
    const parsed = levels.get(ilvl);
    const level = ilvl + 1;
    if (!parsed) {
      out.push({ style: "decimal" });
      continue;
    }
    const levelDef: ListLevelDefinition = { style: parsed.marker };
    let prefix = "";
    let suffix = "";
    if (parsed.type === "bullet") {
      // The glyph may sit between literal prefix/suffix text; split around
      // the first recognised bullet character.
      const index = [...parsed.lvlText].findIndex((char) => bulletStyleForChar(char) !== undefined);
      if (index >= 0) {
        prefix = parsed.lvlText.slice(0, index);
        suffix = parsed.lvlText.slice(index + 1);
      } else {
        prefix = parsed.lvlText;
      }
    } else {
      const template = parseLevelTemplate(parsed.lvlText);
      prefix = template.prefix;
      suffix = template.suffix;
      if (template.includeLevels > 1) {
        levelDef.includeLevels = Math.min(template.includeLevels, level);
      }
    }
    if (prefix !== "") {
      levelDef.prefix = prefix;
    }
    if (suffix !== "") {
      levelDef.suffix = suffix;
    }
    if (parsed.startAt !== 1) {
      levelDef.startAt = parsed.startAt;
    }
    if (parsed.jc === "center" || parsed.jc === "right") {
      levelDef.alignment = parsed.jc;
    }
    if (parsed.leftTwips !== undefined && parsed.leftTwips !== level * 720) {
      levelDef.textIndentCm = twipsToCm(parsed.leftTwips);
    }
    const textTwips = parsed.leftTwips ?? level * 720;
    if (parsed.hangingTwips !== undefined) {
      const numberTwips = textTwips - parsed.hangingTwips;
      if (numberTwips !== Math.max(0, textTwips - 360)) {
        levelDef.numberIndentCm = twipsToCm(Math.max(0, numberTwips));
      }
    }
    if (parsed.pStyle !== undefined) {
      styleLinks = styleLinks ?? {};
      styleLinks[parsed.pStyle] = level;
    }
    out.push(levelDef);
  }
  const def: ListDefinition = { defId, levels: out };
  if (styleLinks) {
    def.styleLinks = styleLinks;
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
  const multiLevelTypes = new Map<string, string>();
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
      const typeEl = firstChildElement(child, WORD_NAMESPACE, "multiLevelType");
      const typeVal = typeEl ? wordAttr(typeEl, "val") : null;
      if (typeVal) {
        multiLevelTypes.set(id, typeVal);
      }
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

  // Composite (SPEC-032) definitions, reconstructed lazily per abstractNum
  // and recorded in first-use order; each num instance becomes a distinct
  // (defId, listId) pair, which is exactly what preserves independent lists
  // and restarts across a round trip (AC-26, AC-28). A startOverride is
  // attached to the FIRST paragraph resolved for its (numId, ilvl) — the
  // reader resolves paragraphs in document order, so that is the item the
  // restart begins at.
  const compositeDefs = new Map<string, ListDefinition>();
  const usedDefs: ListDefinition[] = [];
  const appliedOverrides = new Set<string>();

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
      const level = clampListLevel(ilvl + 1);
      if (levelsAreComposite(multiLevelTypes.get(num.abstractNumId), levels)) {
        let def = compositeDefs.get(num.abstractNumId);
        if (!def) {
          def = definitionFromLevels(`a${num.abstractNumId}`, levels);
          compositeDefs.set(num.abstractNumId, def);
        }
        if (!usedDefs.includes(def)) {
          usedDefs.push(def);
        }
        const list: ListFormatting = {
          type: isBulletStyle(def.levels[level - 1]?.style ?? "decimal") ? "bullet" : "number",
          level,
          defId: def.defId,
          listId: `n${numId}`,
        };
        const override = num.startOverrides.get(ilvl);
        const overrideKey = `${numId}:${ilvl}`;
        if (override !== undefined && !appliedOverrides.has(overrideKey)) {
          appliedOverrides.add(overrideKey);
          list.startAt = override;
        }
        return list;
      }
      // Levels beyond 9 clamp onto level 9 (see the intake ambiguity on
      // unsupported depth) and fall back to a bare level-9 definition when
      // the file defines none for the original ilvl.
      const def = levels.get(ilvl) ?? levels.get(MAX_LIST_LEVEL - 1);
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
    definitionsUsed(): ListDefinition[] {
      return usedDefs;
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

/** The plan entry for a paragraph inside a table cell (SPEC-022), keyed by
 * the full cell path. */
export function cellNumberingRefFor(
  plan: NumberingPlan,
  section: number,
  block: number,
  row: number,
  col: number,
  cellBlock: number,
): NumberingRef | undefined {
  return plan.refs.get(`${section}:${block}:${row}:${col}:${cellBlock}`);
}

function levelXml(ilvl: number, def: PlannedListLevel): string {
  const numFmt = def.type === "bullet" ? "bullet" : NUM_FMT_BY_STYLE[def.style];
  const lvlText = def.type === "bullet" ? BULLET_CHAR_BY_STYLE[def.style] : `%${ilvl + 1}.`;
  return (
    `<w:lvl w:ilvl="${ilvl}">` +
    `<w:start w:val="${def.startAt}"/>` +
    `<w:numFmt w:val="${numFmt}"/>` +
    `<w:lvlText w:val="${escapeXmlAttribute(lvlText)}"/>` +
    '<w:lvlJc w:val="left"/>' +
    `<w:pPr><w:ind w:left="${cmToTwips(def.indentCm)}" w:hanging="${HANGING_TWIPS}"/></w:pPr>` +
    "</w:lvl>"
  );
}

/** A defined level's w:lvlText (SPEC-032 AC-26): the literal prefix, the
 * shown level components as `%N` tokens joined with ".", and the literal
 * suffix — "(", includeLevels 3 at level 3, ")" composes to "(%1.%2.%3)".
 * Bullet levels compose prefix + glyph + suffix instead. */
export function levelTemplate(def: ListDefinition, level: number): string {
  const levelDef = def.levels[level - 1] ?? { style: "decimal" };
  const prefix = levelDef.prefix ?? "";
  const suffix = levelDef.suffix ?? "";
  if (isBulletStyle(levelDef.style)) {
    return `${prefix}${BULLET_CHAR_BY_STYLE[levelDef.style]}${suffix}`;
  }
  const include = Math.min(Math.max(Math.round(levelDef.includeLevels ?? 1), 1), level);
  const tokens: string[] = [];
  for (let component = level - include + 1; component <= level; component++) {
    tokens.push(`%${component}`);
  }
  return `${prefix}${tokens.join(".")}${suffix}`;
}

/** A ListDefinition's per-level effective indents in twips, defaults
 * resolved (text at level × 720, marker 360 before it). */
function definedIndentTwips(def: ListDefinition, level: number): { left: number; hanging: number } {
  const levelDef = def.levels[level - 1];
  const left =
    levelDef?.textIndentCm !== undefined
      ? cmToTwips(levelDef.textIndentCm)
      : cmToTwips(level * LIST_INDENT_STEP_CM);
  const number =
    levelDef?.numberIndentCm !== undefined
      ? cmToTwips(levelDef.numberIndentCm)
      : Math.max(0, left - cmToTwips(LIST_MARKER_WIDTH_CM));
  return { left, hanging: left - number };
}

/** One SPEC-032 definition as a distinct w:abstractNum (AC-26), marked
 * `multilevel` — the discriminator the reader's composite check keys on, so
 * a re-opened definition never degrades into a SPEC-018 adjacency list. */
export function definitionAbstractXml(def: ListDefinition, abstractId: number): string {
  const linkByLevel = new Map<number, string>();
  for (const [styleId, level] of Object.entries(def.styleLinks ?? {})) {
    if (!linkByLevel.has(level)) {
      linkByLevel.set(level, styleId);
    }
  }
  const levels: string[] = [];
  for (let level = 1; level <= MAX_LIST_LEVEL; level++) {
    const levelDef = def.levels[level - 1] ?? { style: "decimal" };
    const numFmt = isBulletStyle(levelDef.style) ? "bullet" : NUM_FMT_BY_STYLE[levelDef.style];
    const link = linkByLevel.get(level);
    const { left, hanging } = definedIndentTwips(def, level);
    const ind =
      hanging >= 0
        ? `<w:ind w:left="${left}" w:hanging="${hanging}"/>`
        : `<w:ind w:left="${left}" w:firstLine="${-hanging}"/>`;
    levels.push(
      `<w:lvl w:ilvl="${level - 1}">` +
        `<w:start w:val="${levelDef.startAt ?? 1}"/>` +
        `<w:numFmt w:val="${numFmt}"/>` +
        (link !== undefined ? `<w:pStyle w:val="${escapeXmlAttribute(link)}"/>` : "") +
        `<w:lvlText w:val="${escapeXmlAttribute(levelTemplate(def, level))}"/>` +
        `<w:lvlJc w:val="${levelDef.alignment ?? "left"}"/>` +
        `<w:pPr>${ind}</w:pPr>` +
        "</w:lvl>",
    );
  }
  return (
    `<w:abstractNum w:abstractNumId="${abstractId}">` +
    '<w:multiLevelType w:val="multilevel"/>' +
    levels.join("") +
    "</w:abstractNum>"
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

  // Defined lists (SPEC-032): one abstractNum per referenced definition,
  // one num per (context, defId, listId) instance — a fresh num is exactly
  // how OOXML restarts a sequence, so restart/continue round-trip for free
  // (AC-26, AC-28). Instances are collected across the whole walk (a
  // top-level instance may span sections); the context key keeps a table
  // cell's occurrences apart from the body's, mirroring the model's
  // per-cell numbering contexts.
  const definedAbstractIds = new Map<string, number>();
  interface DefinedInstance {
    def: ListDefinition;
    refKeys: { key: string; ilvl: number }[];
    startOverrides: Map<number, number>;
  }
  const definedInstances = new Map<string, DefinedInstance>();

  const recordDefined = (list: ListFormatting, contextKey: string, refKey: string) => {
    const def = model.listDefinitions?.find((entry) => entry.defId === list.defId);
    if (!def) {
      return false;
    }
    const instanceKey = `${contextKey} ${list.defId} ${list.listId ?? ""}`;
    let instance = definedInstances.get(instanceKey);
    if (!instance) {
      instance = { def, refKeys: [], startOverrides: new Map() };
      definedInstances.set(instanceKey, instance);
    }
    const ilvl = clampListLevel(list.level) - 1;
    instance.refKeys.push({ key: refKey, ilvl });
    if (list.startAt !== undefined && !instance.startOverrides.has(ilvl)) {
      instance.startOverrides.set(ilvl, list.startAt);
    }
    return true;
  };

  // One pass over a block list: every contiguous run of (SPEC-018) list
  // paragraphs becomes one num referencing a (deduplicated) abstract
  // definition; a defined (SPEC-032) paragraph is recorded for the instance
  // planning above and ends any ad-hoc run. Used for each section's
  // top-level blocks and — with its own key space — for every table cell's
  // block list (SPEC-022: a cell is its own list context, so its runs never
  // join a run outside the table).
  const planBlockList = (
    blocks: readonly Block[],
    keyFor: (index: number) => string,
    contextKey: string,
  ) => {
    let runStart = -1;
    const flushRun = (end: number) => {
      if (runStart === -1) {
        return;
      }
      const items: ListFormatting[] = [];
      for (let index = runStart; index < end; index++) {
        const block = blocks[index];
        if (block.type === "paragraph" && block.list) {
          items.push(block.list);
        }
      }
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
        refs.set(keyFor(runStart + offset), {
          numId,
          ilvl: clampListLevel(item.level) - 1,
        });
      });
      runStart = -1;
    };
    blocks.forEach((block, blockIndex) => {
      const list = block.type === "paragraph" ? block.list : undefined;
      if (list?.defId !== undefined && recordDefined(list, contextKey, keyFor(blockIndex))) {
        flushRun(blockIndex);
        return;
      }
      if (block.type === "paragraph" && block.list) {
        if (runStart === -1) {
          runStart = blockIndex;
        }
      } else {
        flushRun(blockIndex);
      }
    });
    flushRun(blocks.length);
  };

  model.sections.forEach((section, sectionIndex) => {
    planBlockList(section.blocks, (index) => `${sectionIndex}:${index}`, "");
    section.blocks.forEach((block, blockIndex) => {
      if (block.type !== "table") {
        return;
      }
      block.rows.forEach((row, rowIndex) => {
        row.cells.forEach((cell, colIndex) => {
          planBlockList(
            cell.blocks,
            (index) => `${sectionIndex}:${blockIndex}:${rowIndex}:${colIndex}:${index}`,
            `${sectionIndex}:${blockIndex}:${rowIndex}:${colIndex}`,
          );
        });
      });
    });
  });

  // Materialise the defined plan: abstracts in first-use order, then one
  // num per instance carrying its startOverrides.
  for (const instance of definedInstances.values()) {
    const defId = instance.def.defId;
    let abstractId = definedAbstractIds.get(defId);
    if (abstractId === undefined) {
      abstractId = nextAbstract++;
      definedAbstractIds.set(defId, abstractId);
      abstractXml.push(definitionAbstractXml(instance.def, abstractId));
    }
    const numId = nextNum++;
    const overrides = [...instance.startOverrides.entries()]
      .sort((a, b) => a[0] - b[0])
      .map(
        ([ilvl, start]) =>
          `<w:lvlOverride w:ilvl="${ilvl}"><w:startOverride w:val="${start}"/></w:lvlOverride>`,
      )
      .join("");
    numXml.push(
      `<w:num w:numId="${numId}"><w:abstractNumId w:val="${abstractId}"/>${overrides}</w:num>`,
    );
    for (const ref of instance.refKeys) {
      refs.set(ref.key, { numId, ilvl: ref.ilvl });
    }
  }

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
