// ODS content validations (SPEC-130 "ODS wire mapping") — THIS FILE is the
// normative source for how content.xml's `table:content-validations` and
// the cells' `table:content-validation-name` read into Sheet.validations,
// and how the model's rules write back.
//
// Reading (readOdsContentValidations — the reader after its scan;
// readTableValidations — SPEC-130's reader upgrade over one table text):
// - The definitions are the `table:content-validation` children of the
//   envelope prefix's `table:content-validations` (a child of
//   `office:spreadsheet`), their names matched under every prefix the root
//   binds to the table, text and office namespaces. Attributes:
//   `@table:name` → `odsName`; `@table:allow-empty-cell` (true) →
//   `allowBlank`; `@table:display-list` (unsorted): `none` ↔ `showDropdown:
//   false`, `unsorted` ↔ true, `sort-ascending` ↔ true + `odsListOrder`;
//   `@table:base-cell-address` → the cell the condition's relative operands
//   are written at; `@table:condition`:
//   | model | `table:condition` (`of:` or LibreOffice's older `oooc:`) |
//   | `any` | absent |
//   | `whole`/`decimal`/`date`/`time` | `cell-content-is-whole-number()` (`-decimal-number()`, `-date()`, `-time()`), alone or `and` one of: |
//   |   + `between`/`notBetween` | `cell-content-is-between(a,b)`, `cell-content-is-not-between(a,b)` |
//   |   + comparison | `cell-content()<op>a` (`=` `!=` `<` `>` `<=` `>=`; `<>` read as `!=`) |
//   | `textLength` | `cell-content-text-length()<op>a`, `cell-content-text-length-is-between(a,b)`, `…-is-not-between(a,b)` |
//   | `list` literal | `cell-content-is-in-list("a";"b";"c")` (a number item read as its text) |
//   | `list` other | `cell-content-is-in-list(<formula>)` |
//   | `custom` | `is-true-formula(<formula>)` |
//   The two operands of a between form are split at the top-level `,`
//   (ODF's grammar, and what LibreOffice writes and reads) or `;`. Operands
//   are OpenFormula (fromOdsFormula, the condition's own dialect prefix),
//   rebased from the base cell onto each sheet's origin (its first range's
//   top-left) with shiftFormula; a literal list has no reference to rebase.
//   `table:help-message @table:title @table:display (false)` ↔ `input`;
//   `table:error-message @table:title @table:display (false)
//   @table:message-type (stop)` ↔ `error`; each message's `text:p`
//   paragraphs (`text:s`, `text:tab`, `text:line-break` expanded) joined
//   with "\n". A `table:error-macro` (and its `office:event-listeners`) is
//   kept in `sourceXml` only — never run — and the rule's `error.show` is
//   false, so it enforces nothing. A condition this model cannot read (an
//   unknown form or dialect, an operand the dialect cannot parse) reads as
//   `any` with the element kept in `sourceXml`: re-emitted while unchanged.
// - `sourceXml` is the element verbatim. A definition no cell references is
//   an orphan: `SpreadsheetModel.odsOrphanValidations`, verbatim with the
//   whitespace before it.
// - Ranges come from the cells: each `table:table-cell` /
//   `table:covered-table-cell` naming a validation contributes its column
//   run, each row its row run (`number-columns-repeated`,
//   `number-rows-repeated`) — accumulated as row-run × column-run
//   rectangles (ValidationRegions), nothing ever materialised, so a rule
//   over LibreOffice's trailing 1,048,553-row run is one rectangle; a run
//   covering the same columns as one ending on the row just above extends
//   it. Clamped to the grid. A definition referenced from two tables is one
//   DataValidation per sheet, with the same `odsName` and `sourceXml`.
//   Rules are in the order of their definitions; a rule's ranges in
//   document order, but the one whose top-left is the base cell first (the
//   origin the writer wrote it at).
//
// Writing (OdsValidationWriter, one per save — ods/writer.ts):
// - Every rule gets the name its cells are written with (nameOf; the table
//   serialiser writes `table:content-validation-name` from it, the first
//   covering rule winning — ods/tableXml.ts). A table emitted from its
//   retained text (a clean sheet) still names the ORIGINAL definitions, so
//   its rules are claimed first, each under its `odsName`: from its
//   `sourceXml` while it has one, else serialised from the model (a sheet
//   rename or delete remaps a rule's formulas without dirtying its sheet).
//   A name two sheets claim is written once when the one definition reads
//   back as both rules (a rule shared by two clean sheets); a clean sheet
//   whose rules cannot all keep their names (a duplicate's rules have none,
//   or a shared definition has diverged) is rebuilt (rebuilds()). Then the
//   rebuilt and new sheets: an unchanged rule (its `sourceXml` still held)
//   keeps its `odsName` and `sourceXml` unless that name already holds
//   another definition; every other rule is serialised under a fresh
//   `de-val<n>` colliding with no name in use (a definition's, a rule's, an
//   orphan's, or one a kept table names).
// - The block (blockXml, splicePrefix): the source's start tag and closing text kept;
//   the definitions in the source's order (with the whitespace before each),
//   orphans re-emitted verbatim in place, new ones appended; retained text
//   through the writer's sheet-name pass, generated text not. A new block
//   goes after `table:calculation-settings`, before `table:label-ranges`
//   and the first table; an empty one is omitted.
// - A rule serialised: `table:name`, `table:condition` per the table above
//   (operands through toOdsFormula, relative to the rule's origin; between
//   operands joined with `,`; a literal list as quoted items joined with
//   `;` — an ODS item holding a comma became two items when read),
//   `table:allow-empty-cell` (whenever there is a condition, or false),
//   `table:display-list` (always for a list, else when not unsorted),
//   `table:base-cell-address` (the origin, sheet-qualified as LibreOffice
//   writes it: `Form.B2`),
//   then `table:help-message` and `table:error-message` (each written when
//   shown or holding a title, text or a style other than stop), one `text:p`
//   per line.
//
// Layering rule (boundary.spec.ts): framework-free, imports only
// spreadsheet siblings and src/lib/core.
import { OFFICE_NAMESPACE, TABLE_NAMESPACE, TEXT_NAMESPACE } from "../../core/odfXml";
import { scanXml, XmlScanError } from "../../core/xmlScanner";
import { sanitizeXmlText } from "../../core/xmlText";
import type { CellRef, GridRange } from "../address";
import { fromOdsFormula, OOOC_NAMESPACE, OPENFORMULA_NAMESPACE, toOdsFormula } from "../formula/dialect";
import { FormulaSyntaxError } from "../formula/parse";
import { shiftFormula } from "../formula/references";
import { MAX_COLS, MAX_ROWS, type Sheet, type SpreadsheetModel } from "../model";
import {
  formatListLiteral,
  isRangeOperator,
  mintValidationId,
  parseListLiteral,
  sameValidationSettings,
  sheetValidations,
  validationOrigin,
  type DataValidation,
  type ValidationErrorStyle,
  type ValidationOperator,
  type ValidationSettings,
  type ValidationType,
} from "../validation";
import { decodeXmlText, escapeAttributeValue } from "../xmlSplice";
import { formatOdsRangeAddress, parseOdsRangeAddress } from "./rangeAddress";
import { rootNamespaces } from "./sheetNamePass";

// --- the wire's vocabularies --------------------------------------------------------------------

const TYPE_FUNCTIONS: Readonly<Record<string, ValidationType>> = {
  "cell-content-is-whole-number": "whole",
  "cell-content-is-decimal-number": "decimal",
  "cell-content-is-date": "date",
  "cell-content-is-time": "time",
};

const TYPE_FUNCTION_WRITE: Readonly<Partial<Record<ValidationType, string>>> = {
  whole: "cell-content-is-whole-number",
  decimal: "cell-content-is-decimal-number",
  date: "cell-content-is-date",
  time: "cell-content-is-time",
};

const COMPARISONS: Readonly<Record<string, ValidationOperator>> = {
  "<=": "lessThanOrEqual",
  ">=": "greaterThanOrEqual",
  "!=": "notEqual",
  "<>": "notEqual",
  "<": "lessThan",
  ">": "greaterThan",
  "=": "equal",
};

const COMPARISON_WRITE: Readonly<Record<Exclude<ValidationOperator, "between" | "notBetween">, string>> = {
  equal: "=",
  notEqual: "!=",
  greaterThan: ">",
  lessThan: "<",
  greaterThanOrEqual: ">=",
  lessThanOrEqual: "<=",
};

/** The dialect prefixes a condition may carry when the root binds none. */
const DIALECTS: ReadonlyMap<string, string> = new Map([
  ["of", OPENFORMULA_NAMESPACE],
  ["oooc", OOOC_NAMESPACE],
]);

const COMPARISON = /^(<=|>=|!=|<>|<|>|=)([\s\S]*)$/;

const ERROR_STYLES: ReadonlySet<string> = new Set(["stop", "warning", "information"]);

// --- fragments ----------------------------------------------------------------------------------

/** The qualified-name normaliser for fragments of a part whose root declares
 * `namespaces`: a prefix bound to the table, text or office namespace reads
 * as `table:`, `text:`, `office:` (an unbound one as written). */
function qualifier(namespaces: ReadonlyMap<string, string>): (name: string) => string {
  const canonical = new Map<string, string>();
  for (const [prefix, uri] of namespaces) {
    if (uri === TABLE_NAMESPACE) {
      canonical.set(prefix, "table");
    } else if (uri === TEXT_NAMESPACE) {
      canonical.set(prefix, "text");
    } else if (uri === OFFICE_NAMESPACE) {
      canonical.set(prefix, "office");
    }
  }
  return (name) => {
    const colon = name.indexOf(":");
    if (colon <= 0) {
      return name;
    }
    const prefix = canonical.get(name.slice(0, colon));
    return prefix === undefined ? name : `${prefix}${name.slice(colon)}`;
  };
}

/** The `table:content-validations` element of a content prefix: its offsets,
 * start tag, the closing text after its last child and each definition (its
 * name, verbatim text and the whitespace before it). */
interface SourceBlock {
  start: number;
  end: number;
  open: string;
  /** The text between the last definition and the close tag, plus it. */
  close: string;
  definitions: { name: string; xml: string; leading: string }[];
}

/** The prefix's `table:content-validations` (see SourceBlock), or null. */
function locateBlock(prefix: string, namespaces: ReadonlyMap<string, string>): SourceBlock | null {
  if (!/content-validations[\s/>]/.test(prefix)) {
    return null;
  }
  const qualify = qualifier(namespaces);
  let found: SourceBlock | null = null;
  let depth = 0;
  let blockDepth = -1;
  let lastEnd = -1;
  let childStart = -1;
  let childName = "";
  try {
    scanXml(prefix, (event) => {
      if (event.type === "text") {
        return undefined;
      }
      if (event.type === "open") {
        const name = qualify(event.name);
        depth += 1;
        if (blockDepth < 0 && name === "table:content-validations") {
          found = { start: event.start, end: event.end, open: prefix.slice(event.start, event.end), close: "", definitions: [] };
          lastEnd = event.end;
          if (event.selfClosing) {
            found.open = found.open.replace(/\s*\/>$/, ">");
            found.close = `</${event.name}>`;
            return "stop";
          }
          blockDepth = depth;
        } else if (blockDepth > 0 && depth === blockDepth + 1) {
          childStart = event.start;
          childName = name === "table:content-validation" ? attributeOf(event.attrs, "table:name", qualify) : "";
          if (event.selfClosing) {
            finishChild(event.end);
          }
        } else if (blockDepth < 0 && /:(?:table|label-ranges)$/.test(name) && name.startsWith("table:")) {
          return "stop"; // past where the block may be
        }
        if (event.selfClosing) {
          depth -= 1;
        }
        return undefined;
      }
      if (blockDepth > 0 && depth === blockDepth + 1 && childStart >= 0) {
        finishChild(event.end);
      } else if (blockDepth > 0 && depth === blockDepth && found) {
        found.end = event.end;
        found.close = prefix.slice(lastEnd, event.end);
        depth -= 1;
        return "stop";
      }
      depth -= 1;
      return undefined;
    });
  } catch (error) {
    if (!(error instanceof XmlScanError)) {
      throw error;
    }
  }
  const block = found as SourceBlock | null;
  return block && block.close !== "" ? block : null;

  function finishChild(end: number): void {
    const into = found!;
    if (childName !== "") {
      into.definitions.push({ name: childName, xml: prefix.slice(childStart, end), leading: prefix.slice(lastEnd, childStart) });
    } else {
      // Anything else inside the block is carried as text before the next definition.
      into.definitions.push({ name: "", xml: prefix.slice(childStart, end), leading: prefix.slice(lastEnd, childStart) });
    }
    lastEnd = end;
    childStart = -1;
  }
}

function attributeOf(attrs: readonly { name: string; value: string }[], qname: string, qualify: (name: string) => string): string {
  return attrs.find((attr) => qualify(attr.name) === qname)?.value ?? "";
}

// --- reading a definition ------------------------------------------------------------------------

/** A message element as read: its attributes and its paragraphs' text. */
interface WireMessage {
  attrs: ReadonlyMap<string, string>;
  text: string;
}

/** One `table:content-validation` element as read. */
interface WireDefinition {
  attrs: ReadonlyMap<string, string>;
  help?: WireMessage;
  error?: WireMessage;
  macro: boolean;
}

function wireDefinition(xml: string, qualify: (name: string) => string): WireDefinition | null {
  let definition: WireDefinition | null = null;
  let depth = 0;
  let message: { kind: "help" | "error"; attrs: Map<string, string>; paragraphs: string[] } | null = null;
  let paragraph: string | null = null;
  let paragraphDepth = -1;
  const attrsOf = (attrs: readonly { name: string; value: string }[]): Map<string, string> => new Map(attrs.map((attr) => [qualify(attr.name), attr.value]));
  scanXml(xml, (event) => {
    if (event.type === "text") {
      if (paragraph !== null) {
        paragraph += event.text;
      }
      return undefined;
    }
    if (event.type === "close") {
      if (paragraph !== null && depth === paragraphDepth) {
        message?.paragraphs.push(paragraph);
        paragraph = null;
      } else if (message && depth === 2) {
        const read: WireMessage = { attrs: message.attrs, text: message.paragraphs.join("\n") };
        if (message.kind === "help") {
          definition!.help = read;
        } else {
          definition!.error = read;
        }
        message = null;
      }
      depth -= 1;
      return undefined;
    }
    depth += 1;
    const name = qualify(event.name);
    if (depth === 1) {
      definition = { attrs: attrsOf(event.attrs), macro: false };
    } else if (depth === 2 && definition) {
      if (name === "table:help-message" || name === "table:error-message") {
        const kind = name === "table:help-message" ? "help" : "error";
        if (event.selfClosing) {
          definition[kind] = { attrs: attrsOf(event.attrs), text: "" };
        } else {
          message = { kind, attrs: attrsOf(event.attrs), paragraphs: [] };
        }
      } else if (name === "table:error-macro") {
        definition.macro = true;
      }
    } else if (message) {
      if (paragraph === null && (name === "text:p" || name === "text:h")) {
        if (event.selfClosing) {
          message.paragraphs.push("");
        } else {
          paragraph = "";
          paragraphDepth = depth;
        }
      } else if (paragraph !== null) {
        if (name === "text:s") {
          const count = Number(event.attrs.find((attr) => qualify(attr.name) === "text:c")?.value ?? 1);
          paragraph += " ".repeat(Number.isInteger(count) && count > 0 ? Math.min(count, 10_000) : 1);
        } else if (name === "text:tab") {
          paragraph += "\t";
        } else if (name === "text:line-break") {
          paragraph += "\n";
        }
      }
    }
    if (event.selfClosing) {
      depth -= 1;
    }
    return undefined;
  });
  return definition;
}

/** A condition as read: the type and operator, and the operands
 * (canonical, relative to the base cell). */
interface ReadCondition {
  type: ValidationType;
  operator: ValidationOperator;
  formula1?: string;
  formula2?: string;
  /** A literal list's formula1 holds no reference to rebase. */
  literal?: true;
}

/** `inner`'s arguments split at the top-level separators (outside quotes,
 * brackets, parentheses and braces). */
function splitTopLevel(inner: string, separators: string): string[] {
  const out: string[] = [];
  let depth = 0;
  let quote = "";
  let from = 0;
  for (let i = 0; i < inner.length; i += 1) {
    const ch = inner[i]!;
    if (quote !== "") {
      if (ch === quote) {
        quote = "";
      }
    } else if (ch === '"' || ch === "'") {
      quote = ch;
    } else if (ch === "(" || ch === "[" || ch === "{") {
      depth += 1;
    } else if (ch === ")" || ch === "]" || ch === "}") {
      depth -= 1;
    } else if (depth === 0 && separators.includes(ch)) {
      out.push(inner.slice(from, i));
      from = i + 1;
    }
  }
  out.push(inner.slice(from));
  return out;
}

/** `name( … )` spanning the whole of `text` → the text between the
 * parentheses; null when `text` is not that one call. */
function callInner(text: string, name: string): string | null {
  if (!text.startsWith(`${name}(`) || !text.endsWith(")")) {
    return null;
  }
  const inner = text.slice(name.length + 1, -1);
  // The parentheses must match each other: the call ends at the end.
  let depth = 0;
  let quote = "";
  for (const ch of inner) {
    if (quote !== "") {
      if (ch === quote) {
        quote = "";
      }
    } else if (ch === '"' || ch === "'") {
      quote = ch;
    } else if (ch === "(") {
      depth += 1;
    } else if (ch === ")") {
      depth -= 1;
      if (depth < 0) {
        return null;
      }
    }
  }
  return depth === 0 ? inner : null;
}

const STRING_LITERAL = /^"((?:[^"]|"")*)"$/;
const NUMBER_LITERAL = /^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/;

/** Thrown inside readCondition when the condition cannot be read. */
class Unreadable extends Error {}

/** A `table:condition` → its reading; null for one this model cannot read
 * (see the module header). */
function readCondition(condition: string | undefined, namespaces: ReadonlyMap<string, string>): ReadCondition | null {
  const text = condition?.trim() ?? "";
  if (text === "") {
    return { type: "any", operator: "between" };
  }
  const prefixed = /^([A-Za-z_][\w.-]*):(?!=)/.exec(text);
  const dialect = prefixed ? prefixed[1]! : "of";
  const uri = namespaces.get(dialect) ?? DIALECTS.get(dialect);
  if (uri !== OPENFORMULA_NAMESPACE && uri !== OOOC_NAMESPACE) {
    return null;
  }
  const body = (prefixed ? text.slice(prefixed[0].length) : text).trim();
  const operand = (raw: string | undefined): string | undefined => {
    const expression = raw?.trim() ?? "";
    if (expression === "") {
      return undefined;
    }
    try {
      return fromOdsFormula(`${dialect}:=${expression}`, namespaces);
    } catch (error) {
      if (error instanceof FormulaSyntaxError) {
        throw new Unreadable();
      }
      throw error;
    }
  };
  const read = (type: ValidationType, operator: ValidationOperator, formulas: (string | undefined)[]): ReadCondition => {
    const out: ReadCondition = { type, operator };
    const [formula1, formula2] = formulas.map(operand);
    if (formula1 !== undefined) {
      out.formula1 = formula1;
    }
    if (formula2 !== undefined) {
      out.formula2 = formula2;
    }
    return out;
  };
  const between = (type: ValidationType, operator: ValidationOperator, inner: string): ReadCondition => {
    const args = splitTopLevel(inner, ",;");
    if (args.length !== 2) {
      throw new Unreadable();
    }
    return read(type, operator, args);
  };
  const comparison = (type: ValidationType, rest: string): ReadCondition => {
    const match = COMPARISON.exec(rest.trim());
    if (!match) {
      throw new Unreadable();
    }
    return read(type, COMPARISONS[match[1]!]!, [match[2]]);
  };
  try {
    const list = callInner(body, "cell-content-is-in-list");
    if (list !== null) {
      return listCondition(list, operand);
    }
    const custom = callInner(body, "is-true-formula");
    if (custom !== null) {
      return read("custom", "between", [custom]);
    }
    for (const [name, operator] of [
      ["cell-content-text-length-is-between", "between"],
      ["cell-content-text-length-is-not-between", "notBetween"],
    ] as const) {
      const inner = callInner(body, name);
      if (inner !== null) {
        return between("textLength", operator, inner);
      }
    }
    const length = /^cell-content-text-length\(\s*\)([\s\S]*)$/.exec(body);
    if (length) {
      return comparison("textLength", length[1]!);
    }
    const typed = /^(cell-content-is-[a-z-]+)\(\s*\)\s*(?:and\s+([\s\S]*))?$/.exec(body);
    const type = typed ? TYPE_FUNCTIONS[typed[1]!] : undefined;
    if (!typed || type === undefined) {
      return null;
    }
    const rest = typed[2]?.trim();
    if (rest === undefined || rest === "") {
      return { type, operator: "between" };
    }
    for (const [name, operator] of [
      ["cell-content-is-between", "between"],
      ["cell-content-is-not-between", "notBetween"],
    ] as const) {
      const inner = callInner(rest, name);
      if (inner !== null) {
        return between(type, operator, inner);
      }
    }
    const content = /^cell-content\(\s*\)([\s\S]*)$/.exec(rest);
    return content ? comparison(type, content[1]!) : null;
  } catch (error) {
    if (error instanceof Unreadable) {
      return null;
    }
    throw error;
  }
}

/** `cell-content-is-in-list(…)`'s argument: quoted (or number) items joined
 * with `;` → a literal list; one other expression → a formula. */
function listCondition(inner: string, operand: (raw: string | undefined) => string | undefined): ReadCondition {
  const pieces = splitTopLevel(inner, ";").map((piece) => piece.trim());
  const literal = (piece: string): string | null => {
    const quoted = STRING_LITERAL.exec(piece);
    if (quoted) {
      return quoted[1]!.replaceAll('""', '"');
    }
    return NUMBER_LITERAL.test(piece) ? piece : null;
  };
  if (pieces.length > 1 || STRING_LITERAL.test(pieces[0]!)) {
    const items = pieces.map(literal);
    if (items.some((item) => item === null)) {
      throw new Unreadable();
    }
    return { type: "list", operator: "between", formula1: formatListLiteral(items as string[]), literal: true };
  }
  const formula1 = operand(pieces[0]);
  return formula1 === undefined ? { type: "list", operator: "between" } : { type: "list", operator: "between", formula1 };
}

/** One definition as read, independent of any sheet. */
export interface OdsValidationDefinition {
  name: string;
  /** The element verbatim. */
  xml: string;
  /** The base cell its relative operands are written at (null: none). */
  base: CellRef | null;
  /** Null when the condition did not read (the rule is then `any`). */
  condition: ReadCondition | null;
  /** The settings other than the condition's. */
  allowBlank: boolean;
  showDropdown: boolean;
  odsListOrder?: "sort-ascending";
  input: DataValidation["input"];
  error: DataValidation["error"];
}

/** A `table:content-validation` element → its definition (see the module
 * header); null when it is not one. */
function readDefinition(xml: string, namespaces: ReadonlyMap<string, string>): OdsValidationDefinition | null {
  const qualify = qualifier(namespaces);
  let wire: WireDefinition | null;
  try {
    wire = wireDefinition(xml, qualify);
  } catch (error) {
    if (error instanceof XmlScanError) {
      return null;
    }
    throw error;
  }
  if (!wire) {
    return null;
  }
  const attr = (local: string): string | undefined => wire.attrs.get(`table:${local}`);
  const display = attr("display-list") ?? "unsorted";
  const baseRef = attr("base-cell-address") === undefined ? null : parseOdsRangeAddress(attr("base-cell-address")!.trim());
  const definition: OdsValidationDefinition = {
    name: attr("name") ?? "",
    xml,
    base: baseRef && baseRef.start.row !== undefined && baseRef.start.col !== undefined ? { row: baseRef.start.row, col: baseRef.start.col } : null,
    condition: readCondition(attr("condition"), namespaces),
    allowBlank: attr("allow-empty-cell") !== "false",
    showDropdown: display !== "none",
    input: { show: false, title: "", text: "" },
    error: { show: false, style: "stop", title: "", text: "" },
  };
  if (display === "sort-ascending") {
    definition.odsListOrder = "sort-ascending";
  }
  if (wire.help) {
    definition.input = { show: wire.help.attrs.get("table:display") === "true", title: wire.help.attrs.get("table:title") ?? "", text: wire.help.text };
  }
  if (wire.error) {
    const type = wire.error.attrs.get("table:message-type") ?? "stop";
    definition.error = {
      show: wire.error.attrs.get("table:display") === "true",
      style: (ERROR_STYLES.has(type) ? type : "stop") as ValidationErrorStyle,
      title: wire.error.attrs.get("table:title") ?? "",
      text: wire.error.text,
    };
  }
  if (wire.macro) {
    // Kept in sourceXml and never run: the rule enforces nothing.
    definition.error = { ...definition.error, show: false };
  }
  return definition;
}

/** A relative operand written at `base`, rebased onto `origin`. */
function rebase(formula: string, base: CellRef, origin: CellRef): string {
  return shiftFormula(formula, origin.row - base.row, origin.col - base.col);
}

/** The rule a definition makes over `ranges` (see the module header). */
export function validationFromDefinition(definition: OdsValidationDefinition, cellRanges: GridRange[]): DataValidation {
  // The range whose top-left is the base cell leads (the writer's origin), so
  // a rule written from the model reads back with the same origin.
  const base = definition.base;
  const lead = base ? cellRanges.findIndex((range) => range.top === base.row && range.left === base.col) : -1;
  const ranges = lead > 0 ? [cellRanges[lead]!, ...cellRanges.slice(0, lead), ...cellRanges.slice(lead + 1)] : cellRanges;
  const origin = validationOrigin({ ranges });
  const rule: DataValidation = {
    id: mintValidationId(),
    ranges,
    type: "any",
    operator: "between",
    allowBlank: definition.allowBlank,
    showDropdown: definition.showDropdown,
    input: { ...definition.input },
    error: { ...definition.error },
    sourceXml: definition.xml,
    odsName: definition.name,
  };
  const condition = definition.condition;
  if (condition) {
    try {
      const formulas = [condition.formula1, condition.formula2].map((formula) =>
        formula === undefined || condition.literal ? formula : rebase(formula, base ?? origin, origin),
      );
      rule.type = condition.type;
      rule.operator = condition.operator;
      if (formulas[0] !== undefined) {
        rule.formula1 = formulas[0];
      }
      if (formulas[1] !== undefined) {
        rule.formula2 = formulas[1];
      }
    } catch (error) {
      if (!(error instanceof FormulaSyntaxError)) {
        throw error;
      }
    }
  }
  if (definition.odsListOrder !== undefined) {
    rule.odsListOrder = definition.odsListOrder;
  }
  return rule;
}

/** The definitions of a content prefix's `table:content-validations`, in
 * document order (namespaces from the prefix's root start tag). */
export function readOdsValidationDefinitions(prefix: string): { definitions: OdsValidationDefinition[]; leading: Map<string, string> } {
  const namespaces = rootNamespaces(prefix);
  const block = locateBlock(prefix, namespaces);
  const definitions: OdsValidationDefinition[] = [];
  const leading = new Map<string, string>();
  for (const each of block?.definitions ?? []) {
    if (each.name === "") {
      continue;
    }
    const definition = readDefinition(each.xml, namespaces);
    if (definition && definition.name !== "") {
      definitions.push(definition);
      leading.set(definition.name, each.leading);
    }
  }
  return { definitions, leading };
}

// --- ranges from cells ---------------------------------------------------------------------------

/** Accumulates the cells naming each validation into row-run × column-run
 * rectangles (see the module header): call startRow, cell for each named
 * cell run of the row, endRow; nothing is materialised per cell. */
export class ValidationRegions {
  private readonly regions = new Map<string, { ranges: GridRange[]; open: Map<string, GridRange> }>();
  private pending: { name: string; left: number; right: number }[] = [];
  private top = 0;
  private repeat = 1;

  /** A row (run) starting at `top`, `repeat` rows tall. */
  startRow(top: number, repeat: number): void {
    this.pending = [];
    this.top = top;
    this.repeat = Math.max(1, repeat);
  }

  /** `count` cells from column `left` naming `name`. */
  cell(name: string, left: number, count: number): void {
    if (left >= MAX_COLS || count <= 0) {
      return;
    }
    const right = Math.min(left + count, MAX_COLS) - 1;
    const last = this.pending.at(-1);
    if (last && last.name === name && last.right + 1 === left) {
      last.right = right;
    } else {
      this.pending.push({ name, left, right });
    }
  }

  /** The row (run) closing. */
  endRow(): void {
    if (this.top < MAX_ROWS) {
      const bottom = Math.min(this.top + this.repeat, MAX_ROWS) - 1;
      for (const { name, left, right } of this.pending) {
        let region = this.regions.get(name);
        if (!region) {
          region = { ranges: [], open: new Map() };
          this.regions.set(name, region);
        }
        const key = `${left}:${right}`;
        const above = region.open.get(key);
        if (above && above.bottom === this.top - 1) {
          above.bottom = bottom;
        } else {
          const range = { top: this.top, left, bottom, right };
          region.ranges.push(range);
          region.open.set(key, range);
        }
      }
    }
    this.pending = [];
  }

  /** Each name's rectangles, in first-met order. */
  result(): Map<string, GridRange[]> {
    return new Map([...this.regions].map(([name, region]) => [name, region.ranges]));
  }
}

/** Single cells naming validations → each name's rectangles (as the reader
 * accumulates them: horizontal runs, extended down while the same run
 * continues on the next row). */
export function rangesFromCellNames(cells: Iterable<{ row: number; col: number; name: string }>): Map<string, GridRange[]> {
  const sorted = [...cells].toSorted((a, b) => a.row - b.row || a.col - b.col);
  const regions = new ValidationRegions();
  let row = -1;
  for (const cell of sorted) {
    if (cell.row !== row) {
      if (row >= 0) {
        regions.endRow();
      }
      row = cell.row;
      regions.startRow(row, 1);
    }
    regions.cell(cell.name, cell.col, 1);
  }
  if (row >= 0) {
    regions.endRow();
  }
  return regions.result();
}

/** The validation names a table text's cells carry, as rectangles: its rows
 * (nested in header rows, groups and `table:table-rows`) and cells in
 * document order, repeats honoured, a table in its drawing layer skipped.
 * The text is a fragment, so names are matched under the table's own
 * prefix. */
export function tableValidationRanges(tableXml: string): Map<string, GridRange[]> {
  const regions = new ValidationRegions();
  if (!tableXml.includes("content-validation-name")) {
    return regions.result();
  }
  const qname = /<([^\s/>!?]+)/.exec(tableXml)?.[1] ?? "table:table";
  const prefix = qname.includes(":") ? qname.slice(0, qname.indexOf(":") + 1) : "";
  const attr = (attrs: readonly { name: string; value: string }[], local: string): string | undefined => attrs.find((each) => each.name === `${prefix}${local}`)?.value;
  const repeat = (attrs: readonly { name: string; value: string }[], local: string): number => {
    const value = Number(attr(attrs, local));
    return Number.isInteger(value) && value > 1 ? value : 1;
  };
  type Frame = "table" | "group" | "row" | "skip";
  const stack: Frame[] = [];
  let nextRow = 0;
  let col = 0;
  scanXml(tableXml, (event) => {
    if (event.type === "text") {
      return undefined;
    }
    if (event.type === "close") {
      if (stack.pop() === "row") {
        regions.endRow();
      }
      return undefined;
    }
    const parent = stack.at(-1);
    const local = event.name.startsWith(prefix) ? event.name.slice(prefix.length) : undefined;
    let frame: Frame = "skip";
    if (parent === undefined) {
      frame = "table";
    } else if (parent === "table" || parent === "group") {
      if (local === "table-row") {
        const rows = repeat(event.attrs, "number-rows-repeated");
        regions.startRow(nextRow, rows);
        nextRow += rows;
        col = 0;
        frame = "row";
        if (event.selfClosing) {
          regions.endRow();
        }
      } else if (local !== undefined && /^table-(?:rows|header-rows|row-group)$/.test(local)) {
        frame = "group";
      }
    } else if (parent === "row" && (local === "table-cell" || local === "covered-table-cell")) {
      const count = repeat(event.attrs, "number-columns-repeated");
      const name = attr(event.attrs, "content-validation-name");
      if (name !== undefined) {
        regions.cell(name, col, count);
      }
      col += count;
    }
    if (!event.selfClosing) {
      stack.push(frame);
    }
    return undefined;
  });
  return regions.result();
}

// --- the readers' entry points ------------------------------------------------------------------

/** Each sheet's validations from the content prefix's definitions and the
 * names its cells carry (`regions[i]` for sheet i), and the definitions no
 * cell references (verbatim, with the whitespace before each). */
export function readOdsContentValidations(prefix: string, regions: readonly ReadonlyMap<string, GridRange[]>[]): { sheets: DataValidation[][]; orphans: string[] } {
  const { definitions, leading } = readOdsValidationDefinitions(prefix);
  const sheets = regions.map((byName) =>
    definitions.flatMap((definition) => {
      const ranges = byName.get(definition.name);
      return ranges && ranges.length > 0 ? [validationFromDefinition(definition, ranges.map((range) => ({ ...range })))] : [];
    }),
  );
  const orphans = definitions
    .filter((definition) => !regions.some((byName) => (byName.get(definition.name)?.length ?? 0) > 0))
    .map((definition) => (leading.get(definition.name) ?? "") + definition.xml);
  return { sheets, orphans };
}

/** The validations of ONE table text, given the retained content prefix
 * (namespaces from its root start tag) — SPEC-130's reader upgrade. With
 * `rangesByName` (a validation name → the rectangles it covers, in the
 * sheet's coordinates) those ranges are used instead of the table's cells.
 * Touches nothing else. */
export function readTableValidations(prefix: string, tableXml: string, sheetName: string, rangesByName?: ReadonlyMap<string, GridRange[]>): DataValidation[] {
  void sheetName; // the rules are the table's own; the name does not change how they read
  const byName = rangesByName ?? tableValidationRanges(tableXml);
  return readOdsContentValidations(prefix, [byName]).sheets[0]!;
}

// --- writing -------------------------------------------------------------------------------------

/** XML-escaped paragraph text: a run of spaces other than one between two
 * other characters as `text:s`, a tab as `text:tab`. */
function lineXml(line: string): string {
  let out = "";
  let i = 0;
  while (i < line.length) {
    const ch = line[i]!;
    if (ch === " ") {
      let end = i;
      while (end < line.length && line[end] === " ") {
        end += 1;
      }
      const count = end - i;
      out += count === 1 && i > 0 && end < line.length ? " " : count === 1 ? "<text:s/>" : `<text:s text:c="${count}"/>`;
      i = end;
    } else if (ch === "\t") {
      out += "<text:tab/>";
      i += 1;
    } else {
      let end = i;
      while (end < line.length && line[end] !== " " && line[end] !== "\t") {
        end += 1;
      }
      out += line.slice(i, end).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\r/g, "&#13;");
      i = end;
    }
  }
  return out;
}

/** A message's text as `text:p` paragraphs, one per line ("" → none). */
function paragraphsXml(text: string): string {
  const clean = sanitizeXmlText(text);
  if (clean === "") {
    return "";
  }
  return clean
    .split("\n")
    .map((line) => (line === "" ? "<text:p/>" : `<text:p>${lineXml(line)}</text:p>`))
    .join("");
}

/** A canonical formula as a condition operand: OpenFormula without `of:=`. */
function odsOperand(canonical: string | undefined): string {
  return canonical === undefined ? "" : toOdsFormula(canonical.replace(/^=/, "")).slice("of:=".length);
}

/** A rule's `table:condition` ("" for none). */
function conditionOf(rule: ValidationSettings): string {
  const a = odsOperand(rule.formula1);
  const b = odsOperand(rule.formula2);
  switch (rule.type) {
    case "any":
      return "";
    case "list": {
      const items = rule.formula1 === undefined ? null : parseListLiteral(rule.formula1);
      const inner = items ? items.map((item) => `"${item.replaceAll('"', '""')}"`).join(";") : a;
      return `of:cell-content-is-in-list(${inner})`;
    }
    case "custom":
      return `of:is-true-formula(${a})`;
    case "textLength":
      if (isRangeOperator(rule.operator)) {
        return `of:cell-content-text-length-is-${rule.operator === "between" ? "" : "not-"}between(${a},${b})`;
      }
      return `of:cell-content-text-length()${COMPARISON_WRITE[rule.operator as keyof typeof COMPARISON_WRITE]}${a}`;
    default: {
      const type = `of:${TYPE_FUNCTION_WRITE[rule.type]}()`;
      if (rule.formula1 === undefined && rule.formula2 === undefined) {
        return type;
      }
      if (isRangeOperator(rule.operator)) {
        return `${type} and cell-content-is-${rule.operator === "between" ? "" : "not-"}between(${a},${b})`;
      }
      return `${type} and cell-content()${COMPARISON_WRITE[rule.operator as keyof typeof COMPARISON_WRITE]}${a}`;
    }
  }
}

/** A rule serialised as `table:content-validation` named `name`, its base
 * cell the origin on `sheetName` (see the module header). */
export function odsValidationXml(name: string, rule: DataValidation, sheetName: string): string {
  const origin = validationOrigin(rule);
  const condition = conditionOf(rule);
  let attributes = ` table:name="${escapeAttributeValue(name)}"`;
  if (condition !== "") {
    attributes += ` table:condition="${escapeAttributeValue(condition)}"`;
  }
  if (condition !== "" || !rule.allowBlank) {
    attributes += ` table:allow-empty-cell="${rule.allowBlank}"`;
  }
  const display = !rule.showDropdown ? "none" : rule.odsListOrder === "sort-ascending" ? "sort-ascending" : "unsorted";
  if (rule.type === "list" || display !== "unsorted") {
    attributes += ` table:display-list="${display}"`;
  }
  // LibreOffice's form: the sheet without `$` (`Form.B2`).
  const base = formatOdsRangeAddress({ sheet: sheetName, start: { row: origin.row, col: origin.col, rowAbs: false, colAbs: false } }).slice(1);
  attributes += ` table:base-cell-address="${escapeAttributeValue(base)}"`;
  let children = "";
  const input = rule.input;
  if (input.show || input.title !== "" || input.text !== "") {
    children +=
      `<table:help-message${input.title === "" ? "" : ` table:title="${escapeAttributeValue(input.title)}"`} table:display="${input.show}">` +
      `${paragraphsXml(input.text)}</table:help-message>`;
  }
  const error = rule.error;
  if (error.show || error.title !== "" || error.text !== "" || error.style !== "stop") {
    children +=
      `<table:error-message table:message-type="${error.style}"${error.title === "" ? "" : ` table:title="${escapeAttributeValue(error.title)}"`} table:display="${error.show}">` +
      `${paragraphsXml(error.text)}</table:error-message>`;
  }
  return children === "" ? `<table:content-validation${attributes}/>` : `<table:content-validation${attributes}>${children}</table:content-validation>`;
}

/** Whether a rule's `sourceXml` is an ODS definition (a model read from
 * another format carries none). */
function isOdsSource(sourceXml: string | undefined): sourceXml is string {
  return sourceXml !== undefined && /^\s*<[\w.-]+:content-validation[\s/>]/.test(sourceXml);
}

const NAME_ATTRIBUTE = /\stable:name\s*=\s*"([^"]*)"/;
const CELL_NAMES = /[\s:]content-validation-name\s*=\s*"([^"]*)"/g;

/** A definition the block will hold. */
interface Claim {
  name: string;
  xml: string;
  /** Retained text (passed by the sheet-name pass), not generated. */
  retained: boolean;
  /** How it reads (for sharing). */
  definition: OdsValidationDefinition | null;
}

/** One save's content validations (see the module header). */
export class OdsValidationWriter {
  private readonly claims = new Map<string, Claim>();
  private readonly order: string[] = [];
  private readonly names = new Map<string, string>();
  private readonly rebuilt = new Set<Sheet>();
  private readonly taken = new Set<string>();
  private readonly orphans: { name: string; xml: string }[];
  private readonly namespaces: ReadonlyMap<string, string>;
  private readonly source: SourceBlock | null;
  private counter = 0;

  /** `prefix` is the retained content prefix (undefined from scratch);
   * `keepsSource(sheet)` whether the caller would emit a sheet's table from
   * its retained text as it is. */
  constructor(model: Pick<SpreadsheetModel, "sheets" | "odsOrphanValidations">, prefix: string | undefined, keepsSource: (sheet: Sheet) => boolean) {
    this.namespaces = prefix === undefined ? new Map() : rootNamespaces(prefix);
    this.source = prefix === undefined ? null : locateBlock(prefix, this.namespaces);
    this.orphans = (model.odsOrphanValidations ?? []).map((xml) => ({ name: decodeXmlText(NAME_ATTRIBUTE.exec(xml)?.[1] ?? ""), xml }));
    // Every name in use: the source's definitions, the orphans', the rules',
    // and the names the kept tables carry.
    for (const each of this.source?.definitions ?? []) {
      this.taken.add(each.name);
    }
    for (const orphan of this.orphans) {
      this.taken.add(orphan.name);
      this.claims.set(orphan.name, { name: orphan.name, xml: orphan.xml, retained: true, definition: null });
    }
    const kept = model.sheets.filter(keepsSource);
    for (const sheet of model.sheets) {
      for (const rule of sheetValidations(sheet)) {
        if (rule.odsName !== undefined) {
          this.taken.add(rule.odsName);
        }
      }
    }
    for (const sheet of kept) {
      for (const match of (sheet.sourceXml ?? "").matchAll(CELL_NAMES)) {
        this.taken.add(decodeXmlText(match[1]!));
      }
    }
    for (const sheet of kept) {
      if (!this.keep(sheet)) {
        this.rebuilt.add(sheet);
      }
    }
    for (const sheet of model.sheets) {
      if (!keepsSource(sheet) || this.rebuilt.has(sheet)) {
        this.assign(sheet);
      }
    }
  }

  /** Whether a sheet the caller would keep must be rebuilt instead (its
   * rules cannot keep the names its retained table carries). */
  rebuilds(sheet: Sheet): boolean {
    return this.rebuilt.has(sheet);
  }

  /** The name a sheet's rule is written under. */
  nameOf(sheet: Sheet, rule: DataValidation): string | undefined {
    return this.names.get(`${sheet.id}\u0000${rule.id}`);
  }

  /** The generated (not retained) text the block holds — for the root's
   * namespace declarations. */
  get generatedXml(): string {
    return [...this.claims.values()]
      .filter((claim) => !claim.retained)
      .map((claim) => claim.xml)
      .join("");
  }

  /** The block (see the module header); "" when it holds no definition.
   * `pass` is the sheet-name pass for retained text. */
  blockXml(pass: (text: string) => string = (text) => text): string {
    if (this.claims.size === 0) {
      return "";
    }
    const emitted = new Set<string>();
    let body = "";
    const emit = (claim: Claim, leading: string): void => {
      emitted.add(claim.name);
      const isOrphan = this.orphans.some((orphan) => orphan.name === claim.name);
      body += isOrphan ? pass(claim.xml) : leading + (claim.retained ? pass(claim.xml) : claim.xml);
    };
    for (const each of this.source?.definitions ?? []) {
      const claim = this.claims.get(each.name);
      if (each.name !== "" && claim && !emitted.has(each.name)) {
        emit(claim, each.leading);
      }
    }
    for (const orphan of this.orphans) {
      if (!emitted.has(orphan.name)) {
        emit(this.claims.get(orphan.name)!, "");
      }
    }
    for (const name of this.order) {
      if (!emitted.has(name)) {
        emit(this.claims.get(name)!, "");
      }
    }
    const open = this.source?.open ?? "<table:content-validations>";
    const close = this.source?.close ?? "</table:content-validations>";
    return open + body + close;
  }

  /** A content prefix (the retained one, the sheet-name pass applied) with
   * its block regenerated: replaced where it was, removed when empty, else
   * inserted after `table:calculation-settings` — before
   * `table:label-ranges` and the first table. */
  splicePrefix(prefix: string, pass: (text: string) => string = (text) => text): string {
    const block = this.blockXml(pass);
    const found = locateBlock(prefix, rootNamespaces(prefix));
    if (found) {
      return prefix.slice(0, found.start) + block + prefix.slice(found.end);
    }
    if (block === "") {
      return prefix;
    }
    const at = insertionPoint(prefix, rootNamespaces(prefix));
    return prefix.slice(0, at) + block + prefix.slice(at);
  }

  // --- claims ---

  /** Claims a kept sheet's names; false (nothing claimed) when one cannot
   * be kept. */
  private keep(sheet: Sheet): boolean {
    const made: string[] = [];
    const named: [string, string][] = [];
    for (const rule of sheetValidations(sheet)) {
      const name = rule.odsName;
      if (name === undefined || rule.ranges.length === 0) {
        return this.rollback(made);
      }
      const existing = this.claims.get(name);
      if (existing) {
        if (!this.shares(existing, rule)) {
          return this.rollback(made);
        }
      } else {
        this.claim(this.candidate(name, rule, sheet.name));
        made.push(name);
      }
      named.push([`${sheet.id}\u0000${rule.id}`, name]);
    }
    for (const [key, name] of named) {
      this.names.set(key, name);
    }
    return true;
  }

  private rollback(made: readonly string[]): false {
    for (const name of made) {
      this.claims.delete(name);
      this.order.splice(this.order.indexOf(name), 1);
    }
    return false;
  }

  /** Names a rebuilt or new sheet's rules (see the module header). */
  private assign(sheet: Sheet): void {
    for (const rule of sheetValidations(sheet)) {
      if (rule.ranges.length === 0) {
        continue;
      }
      let name: string | undefined;
      if (rule.odsName !== undefined && isOdsSource(rule.sourceXml)) {
        const existing = this.claims.get(rule.odsName);
        if (!existing) {
          this.claim(this.candidate(rule.odsName, rule, sheet.name));
          name = rule.odsName;
        } else if (this.shares(existing, rule)) {
          name = rule.odsName;
        }
      }
      if (name === undefined) {
        name = this.freshName();
        this.claim({ name, xml: odsValidationXml(name, rule, sheet.name), retained: false, definition: null });
      }
      this.names.set(`${sheet.id}\u0000${rule.id}`, name);
    }
  }

  /** The definition a rule claims its own name with: its source while it
   * has one, else serialised from the model. */
  private candidate(name: string, rule: DataValidation, sheetName: string): Claim {
    return isOdsSource(rule.sourceXml)
      ? { name, xml: rule.sourceXml.trimStart(), retained: true, definition: null }
      : { name, xml: odsValidationXml(name, rule, sheetName), retained: false, definition: null };
  }

  private claim(claim: Claim): void {
    this.claims.set(claim.name, claim);
    this.order.push(claim.name);
  }

  /** Whether a claimed definition reads back as `rule` on its sheet. */
  private shares(claim: Claim, rule: DataValidation): boolean {
    if (this.orphans.some((orphan) => orphan.name === claim.name)) {
      return false;
    }
    if (claim.retained && claim.xml === rule.sourceXml?.trimStart()) {
      return true;
    }
    claim.definition ??= readDefinition(claim.xml, this.namespaces);
    if (!claim.definition) {
      return false;
    }
    const read = validationFromDefinition(claim.definition, rule.ranges);
    return sameValidationSettings(read, rule);
  }

  /** A `de-val<n>` no name in use has. */
  private freshName(): string {
    let name: string;
    do {
      this.counter += 1;
      name = `de-val${this.counter}`;
    } while (this.taken.has(name) || this.claims.has(name));
    this.taken.add(name);
    return name;
  }
}

/** Where a new block goes in a content prefix: after
 * `table:calculation-settings`, else before `table:label-ranges`, else at
 * the end (the first table follows). */
function insertionPoint(prefix: string, namespaces: ReadonlyMap<string, string>): number {
  const qualify = qualifier(namespaces);
  let at = prefix.length;
  let depth = 0;
  let settingsDepth = -1;
  try {
    scanXml(prefix, (event) => {
      if (event.type === "text") {
        return undefined;
      }
      if (event.type === "close") {
        if (depth === settingsDepth) {
          at = event.end;
          return "stop";
        }
        depth -= 1;
        return undefined;
      }
      const name = qualify(event.name);
      depth += 1;
      if (name === "table:calculation-settings") {
        if (event.selfClosing) {
          at = event.end;
          return "stop";
        }
        settingsDepth = depth;
      } else if (name === "table:label-ranges" && settingsDepth < 0) {
        at = event.start;
        return "stop";
      }
      if (event.selfClosing) {
        depth -= 1;
      }
      return undefined;
    });
  } catch (error) {
    if (!(error instanceof XmlScanError)) {
      throw error;
    }
  }
  return at;
}

/** A cell's verbatim preserved attributes without a stale
 * `table:content-validation-name` (records stored before SPEC-130 captured
 * it; the model's name wins). */
export function withoutValidationName(attrs: string): string {
  return attrs.includes("content-validation-name") ? attrs.replace(/\s+[\w.-]*:?content-validation-name\s*=\s*(?:"[^"]*"|'[^']*')/g, "") : attrs;
}
