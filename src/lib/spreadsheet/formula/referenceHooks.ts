// SPEC-133's evaluator hooks (SPEC-133 "Resolution") — the `resolveName` and
// `resolveStructured` members SPEC-120's EvalContext leaves to it, and the
// static areas the dependency graph learns from names and tables.
//
// - resolveName finds the DefinedName (names.ts resolveDefinedName; a name
//   that is no defined name but a table's is the table's data body) and
//   evaluates its formula — parsed once per text — with origin A1 and the
//   caller's host, so a relative part means `part + (host − A1)`, wrapping
//   modulo the grid instead of becoming #REF! (Excel's relative names). A
//   name evaluates in array mode, as Excel evaluates names: `=SUM(Doubled)`
//   over `Doubled = $A$1:$A$3*2` sums the array; a value position still
//   intersects implicitly. An unknown, unparsed or placeholder name is
//   #NAME?.
// - Termination: each top-level name lookup starts a chain. A name already
//   being evaluated in the chain (a cycle: `Foo = Foo+Foo`, `Foo = Bar+Bar`
//   with `Bar = Foo+Foo`) is #NAME? at once; a chain nesting past
//   MAX_NAME_DEPTH (64) names is #NAME?; and within a chain each name's
//   value is computed once and reused — the host is the same throughout —
//   so a name used twice by every name below it costs one evaluation per
//   name, never 2^depth.
// - Both hooks call ctx.recordDependency for every area they resolve to; the
//   graph also learns those areas statically (referenceAreas), so a formula
//   using a name recalculates when the name's cells change even before its
//   first evaluation.
import type { DefinedName, SpreadsheetModel } from "../model";
import { findTable, dataBody, resolveStructuredReference } from "../tables";
import { resolveDefinedName } from "../names";
import type { NodeOf } from "./ast";
import { resolveReference } from "./areas";
import { evaluate, formulaFacts, type EvalContext } from "./evaluate";
import { FormulaSyntaxError, parseFormulaAt } from "./parse";
import { isError, isRange, NAME_ERROR, type Area, type EvalValue } from "./values";

/** Names nest this deep before #NAME? (SPEC-133). */
export const MAX_NAME_DEPTH = 64;

const A1 = { row: 0, col: 0 };

export type BaseContext = Omit<EvalContext, "host" | "origin" | "arrayMode">;

/** One name chain's state: the names being evaluated and the values found. */
interface Chain {
  active: Set<DefinedName>;
  memo: Map<DefinedName, EvalValue>;
  depth: number;
}

/** A defined name's formula parsed relative to A1, or null when it does not
 * parse (parseFormulaAt caches per text). */
function parsedName(name: DefinedName): Exclude<ReturnType<typeof parseFormulaAt>, FormulaSyntaxError> | null {
  if (name.formulaUnparsed) {
    return null;
  }
  const parsed = parseFormulaAt(name.formula, A1);
  return parsed instanceof FormulaSyntaxError ? null : parsed;
}

/** A table's data body as a range, when `name` is a table's name. */
function tableByName(model: SpreadsheetModel, name: string): EvalValue | undefined {
  const found = findTable(model, name);
  return found ? { kind: "range", areas: [{ sheetId: found.sheet.id, ...dataBody(found.table) }] } : undefined;
}

function recordAreas(ctx: Pick<EvalContext, "recordDependency">, value: EvalValue): void {
  if (isRange(value)) {
    for (const area of value.areas) {
      ctx.recordDependency(area);
    }
  }
}

/** The resolveName hook of one chain. */
function nameResolver(base: BaseContext, chain: Chain | null): NonNullable<EvalContext["resolveName"]> {
  return (name, sheetQualifier, host) => {
    const defined = resolveDefinedName(base.model, name, sheetQualifier, host.sheetId);
    if (!defined) {
      const table = sheetQualifier === undefined ? tableByName(base.model, name) : undefined;
      if (table) {
        recordAreas(base, table);
        return table;
      }
      return NAME_ERROR;
    }
    const current: Chain = chain ?? { active: new Set(), memo: new Map(), depth: 0 };
    const known = current.memo.get(defined);
    if (known !== undefined) {
      recordAreas(base, known);
      return known;
    }
    if (current.active.has(defined) || current.depth >= MAX_NAME_DEPTH) {
      return NAME_ERROR;
    }
    const parsed = parsedName(defined);
    if (!parsed) {
      return NAME_ERROR;
    }
    current.active.add(defined);
    current.depth += 1;
    let value: EvalValue;
    try {
      const ctx: EvalContext = {
        ...base,
        resolveName: nameResolver(base, current),
        resolveStructured: structuredResolver(base),
        host,
        origin: parsed.origin,
        arrayMode: true,
        wrapReferences: true,
      };
      value = evaluate(parsed.node, ctx);
    } finally {
      current.depth -= 1;
      current.active.delete(defined);
    }
    current.memo.set(defined, value);
    recordAreas(base, value);
    return value;
  };
}

function structuredResolver(base: BaseContext): NonNullable<EvalContext["resolveStructured"]> {
  return (node, host) => {
    const value = resolveStructuredReference(base.model, node, host);
    recordAreas(base, value);
    return value;
  };
}

/** `base` with SPEC-133's hooks: names and structured references evaluate. */
export function withReferenceHooks<T extends BaseContext>(base: T): T {
  return { ...base, resolveName: nameResolver(base, null), resolveStructured: structuredResolver(base) };
}

/** A defined name's current value at `host` (the Name Manager's column):
 * its formula evaluated as a cell using it would see it. */
export function evaluateDefinedName(base: BaseContext, name: DefinedName, host: { sheetId: string; row: number; col: number }): EvalValue {
  const hooked = withReferenceHooks(base);
  const chain: Chain = { active: new Set(), memo: new Map(), depth: 0 };
  const parsed = parsedName(name);
  if (!parsed) {
    return NAME_ERROR;
  }
  chain.active.add(name);
  chain.depth = 1;
  return evaluate(parsed.node, {
    ...hooked,
    resolveName: nameResolver(base, chain),
    host,
    origin: parsed.origin,
    arrayMode: true,
    wrapReferences: true,
  });
}

// --- static areas (the dependency graph) -------------------------------------------------------

/** The areas a name's formula reads at `host`, nested names and table
 * references included — each name visited once, MAX_NAME_DEPTH deep. */
function nameAreas(model: SpreadsheetModel, defined: DefinedName, host: { sheetId: string; row: number; col: number }, seen: Set<DefinedName>, out: Area[]): void {
  if (seen.has(defined) || seen.size >= MAX_NAME_DEPTH) {
    return;
  }
  seen.add(defined);
  const parsed = parsedName(defined);
  if (!parsed) {
    return;
  }
  const facts = formulaFacts(parsed.node);
  for (const ref of facts.refs) {
    const resolved = resolveReference(ref, { model, host, origin: parsed.origin, wrapReferences: true });
    if (!isError(resolved)) {
      out.push(...resolved.areas);
    }
  }
  collectNodes(model, facts.names, facts.structured, host, seen, out);
}

function collectNodes(
  model: SpreadsheetModel,
  names: readonly NodeOf<"name">[],
  structured: readonly NodeOf<"structured">[],
  host: { sheetId: string; row: number; col: number },
  seen: Set<DefinedName>,
  out: Area[],
): void {
  for (const node of names) {
    const defined = resolveDefinedName(model, node.name, node.sheet, host.sheetId);
    if (defined) {
      nameAreas(model, defined, host, seen, out);
    } else if (node.sheet === undefined) {
      const table = findTable(model, node.name);
      if (table) {
        out.push({ sheetId: table.sheet.id, ...dataBody(table.table) });
      }
    }
  }
  for (const node of structured) {
    const resolved = resolveStructuredReference(model, node, host);
    if (!isError(resolved)) {
      out.push(...resolved.areas);
    }
  }
}

/** The areas a formula's names and structured references read at `host`
 * — what the dependency graph adds to its static precedents. */
export function referenceAreas(
  model: SpreadsheetModel,
  facts: { names: readonly NodeOf<"name">[]; structured: readonly NodeOf<"structured">[] },
  host: { sheetId: string; row: number; col: number },
): Area[] {
  if (facts.names.length === 0 && facts.structured.length === 0) {
    return [];
  }
  const out: Area[] = [];
  collectNodes(model, facts.names, facts.structured, host, new Set(), out);
  return out;
}
