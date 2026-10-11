// Information, SPEC-127's library: ISEVEN ISODD ISREF ISFORMULA TYPE
// ERROR.TYPE N SHEET SHEETS CELL INFO.
//
// - TYPE: 1 number or blank, 2 text, 4 logical, 16 error, 64 array (a
//   multi-cell reference included). ERROR.TYPE: #NULL! 1, #DIV/0! 2, #VALUE!
//   3, #REF! 4, #NAME? 5, #NUM! 6, #N/A 7, #GETTING_DATA 8, and the spec's
//   pins #SPILL! 9, #CONNECT! 10, #BLOCKED! 11, #UNKNOWN! 12, #FIELD! 13,
//   #CALC! 14 (Microsoft's page lists 1–8 only); anything else #N/A. Both
//   accept errors.
// - SHEET is a sheet's 1-based position in the workbook and SHEETS a count,
//   hidden and very hidden sheets counted.
// - CELL (volatile) supports "address" (references on the host sheet),
//   "col", "row", "contents", "type", "filename" (`[<title>.<ext>]<sheet>`:
//   no folder exists here, and MID(CELL("filename",A1),…) still finds the
//   sheet name), "protect" and "prefix"; INFO (volatile) supports "numfile"
//   (grid sheets in the workbook) and "recalc" ("Automatic"). Every other
//   info type Excel knows is a valid call this library does not compute
//   (UnsupportedCaseError) — INFO("system") included, since the engine may
//   not read the platform and a value that differs per machine would differ
//   per save; an unknown one is #VALUE!. Without a reference CELL reads the
//   formula's own cell.
import { colToLetters, getCell } from "../../address";
import type { CellError } from "../../model";
import { cellStyle, effectiveStyleId } from "../../styles";
import { sheetById, sheetIndexByName, sheetIndexOf } from "../areas";
import type { EvalContext } from "../evaluate";
import { UnsupportedCaseError, type ArgList, type FunctionSpec } from "../registry";
import { isArray, isEmpty, isError, isRange, NA_ERROR, REF_ERROR, toNumber, VALUE_ERROR, type Area, type ScalarValue } from "../values";
import { arg, optional, scalarArg, textArg } from "./common";

/** ERROR.TYPE's numbers (see the module header). */
const ERROR_NUMBERS: Readonly<Record<string, number>> = {
  "#NULL!": 1,
  "#DIV/0!": 2,
  "#VALUE!": 3,
  "#REF!": 4,
  "#NAME?": 5,
  "#NUM!": 6,
  "#N/A": 7,
  "#GETTING_DATA": 8,
  "#SPILL!": 9,
  "#CONNECT!": 10,
  "#BLOCKED!": 11,
  "#UNKNOWN!": 12,
  "#FIELD!": 13,
  "#CALC!": 14,
};

const CELL_SUPPORTED = new Set(["address", "col", "row", "contents", "type", "filename", "protect", "prefix"]);
const CELL_KNOWN = new Set([...CELL_SUPPORTED, "color", "format", "parentheses", "width"]);
const INFO_KNOWN = new Set(["directory", "numfile", "origin", "osversion", "recalc", "release", "system"]);

/** ISEVEN/ISODD's integer: numbers and numeric text, truncated. */
function parityArg(args: ArgList, ctx: EvalContext): number | CellError {
  const value = scalarArg(args, 0);
  if (typeof value === "boolean") {
    return VALUE_ERROR;
  }
  const n = toNumber(value, ctx.settings);
  return isError(n) ? n : Math.trunc(n);
}

/** TYPE of a scalar. */
function typeOf(value: ScalarValue): number {
  if (isError(value)) {
    return 16;
  }
  return typeof value === "string" ? 2 : typeof value === "boolean" ? 4 : 1;
}

/** CELL's "prefix": the alignment mark of a text cell. */
function prefixOf(ctx: EvalContext, area: Area, value: ScalarValue): string {
  if (typeof value !== "string") {
    return "";
  }
  const sheet = sheetById(ctx.model, area.sheetId);
  const horizontal = sheet ? cellStyle(ctx.model, effectiveStyleId(sheet, area.top, area.left)).alignment?.horizontal : undefined;
  switch (horizontal) {
    case "right":
      return '"';
    case "center":
      return "^";
    case "fill":
      return "\\";
    case undefined:
    case "general":
    case "left":
      return "'";
    default:
      return "";
  }
}

export const INFORMATION_EXTRA_FUNCTIONS: FunctionSpec[] = [
  {
    name: "ISEVEN",
    category: "Information",
    minArgs: 1,
    maxArgs: 1,
    argKinds: ["value"],
    summary: "Returns TRUE when a number is even.",
    args: [arg("number", "The number to test, truncated.")],
    impl(args, ctx) {
      const n = parityArg(args, ctx);
      return isError(n) ? n : n % 2 === 0;
    },
  },
  {
    name: "ISODD",
    category: "Information",
    minArgs: 1,
    maxArgs: 1,
    argKinds: ["value"],
    summary: "Returns TRUE when a number is odd.",
    args: [arg("number", "The number to test, truncated.")],
    impl(args, ctx) {
      const n = parityArg(args, ctx);
      return isError(n) ? n : n % 2 !== 0;
    },
  },
  {
    name: "ISREF",
    category: "Information",
    minArgs: 1,
    maxArgs: 1,
    argKinds: ["any"],
    acceptsErrors: true,
    summary: "Returns TRUE when the value is a reference.",
    args: [arg("value", "The value to test.")],
    impl(args) {
      return isRange(args.get(0));
    },
  },
  {
    name: "ISFORMULA",
    category: "Information",
    minArgs: 1,
    maxArgs: 1,
    argKinds: ["range"],
    summary: "Returns TRUE when a referenced cell holds a formula.",
    args: [arg("reference", "The cell to test (a range's top-left cell).")],
    impl(args, ctx) {
      const value = args.get(0);
      if (!isRange(value)) {
        return isError(value) ? value : VALUE_ERROR;
      }
      const area = value.areas[0]!;
      const sheet = sheetById(ctx.model, area.sheetId);
      return sheet ? getCell(sheet, area.top, area.left)?.formula !== undefined : false;
    },
  },
  {
    name: "TYPE",
    category: "Information",
    minArgs: 1,
    maxArgs: 1,
    argKinds: ["any"],
    acceptsErrors: true,
    summary: "Returns a number for the type of a value: 1, 2, 4, 16 or 64.",
    args: [arg("value", "The value whose type to return.")],
    impl(args, ctx) {
      const value = args.get(0);
      if (isArray(value)) {
        return 64;
      }
      if (isRange(value)) {
        const area = value.areas[0]!;
        if (value.areas.length > 1 || area.top !== area.bottom || area.left !== area.right) {
          return value.areas.length > 1 ? 16 : 64;
        }
        return typeOf(ctx.valueAt(area.sheetId, area.top, area.left));
      }
      return typeOf(value);
    },
  },
  {
    name: "ERROR.TYPE",
    category: "Information",
    minArgs: 1,
    maxArgs: 1,
    argKinds: ["value"],
    acceptsErrors: true,
    summary: "Returns the number of an error value, #N/A for anything else.",
    args: [arg("error_val", "The error value to identify.")],
    impl(args) {
      const value = scalarArg(args, 0);
      return isError(value) ? (ERROR_NUMBERS[value.error] ?? NA_ERROR) : NA_ERROR;
    },
  },
  {
    name: "N",
    category: "Information",
    minArgs: 1,
    maxArgs: 1,
    argKinds: ["value"],
    summary: "Returns a value as a number: TRUE is 1, text and FALSE are 0.",
    args: [arg("value", "The value to convert.")],
    impl(args) {
      const value = scalarArg(args, 0);
      return typeof value === "number" ? value : value === true ? 1 : 0;
    },
  },
  {
    name: "SHEET",
    category: "Information",
    minArgs: 0,
    maxArgs: 1,
    argKinds: ["any"],
    summary: "Returns the position of a sheet in the workbook.",
    args: [optional("value", "A reference or a sheet name (default the formula's own sheet).")],
    impl(args, ctx) {
      if (args.length === 0 || args.isMissing(0)) {
        return sheetIndexOf(ctx.model, ctx.host.sheetId) + 1;
      }
      const value = args.get(0);
      if (isRange(value)) {
        return sheetIndexOf(ctx.model, value.areas[0]!.sheetId) + 1;
      }
      if (typeof value === "string") {
        const index = sheetIndexByName(ctx.model, value);
        return index < 0 ? NA_ERROR : index + 1;
      }
      return isError(value) ? value : NA_ERROR;
    },
  },
  {
    name: "SHEETS",
    category: "Information",
    minArgs: 0,
    maxArgs: 1,
    argKinds: ["any"],
    summary: "Returns the number of sheets in a reference or in the workbook.",
    args: [optional("reference", "A reference, perhaps 3-D (default the whole workbook).")],
    impl(args, ctx) {
      if (args.length === 0 || args.isMissing(0)) {
        return ctx.model.sheets.length;
      }
      const value = args.get(0);
      if (isRange(value)) {
        return new Set(value.areas.map((area) => area.sheetId)).size;
      }
      return isError(value) ? value : REF_ERROR;
    },
  },
  {
    name: "CELL",
    category: "Information",
    minArgs: 1,
    maxArgs: 2,
    argKinds: ["value", "range"],
    volatile: true,
    summary: "Returns information about a cell: its address, row, contents, type and more.",
    args: [
      arg("info_type", "\"address\", \"col\", \"row\", \"contents\", \"type\", \"filename\", \"protect\" or \"prefix\"."),
      optional("reference", "The cell to describe (default the formula's own cell)."),
    ],
    impl(args, ctx) {
      const type = textArg(args, 0);
      if (isError(type)) {
        return type;
      }
      const info = type.toLowerCase();
      if (!CELL_KNOWN.has(info)) {
        return VALUE_ERROR;
      }
      if (!CELL_SUPPORTED.has(info)) {
        throw new UnsupportedCaseError(`CELL("${info}") isn't computed`);
      }
      let area: Area = { sheetId: ctx.host.sheetId, top: ctx.host.row, left: ctx.host.col, bottom: ctx.host.row, right: ctx.host.col };
      if (args.length > 1 && !args.isMissing(1)) {
        const value = args.get(1);
        if (!isRange(value)) {
          return isError(value) ? value : VALUE_ERROR;
        }
        area = value.areas[0]!;
      }
      const sheet = sheetById(ctx.model, area.sheetId);
      const value = ctx.valueAt(area.sheetId, area.top, area.left);
      switch (info) {
        case "address":
          if (area.sheetId !== ctx.host.sheetId) {
            throw new UnsupportedCaseError("CELL(\"address\") of another sheet isn't computed");
          }
          return `$${colToLetters(area.left)}$${area.top + 1}`;
        case "col":
          return area.left + 1;
        case "row":
          return area.top + 1;
        case "contents":
          return isEmpty(value) ? 0 : value;
        case "type":
          return isEmpty(value) ? "b" : typeof value === "string" ? "l" : "v";
        case "filename":
          return `[${ctx.model.title}.${ctx.model.format}]${sheet?.name ?? ""}`;
        case "protect":
          return sheet && cellStyle(ctx.model, effectiveStyleId(sheet, area.top, area.left)).protection?.locked === false ? 0 : 1;
        default:
          return prefixOf(ctx, area, value);
      }
    },
  },
  {
    name: "INFO",
    category: "Information",
    minArgs: 1,
    maxArgs: 1,
    argKinds: ["value"],
    volatile: true,
    summary: "Returns information about the workbook's environment.",
    args: [arg("type_text", "\"numfile\" (the number of sheets) or \"recalc\" (the calculation mode).")],
    impl(args, ctx) {
      const type = textArg(args, 0);
      if (isError(type)) {
        return type;
      }
      const info = type.toLowerCase();
      if (!INFO_KNOWN.has(info)) {
        return VALUE_ERROR;
      }
      if (info === "numfile") {
        return ctx.model.sheets.filter((sheet) => sheet.kind === "grid").length;
      }
      if (info === "recalc") {
        return "Automatic";
      }
      throw new UnsupportedCaseError(`INFO("${info}") isn't computed`);
    },
  },
];
