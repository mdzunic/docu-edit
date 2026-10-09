// Text (SPEC-120 "The core functions"): CONCAT CONCATENATE TEXTJOIN LEN LEFT
// RIGHT MID UPPER LOWER PROPER TRIM SUBSTITUTE REPLACE FIND SEARCH TEXT VALUE
// REPT EXACT.
//
// - Positions and lengths count UTF-16 code units, as Excel does
//   (`LEN("😀")` is 2). UPPER/LOWER are locale-independent; PROPER
//   capitalises every letter that follows a non-letter (`PROPER("it's")` is
//   "It'S"); TRIM removes U+0020 only (and collapses its runs).
// - FIND is case-sensitive without wildcards; SEARCH is case-insensitive with
//   wildcards (criteria.ts); not found is #VALUE!.
// - CONCATENATE takes value arguments; CONCAT and TEXTJOIN read ranges
//   row-major.
// - TEXT formats through SPEC-118's formatValue: the code is the invariant
//   (en-US) code a cell's number format is, the output is localised, colour
//   ignored, and an invalid code is #VALUE!. VALUE reads text through
//   valueFromText (a number passes through, an empty cell is 0, "" #VALUE!).
import type { CellError } from "../../model";
import { FormatCodeError, parseFormatCode } from "../../numberFormat/parse";
import { formatValue, plainText } from "../../numberFormat/render";
import { forEachScalar, materialize } from "../areas";
import type { EvalContext } from "../evaluate";
import type { ArgList, FunctionSpec } from "../registry";
import {
  isEmpty,
  isError,
  isRange,
  MAX_TEXT_LENGTH,
  toLogical,
  toText,
  valueFromText,
  VALUE_ERROR,
  type EvalValue,
  type ScalarValue,
} from "../values";
import { arg, integerArg, numberOr, optional, repeating, scalarArg, textArg } from "./common";
import { searchText } from "./criteria";

/** Appends the texts of a value, read row-major, to `out`; `keepEmpty`
 * gives every empty position of a range or array a "" too. Returns the first
 * error, or #VALUE! once the joined text could not fit. */
function readTexts(
  ctx: EvalContext,
  value: EvalValue,
  keepEmpty: boolean,
  out: string[],
): CellError | undefined {
  let failure: CellError | undefined;
  const take = (scalar: ScalarValue): boolean => {
    const text = toText(scalar);
    if (isError(text)) {
      failure = text;
      return false;
    }
    out.push(text);
    if (out.length > MAX_TEXT_LENGTH + 1) {
      failure = VALUE_ERROR;
      return false;
    }
    return true;
  };
  if (keepEmpty && isRange(value)) {
    for (const area of value.areas) {
      const array = materialize(ctx, { kind: "range", areas: [area] });
      if (isError(array)) {
        return array;
      }
      if (!array.data.every(take)) {
        return failure;
      }
    }
    return undefined;
  }
  forEachScalar(ctx, value, (scalar, direct) => (direct || keepEmpty || !isEmpty(scalar) ? take(scalar) : true));
  return failure;
}

/** A joined text, #VALUE! past MAX_TEXT_LENGTH. */
function joined(parts: readonly string[], delimiter = ""): string | CellError {
  const text = parts.join(delimiter);
  return text.length > MAX_TEXT_LENGTH ? VALUE_ERROR : text;
}

/** One text in, one value out. */
function textual(name: string, summary: string, description: string, f: (text: string) => ScalarValue): FunctionSpec {
  return {
    name,
    category: "Text",
    minArgs: 1,
    maxArgs: 1,
    argKinds: ["value"],
    summary,
    args: [arg("text", description)],
    impl(args) {
      const text = textArg(args, 0);
      return isError(text) ? text : f(text);
    },
  };
}

const LETTER = /\p{L}/u;

/** Every letter after a non-letter upper-cased, every other letter lower. */
function proper(text: string): string {
  let out = "";
  let afterLetter = false;
  for (const ch of text) {
    const letter = LETTER.test(ch);
    out += letter ? (afterLetter ? ch.toLowerCase() : ch.toUpperCase()) : ch;
    afterLetter = letter;
  }
  return out;
}

/** U+0020 trimmed from both ends and its inner runs collapsed to one. */
function trimSpaces(text: string): string {
  let out = "";
  let pendingSpace = false;
  for (const ch of text) {
    if (ch === " ") {
      pendingSpace = out.length > 0;
      continue;
    }
    if (pendingSpace) {
      out += " ";
      pendingSpace = false;
    }
    out += ch;
  }
  return out;
}

/** `find_text`, `within_text` and `start_num` of FIND/SEARCH, validated:
 * a start before 1 or past the text is #VALUE!. */
function findArguments(args: ArgList, ctx: EvalContext): { find: string; within: string; start: number } | CellError {
  const find = textArg(args, 0);
  if (isError(find)) {
    return find;
  }
  const within = textArg(args, 1);
  if (isError(within)) {
    return within;
  }
  const start = numberOr(args, 2, ctx, 1);
  if (isError(start)) {
    return start;
  }
  const position = Math.trunc(start);
  if (position < 1 || position > Math.max(within.length, 1)) {
    return VALUE_ERROR;
  }
  return { find, within, start: position - 1 };
}

/** The first, last or middle characters of a text. */
function slicing(name: "LEFT" | "RIGHT", summary: string): FunctionSpec {
  return {
    name,
    category: "Text",
    minArgs: 1,
    maxArgs: 2,
    argKinds: ["value", "value"],
    summary,
    args: [arg("text", "The text."), optional("num_chars", "How many characters (default 1).")],
    impl(args, ctx) {
      const text = textArg(args, 0);
      if (isError(text)) {
        return text;
      }
      const count = numberOr(args, 1, ctx, 1);
      if (isError(count)) {
        return count;
      }
      const n = Math.trunc(count);
      if (n < 0) {
        return VALUE_ERROR;
      }
      return name === "LEFT" ? text.slice(0, n) : text.slice(Math.max(0, text.length - n));
    },
  };
}

export const TEXT_FUNCTIONS: FunctionSpec[] = [
  {
    name: "CONCAT",
    category: "Text",
    minArgs: 1,
    maxArgs: 254,
    argKinds: ["any"],
    repeat: { from: 0, count: 1 },
    summary: "Joins texts, reading ranges row by row.",
    args: [repeating("text", "A text, or a range or array of texts.")],
    impl(args, ctx) {
      const parts: string[] = [];
      for (let i = 0; i < args.length; i += 1) {
        const failure = readTexts(ctx, args.get(i), false, parts);
        if (failure) {
          return failure;
        }
      }
      return joined(parts);
    },
  },
  {
    name: "CONCATENATE",
    category: "Text",
    minArgs: 1,
    maxArgs: 255,
    argKinds: ["value"],
    repeat: { from: 0, count: 1 },
    summary: "Joins texts.",
    args: [repeating("text", "A text to join.")],
    impl(args) {
      const parts: string[] = [];
      for (let i = 0; i < args.length; i += 1) {
        const text = textArg(args, i);
        if (isError(text)) {
          return text;
        }
        parts.push(text);
      }
      return joined(parts);
    },
  },
  {
    name: "TEXTJOIN",
    category: "Text",
    minArgs: 3,
    maxArgs: 254,
    argKinds: ["any", "value", "any"],
    repeat: { from: 2, count: 1 },
    summary: "Joins texts with a delimiter, optionally skipping empty ones.",
    args: [
      arg("delimiter", "The text placed between the items, or a range of them used in turn."),
      arg("ignore_empty", "TRUE skips empty items."),
      repeating("text", "A text, or a range or array of texts (up to 252)."),
    ],
    impl(args, ctx) {
      // A range or array of delimiters is used in turn (Excel's TEXTJOIN).
      const delimiters: string[] = [];
      const failed = readTexts(ctx, args.get(0), true, delimiters);
      if (failed) {
        return failed;
      }
      const ignoreEmpty = toLogical(scalarArg(args, 1));
      if (isError(ignoreEmpty)) {
        return ignoreEmpty;
      }
      const parts: string[] = [];
      const keepEmpty = !ignoreEmpty && delimiters.some((delimiter) => delimiter.length > 0);
      for (let i = 2; i < args.length; i += 1) {
        const failure = readTexts(ctx, args.get(i), keepEmpty, parts);
        if (failure) {
          return failure;
        }
      }
      const items = ignoreEmpty ? parts.filter((part) => part.length > 0) : parts;
      if (delimiters.length <= 1) {
        return joined(items, delimiters[0] ?? "");
      }
      return joined(items.flatMap((item, i) => (i === 0 ? [item] : [delimiters[(i - 1) % delimiters.length]!, item])));
    },
  },
  textual("LEN", "Returns the number of characters in a text.", "The text to measure.", (text) => text.length),
  slicing("LEFT", "Returns the first characters of a text."),
  slicing("RIGHT", "Returns the last characters of a text."),
  {
    name: "MID",
    category: "Text",
    minArgs: 3,
    maxArgs: 3,
    argKinds: ["value", "value", "value"],
    summary: "Returns characters from the middle of a text.",
    args: [arg("text", "The text."), arg("start_num", "The position of the first character, from 1."), arg("num_chars", "How many characters.")],
    impl(args, ctx) {
      const text = textArg(args, 0);
      if (isError(text)) {
        return text;
      }
      const start = integerArg(args, 1, ctx);
      if (isError(start)) {
        return start;
      }
      const count = integerArg(args, 2, ctx);
      if (isError(count)) {
        return count;
      }
      if (start < 1 || count < 0) {
        return VALUE_ERROR;
      }
      return text.slice(start - 1, start - 1 + count);
    },
  },
  textual("UPPER", "Converts a text to upper case.", "The text to convert.", (text) => text.toUpperCase()),
  textual("LOWER", "Converts a text to lower case.", "The text to convert.", (text) => text.toLowerCase()),
  textual("PROPER", "Capitalises the first letter of each word.", "The text to convert.", proper),
  textual("TRIM", "Removes extra spaces from a text.", "The text to trim.", trimSpaces),
  {
    name: "SUBSTITUTE",
    category: "Text",
    minArgs: 3,
    maxArgs: 4,
    argKinds: ["value", "value", "value", "value"],
    summary: "Replaces occurrences of a text with another text.",
    args: [
      arg("text", "The text to change."),
      arg("old_text", "The text to replace."),
      arg("new_text", "The text to put in its place."),
      optional("instance_num", "Which occurrence to replace (default every one)."),
    ],
    impl(args, ctx) {
      const texts: string[] = [];
      for (let i = 0; i < 3; i += 1) {
        const text = textArg(args, i);
        if (isError(text)) {
          return text;
        }
        texts.push(text);
      }
      const [text, oldText, newText] = texts as [string, string, string];
      if (args.length > 3) {
        const instance = integerArg(args, 3, ctx);
        if (isError(instance)) {
          return instance;
        }
        if (instance < 1) {
          return VALUE_ERROR;
        }
        if (oldText.length === 0) {
          return text;
        }
        let at = -1;
        for (let n = 0; n < instance; n += 1) {
          at = text.indexOf(oldText, at < 0 ? 0 : at + oldText.length);
          if (at < 0) {
            return text;
          }
        }
        return joined([text.slice(0, at), newText, text.slice(at + oldText.length)]);
      }
      return oldText.length === 0 ? text : joined(text.split(oldText), newText);
    },
  },
  {
    name: "REPLACE",
    category: "Text",
    minArgs: 4,
    maxArgs: 4,
    argKinds: ["value", "value", "value", "value"],
    summary: "Replaces characters at a position with another text.",
    args: [
      arg("old_text", "The text to change."),
      arg("start_num", "The position of the first character to replace, from 1."),
      arg("num_chars", "How many characters to replace."),
      arg("new_text", "The text to put in their place."),
    ],
    impl(args, ctx) {
      const text = textArg(args, 0);
      if (isError(text)) {
        return text;
      }
      const start = integerArg(args, 1, ctx);
      if (isError(start)) {
        return start;
      }
      const count = integerArg(args, 2, ctx);
      if (isError(count)) {
        return count;
      }
      const replacement = textArg(args, 3);
      if (isError(replacement)) {
        return replacement;
      }
      if (start < 1 || count < 0) {
        return VALUE_ERROR;
      }
      return joined([text.slice(0, start - 1), replacement, text.slice(start - 1 + count)]);
    },
  },
  {
    name: "FIND",
    category: "Text",
    minArgs: 2,
    maxArgs: 3,
    argKinds: ["value", "value", "value"],
    summary: "Finds a text within another, case-sensitively.",
    args: [
      arg("find_text", "The text to find."),
      arg("within_text", "The text to search."),
      optional("start_num", "The position to start at (default 1)."),
    ],
    impl(args, ctx) {
      const call = findArguments(args, ctx);
      if (isError(call)) {
        return call;
      }
      const at = call.within.indexOf(call.find, call.start);
      return at < 0 ? VALUE_ERROR : at + 1;
    },
  },
  {
    name: "SEARCH",
    category: "Text",
    minArgs: 2,
    maxArgs: 3,
    argKinds: ["value", "value", "value"],
    summary: "Finds a text within another, ignoring case and allowing wildcards.",
    args: [
      arg("find_text", "The text to find; ? and * are wildcards."),
      arg("within_text", "The text to search."),
      optional("start_num", "The position to start at (default 1)."),
    ],
    impl(args, ctx) {
      const call = findArguments(args, ctx);
      if (isError(call)) {
        return call;
      }
      const at = searchText(call.within, call.find, call.start, ctx.settings);
      return at < 0 ? VALUE_ERROR : at + 1;
    },
  },
  {
    name: "TEXT",
    category: "Text",
    minArgs: 2,
    maxArgs: 2,
    argKinds: ["value", "value"],
    summary: "Formats a value as text with a number format.",
    args: [arg("value", "The value to format."), arg("format_text", "The number format, as in en-US (\"#,##0.00\").")],
    impl(args, ctx) {
      const code = textArg(args, 1);
      if (isError(code)) {
        return code;
      }
      try {
        parseFormatCode(code);
      } catch (error) {
        if (error instanceof FormatCodeError) {
          return VALUE_ERROR;
        }
        throw error;
      }
      let value = scalarArg(args, 0);
      if (isEmpty(value)) {
        value = 0;
      } else if (typeof value === "string") {
        value = valueFromText(value, ctx.settings.locale, ctx.settings.dateSystem) ?? value;
      }
      if (isError(value)) {
        return value;
      }
      const formatted = formatValue(value, code, {
        locale: ctx.settings.locale,
        dateSystem: ctx.settings.dateSystem,
        dateQuirks: ctx.settings.dateQuirks,
      });
      return plainText(formatted);
    },
  },
  {
    name: "VALUE",
    category: "Text",
    minArgs: 1,
    maxArgs: 1,
    argKinds: ["value"],
    summary: "Converts a text that looks like a number, date or time to a number.",
    args: [arg("text", "The text to convert.")],
    impl(args, ctx) {
      const value = scalarArg(args, 0);
      if (typeof value === "number") {
        return value;
      }
      if (isEmpty(value)) {
        return 0;
      }
      if (typeof value !== "string") {
        return VALUE_ERROR;
      }
      return valueFromText(value, ctx.settings.locale, ctx.settings.dateSystem) ?? VALUE_ERROR;
    },
  },
  {
    name: "REPT",
    category: "Text",
    minArgs: 2,
    maxArgs: 2,
    argKinds: ["value", "value"],
    summary: "Repeats a text a number of times.",
    args: [arg("text", "The text to repeat."), arg("number_times", "How many times.")],
    impl(args, ctx) {
      const text = textArg(args, 0);
      if (isError(text)) {
        return text;
      }
      const times = integerArg(args, 1, ctx);
      if (isError(times)) {
        return times;
      }
      if (times < 0 || text.length * times > MAX_TEXT_LENGTH) {
        return VALUE_ERROR;
      }
      return text.repeat(times);
    },
  },
  {
    name: "EXACT",
    category: "Text",
    minArgs: 2,
    maxArgs: 2,
    argKinds: ["value", "value"],
    summary: "Returns TRUE when two texts are identical, case included.",
    args: [arg("text1", "The first text."), arg("text2", "The second text.")],
    impl(args) {
      const a = textArg(args, 0);
      if (isError(a)) {
        return a;
      }
      const b = textArg(args, 1);
      return isError(b) ? b : a === b;
    },
  },
];
