// Information (SPEC-120 "The core functions"): ISBLANK ISNUMBER ISTEXT
// ISERROR ISERR ISNA ISLOGICAL NA. The IS functions accept errors (an error
// argument is what ISERROR asks about, never a result to propagate); ISBLANK
// is TRUE only for an empty cell — a `""` result is not blank.
import type { FunctionSpec } from "../registry";
import { isEmpty, isError, NA_ERROR, type ScalarValue } from "../values";
import { arg, scalarArg } from "./common";

function predicate(name: string, summary: string, test: (value: ScalarValue) => boolean): FunctionSpec {
  return {
    name,
    category: "Information",
    minArgs: 1,
    maxArgs: 1,
    argKinds: ["value"],
    acceptsErrors: true,
    summary,
    args: [arg("value", "The value to test.")],
    impl(args) {
      return test(scalarArg(args, 0));
    },
  };
}

export const INFORMATION_FUNCTIONS: FunctionSpec[] = [
  predicate("ISBLANK", "Returns TRUE when the value is an empty cell.", (value) => isEmpty(value)),
  predicate("ISNUMBER", "Returns TRUE when the value is a number.", (value) => typeof value === "number"),
  predicate("ISTEXT", "Returns TRUE when the value is text.", (value) => typeof value === "string"),
  predicate("ISERROR", "Returns TRUE when the value is any error.", (value) => isError(value)),
  predicate("ISERR", "Returns TRUE when the value is an error other than #N/A.", (value) => isError(value) && value.error !== "#N/A"),
  predicate("ISNA", "Returns TRUE when the value is #N/A.", (value) => isError(value) && value.error === "#N/A"),
  predicate("ISLOGICAL", "Returns TRUE when the value is a logical value.", (value) => typeof value === "boolean"),
  {
    name: "NA",
    category: "Information",
    minArgs: 0,
    maxArgs: 0,
    argKinds: [],
    summary: "Returns the #N/A error.",
    args: [],
    impl() {
      return NA_ERROR;
    },
  },
];
