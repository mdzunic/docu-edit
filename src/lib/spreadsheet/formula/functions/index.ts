// The core library (SPEC-120 "The core functions"): exactly 100 functions in
// seven files, registered once when this module loads (evaluate.ts imports
// it). SPEC-127 registers the rest of the library through the same
// registerFunction.
import { registerFunction, type FunctionSpec } from "../registry";
import { AGGREGATE_FUNCTIONS } from "./aggregate";
import { DATE_TIME_FUNCTIONS } from "./dateTime";
import { INFORMATION_FUNCTIONS } from "./information";
import { LOGICAL_FUNCTIONS } from "./logical";
import { LOOKUP_FUNCTIONS } from "./lookup";
import { MATH_FUNCTIONS } from "./math";
import { TEXT_FUNCTIONS } from "./text";

/** Every core function, by file. */
export const CORE_FUNCTION_FILES: Readonly<Record<string, readonly FunctionSpec[]>> = {
  "math.ts": MATH_FUNCTIONS,
  "aggregate.ts": AGGREGATE_FUNCTIONS,
  "logical.ts": LOGICAL_FUNCTIONS,
  "information.ts": INFORMATION_FUNCTIONS,
  "text.ts": TEXT_FUNCTIONS,
  "dateTime.ts": DATE_TIME_FUNCTIONS,
  "lookup.ts": LOOKUP_FUNCTIONS,
};

for (const specs of Object.values(CORE_FUNCTION_FILES)) {
  for (const spec of specs) {
    registerFunction(spec);
  }
}
