// The function library, registered once when this module loads (evaluate.ts
// imports it): SPEC-120's core functions — exactly its table, in the seven
// files it names (the spec's prose calls them "the 100 core functions"; its
// table lists 101, and the table is what is registered) — then SPEC-127's
// library of about two hundred more, by category, through the same
// registerFunction.
import { registerFunction, type FunctionSpec } from "../registry";
import { AGGREGATE_FUNCTIONS } from "./aggregate";
import { DATE_EXTRA_FUNCTIONS } from "./dateExtra";
import { DATE_TIME_FUNCTIONS } from "./dateTime";
import { ENGINEERING_FUNCTIONS } from "./engineering";
import { ERF_FUNCTIONS } from "./erf";
import { FINANCIAL_FUNCTIONS } from "./financial";
import { INFORMATION_FUNCTIONS } from "./information";
import { INFORMATION_EXTRA_FUNCTIONS } from "./informationExtra";
import { LOGICAL_FUNCTIONS } from "./logical";
import { LOOKUP_FUNCTIONS } from "./lookup";
import { LOOKUP_EXTRA_FUNCTIONS } from "./lookupExtra";
import { MATH_FUNCTIONS } from "./math";
import { MATH_EXTRA_FUNCTIONS } from "./mathExtra";
import { STATISTICAL_FUNCTIONS } from "./statistical";
import { SUBTOTAL_FUNCTIONS } from "./subtotal";
import { TEXT_FUNCTIONS } from "./text";
import { TEXT_EXTRA_FUNCTIONS } from "./textExtra";
import { TRIG_FUNCTIONS } from "./trig";

/** Every core function, by file (SPEC-120). */
export const CORE_FUNCTION_FILES: Readonly<Record<string, readonly FunctionSpec[]>> = {
  "math.ts": MATH_FUNCTIONS,
  "aggregate.ts": AGGREGATE_FUNCTIONS,
  "logical.ts": LOGICAL_FUNCTIONS,
  "information.ts": INFORMATION_FUNCTIONS,
  "text.ts": TEXT_FUNCTIONS,
  "dateTime.ts": DATE_TIME_FUNCTIONS,
  "lookup.ts": LOOKUP_FUNCTIONS,
};

/** SPEC-127's library, by file. */
export const LIBRARY_FUNCTION_FILES: Readonly<Record<string, readonly FunctionSpec[]>> = {
  "lookupExtra.ts": LOOKUP_EXTRA_FUNCTIONS,
  "textExtra.ts": TEXT_EXTRA_FUNCTIONS,
  "dateExtra.ts": DATE_EXTRA_FUNCTIONS,
  "mathExtra.ts": MATH_EXTRA_FUNCTIONS,
  "trig.ts": TRIG_FUNCTIONS,
  "subtotal.ts": SUBTOTAL_FUNCTIONS,
  "statistical.ts": STATISTICAL_FUNCTIONS,
  "financial.ts": FINANCIAL_FUNCTIONS,
  "informationExtra.ts": INFORMATION_EXTRA_FUNCTIONS,
  "engineering.ts": ENGINEERING_FUNCTIONS,
  "erf.ts": ERF_FUNCTIONS,
};

for (const specs of [...Object.values(CORE_FUNCTION_FILES), ...Object.values(LIBRARY_FUNCTION_FILES)]) {
  for (const spec of specs) {
    registerFunction(spec);
  }
}
