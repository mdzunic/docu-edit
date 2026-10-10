// The .ods half of SPEC-124's retained references (see
// ../retainedReferences.ts). Stub until the table below is implemented.
import type { SpreadsheetModel } from "../model";
import type { RetainedReferenceMapper } from "../retainedReferences";

/** Adjusts every retained ods reference through `mapper`; returns the ids
 * of the sheets whose own retained text changed. */
export function rewriteOdsRetainedReferences(model: SpreadsheetModel, mapper: RetainedReferenceMapper): Set<string> {
  void model;
  void mapper;
  return new Set();
}
