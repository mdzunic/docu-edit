// The single error type every validation failure in the generation API
// throws (SPEC-045 AC-7): its message always names the offending path
// (e.g. "body[3].cells: expected array") and the expected type/constraint,
// so a caller (human or system) can locate the bad field without a
// debugger. Never thrown for anything the writer itself would reject —
// this module is the gate that runs before any byte is written (AC-6).
export class DocumentGenerationError extends Error {}
