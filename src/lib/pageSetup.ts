// Moved to src/lib/core/pageSetup.ts — paper sizes, orientation and margin
// validation are editing-surface geometry shared with any future product
// (a Slides surface uses the same matching/validation over its own size
// catalogue). This shim keeps existing import paths working.
export * from "./core/pageSetup";
