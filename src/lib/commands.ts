// The editing command API (SPEC-011 and successors), split into one module
// per command family: every content mutation the editor performs is a
// command in one of the commands.<family>.ts siblings, each sitting next to
// its commands.<family>.spec.ts where one exists, and this barrel re-exports
// them so every importer (and every spec) keeps importing "./commands".
//
// The export surface is guarded by commands.undo.spec.ts: its completeness
// check requires every function exported here to be a registered undoable
// command or a documented query. That is why four modules below are
// re-exported by NAME rather than `export *` — they additionally export
// module-internal helpers to sibling command modules (commands.editing's
// `result`, commands.paragraph's snapCm/updateTouchedParagraphs,
// commands.multilevel's definition internals, commands.styles' re-bake
// layers), which must not leak into this surface.
//
// The modules' import graph is acyclic: commands.editing.ts (shared types,
// selection primitives, base text edits, caret queries) is the base layer;
// commands.paragraph.ts adds the touched-paragraph walk and the snap grid;
// commands.multilevel.ts and commands.styles.ts export list/style internals
// consumed by commands.styles, commands.customStyles and
// commands.clearFormatting above them.
export {
  blockCount,
  blockOrdinal,
  caretSelection,
  deleteBackward,
  deleteForward,
  deleteRange,
  documentEnd,
  documentStart,
  insertLineBreak,
  insertText,
  isCollapsed,
  noCommandResult,
  noParagraphCommand,
  noPendingFormat,
  noSelection,
  orderedRange,
  selectAll,
  splitBlock,
} from "./commands.editing";
export type { CommandResult, ModelSelection, ParagraphCommand } from "./commands.editing";
export * from "./commands.breaks";
export * from "./commands.formatting";
export {
  adjustIndent,
  applyParagraphFormat,
  emptyParagraphFormatting,
  emptyParagraphPaginationState,
  firstTouchedFormatting,
  hasSpaceBefore,
  INDENT_STEP_CM,
  paragraphPaginationState,
  setAlignment,
  SPACE_BEFORE_TOGGLE_PT,
  toggleSpaceBefore,
  touchedBlocks,
} from "./commands.paragraph";
export type {
  BlockAddress,
  ParagraphFormatPatch,
  ParagraphPaginationState,
} from "./commands.paragraph";
export * from "./commands.pageSetup";
export * from "./commands.headers";
export * from "./commands.notes";
export * from "./commands.merge";
export * from "./commands.shapes";
export * from "./commands.lists";
export {
  applyListDefinition,
  applyMultilevelPreset,
  continueListNumbering,
  linkHeadingNumbering,
  normalizeDefinitionLevels,
  restartListNumbering,
  selectionListDefinitionLevels,
  selectionTabStops,
  setListNumberingValue,
  setTabStops,
} from "./commands.multilevel";
export {
  applyParagraphStyle,
  emptyToolbarStyleState,
  resetParagraphStyle,
  toolbarStyleState,
  updateStyleFromSelection,
} from "./commands.styles";
export type { ToolbarStyleState } from "./commands.styles";
export * from "./commands.customStyles";
export * from "./commands.clearFormatting";
