// Format → Theme… apply (SPEC-077 AC-14, AC-18): replacing the document
// theme is the smallest edit in the command family — the ONLY model change
// is `model.theme`. No colour/font reference anywhere in the document is
// rewritten and no literal is converted: every reference re-resolves against
// the new scheme through documentThemeOf at render time, which is the whole
// point of references. The blocks, sections and styles keep their object
// identity, so an undo/redo (one history step, D18) swings every resolved
// colour and font at once.
import { orderedRange, type CommandResult, type ModelSelection } from "./commands.editing";
import type { DocumentTheme } from "./core/theme";
import type { DocTheme, DocumentModel } from "./documentModel";

/** Applies `theme` (a gallery card or the Customise pane's draft) as the
 * document theme. The stored theme is a fresh clone — never an alias of a
 * BUILT_IN_THEMES entry or a live dialog draft — marked `dirty` so the
 * writers regenerate the theme part on save. When the document already has
 * a theme part (`partName`), it is kept so the part gets spliced in place;
 * a theme applied to a partless document (blank, .odt, .md) has none. */
export function applyDocumentTheme(
  model: DocumentModel,
  theme: DocumentTheme,
  selection: ModelSelection,
): CommandResult {
  const next: DocTheme = {
    name: theme.name,
    colors: { ...theme.colors },
    fonts: { ...theme.fonts },
    dirty: true,
  };
  if (model.theme?.partName !== undefined) {
    next.partName = model.theme.partName;
  }
  const { start, end } = orderedRange(model, selection);
  return { model: { ...model, theme: next }, selection: { anchor: start, focus: end } };
}
