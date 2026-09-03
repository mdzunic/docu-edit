// Format → Theme… (SPEC-077 AC-9..AC-14): opens the theme gallery and
// applies the chosen (or customised) theme through the pure
// applyDocumentTheme command — ONE applyEdit call, so ONE undo entry
// ("Apply theme", AC-14). Cancelling never reaches applyEdit, so the model,
// history and dirty state stay untouched (AC-19); the dialog's previews and
// Customise draft are purely local to the component.
//
// DocumentView keeps only the `useThemeGallery(editorContext)` wiring call
// and the template hookup.
import { computed, ref } from "vue";
import { caretSelection, documentStart } from "../../lib/commands";
import { applyDocumentTheme } from "../../lib/commands.theme";
import type { DocumentTheme } from "../../lib/core/theme";
import { documentThemeOf, type DocumentModel } from "../../lib/documentModel";
import { blockText } from "../../lib/textEditing";
import type { EditorContext } from "./context";

/** One line of a gallery card's miniature (AC-11): a leading paragraph of
 * the document as plain text, flagged heading/body so the card styles it
 * with its own theme's fonts and colours. */
export interface ThemePreviewParagraph {
  text: string;
  heading: boolean;
}

/** The miniature stays cheap and deterministic: at most this many lines… */
const PREVIEW_PARAGRAPH_LIMIT = 6;
/** …of at most this many characters each. */
const PREVIEW_TEXT_LIMIT = 120;

/** The gallery card miniature's source (AC-11): the first non-empty
 * paragraphs of the first section — the document's first page, near enough,
 * without cloning the live editor DOM or waiting on layout. An empty result
 * (blank document) is the dialog's cue to fall back to its specimen sheet.
 * Exported for its unit test — pure, no Vue. */
export function themePreviewParagraphs(model: DocumentModel | null): ThemePreviewParagraph[] {
  if (!model) {
    return [];
  }
  const preview: ThemePreviewParagraph[] = [];
  for (const block of model.sections[0]?.blocks ?? []) {
    if (block.type !== "paragraph") {
      continue;
    }
    const text = blockText(block).trim();
    if (text === "") {
      continue;
    }
    preview.push({ text: text.slice(0, PREVIEW_TEXT_LIMIT), heading: block.headingLevel !== undefined });
    if (preview.length >= PREVIEW_PARAGRAPH_LIMIT) {
      break;
    }
  }
  return preview;
}

/** Typed pass-through for the preview prop arriving through a runtime-
 * declared Array (espree bars the annotation in `.vue` scripts — see
 * lib/pageRender.asRenderEntries for the pattern). */
export function asThemePreviewParagraphs(value: unknown): ThemePreviewParagraph[] {
  return (value ?? []) as ThemePreviewParagraph[];
}

export type ThemeGalleryDeps = EditorContext;

export function useThemeGallery(deps: ThemeGalleryDeps) {
  const { doc, editable, applyEdit } = deps;

  const open = ref(false);

  /** The gallery's first card (AC-10): the document's current theme. */
  const currentTheme = computed(() => documentThemeOf(doc.value));

  const previewParagraphs = computed(() => themePreviewParagraphs(doc.value));

  function openThemeGallery(): void {
    if (!doc.value || !editable.value) {
      return;
    }
    open.value = true;
  }

  function closeThemeGallery(): void {
    open.value = false;
  }

  /** A card click or the Customise pane's Apply. The selection is only
   * carried through so the caret survives the undo round-trip — the command
   * itself touches nothing but `model.theme`. */
  function applyTheme(theme: DocumentTheme): void {
    const model = doc.value;
    open.value = false;
    if (!model || !editable.value) {
      return;
    }
    const root = deps.contentRoot.value as HTMLElement | null;
    const selection = deps.currentSelection(root) ?? caretSelection(deps.getCaret() ?? documentStart());
    applyEdit(applyDocumentTheme(model, theme, selection), selection, "Apply theme");
  }

  return { open, currentTheme, previewParagraphs, openThemeGallery, closeThemeGallery, applyTheme };
}
