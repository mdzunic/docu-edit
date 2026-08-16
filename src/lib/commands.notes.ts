// --- footnotes and endnotes (SPEC-031) --------------------------------------
import {
  isCollapsed,
  orderedRange,
  result,
  type CommandResult,
  type ModelSelection,
} from "./commands.editing";
import {
  cloneContainers,
  createEmptyNoteContent,
  createNoteRun,
  type Block,
  type DocumentModel,
  type NoteKind,
  type NoteNumberFormat,
  type NoteRestart,
  type NoteSettings,
} from "./documentModel";
import { createNoteId } from "./notes";
import {
  deleteRange as deleteRangeInPlace,
  formattingAtOffset,
  insertRun as insertRunInPlace,
  paragraphAt,
} from "./textEditing";

/** Inserts a footnote/endnote reference mark at the selection (SPEC-031
 * AC-1, AC-3): a superscript marker run at the caret plus an empty note
 * content for it in `model.notes`. A non-collapsed selection is deleted
 * first (the typing-over rule); an atom caret (table/image/break edge) is
 * inert. The returned selection sits after the mark; the view then moves
 * the caret into the note's own area (AC-2, AC-4). Displayed numbers are
 * derived from mark order (lib/notes.ts), so every other mark renumbers by
 * construction (AC-8). Pass `noteId` to control the id (the readers do);
 * otherwise one is generated. */
export function insertNote(
  model: DocumentModel,
  selection: ModelSelection,
  kind: NoteKind,
  noteId: string = createNoteId(),
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  const base = isCollapsed(selection) ? start : deleteRangeInPlace(next, start, end);
  const paragraph = paragraphAt(next, base);
  if (!paragraph) {
    return result(next, base);
  }
  // The mark inherits the surrounding text's formatting (like insertText),
  // minus any link — a mark at a link's edge must not join the link.
  const formatting = { ...formattingAtOffset(paragraph, base.offset) };
  delete formatting.linkHref;
  delete formatting.linkAnchor;
  const caret = insertRunInPlace(next, base, createNoteRun(noteId, kind, formatting));
  if (!next.notes?.some((note) => note.noteId === noteId)) {
    next.notes = [...(next.notes ?? []), createEmptyNoteContent(noteId, kind)];
  }
  return result(next, caret);
}

/** Sets a note stream's numbering settings (SPEC-031 AC-17..20).
 * `undefined` leaves a field unchanged, so the dialog submits only what it
 * shows; the selection passes through — this is a document-wide setting,
 * not a text edit. */
export function setNoteSettings(
  model: DocumentModel,
  selection: ModelSelection,
  kind: NoteKind,
  patch: { numberFormat?: NoteNumberFormat; startAt?: number; restart?: NoteRestart },
): CommandResult {
  const key = kind === "footnote" ? "footnoteSettings" : "endnoteSettings";
  const settings: NoteSettings = { ...model[key] };
  if (patch.numberFormat !== undefined) {
    settings.numberFormat = patch.numberFormat;
  }
  if (patch.startAt !== undefined) {
    settings.startAt = Math.max(1, Math.floor(patch.startAt));
  }
  if (patch.restart !== undefined) {
    if (patch.restart === "continuous") {
      delete settings.restart;
    } else {
      settings.restart = patch.restart;
    }
  }
  const next: DocumentModel = { ...model, [key]: settings };
  // The .docx writer keeps a preserved body-level sectPr verbatim; the new
  // settings must reach its w:footnotePr/w:endnotePr, so drop it and let
  // the writer regenerate — the same rule every page-setup command follows.
  if (next.sourceXml?.docxSectPr !== undefined) {
    const sourceXml = { ...next.sourceXml };
    delete sourceXml.docxSectPr;
    next.sourceXml = sourceXml;
  }
  return { model: next, selection };
}

/** Addresses one note's content as a throwaway single-section document —
 * the withHeaderFooterModel seam reused for notes (SPEC-031 AC-6), so
 * every body-editing primitive works unmodified against a note's blocks.
 * A missing or empty note starts from one empty paragraph. */
export function withNoteModel(model: DocumentModel, noteId: string): DocumentModel {
  const note = model.notes?.find((entry) => entry.noteId === noteId);
  const blocks: Block[] =
    note && note.blocks.length > 0 ? note.blocks : [{ type: "paragraph", runs: [] }];
  return { ...model, sections: [{ blocks }] };
}

/** Folds a `withNoteModel` command result back into the real document: the
 * synthetic section's blocks become the note's content, the real sections
 * come back from `before`, and every other model-level change the inner
 * command made (new media, updated styles, …) is kept from `result`. */
export function commitNoteModel(
  before: DocumentModel,
  noteId: string,
  result: DocumentModel,
): DocumentModel {
  const blocks = result.sections[0].blocks;
  const notes = (before.notes ?? []).map((note) =>
    note.noteId === noteId ? { ...note, blocks } : note,
  );
  return { ...result, sections: before.sections, notes };
}
