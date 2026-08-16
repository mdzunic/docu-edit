// --- headers, footers and page numbers (SPEC-029) ---------------------------
import { type CommandResult, type ModelSelection } from "./commands.editing";
import {
  createEmptyHeaderFooterContent,
  createFieldRun,
  sectionHeaderFooter,
  type DocumentModel,
  type HeaderFooterKind,
  type HeaderFooterPart,
  type PageNumberFormat,
  type Paragraph,
  type ParagraphAlignment,
  type Run,
  type Section,
} from "./documentModel";

/** Sets a section's header/footer options (SPEC-029 AC-11, AC-12, AC-22,
 * AC-23): `titlePage` is "Different first page", `evenAndOddHeaders` is
 * "Different odd and even pages", and `pageNumberStart` is "Start numbering
 * at" — `null` clears it back to "continue from the previous section".
 * `undefined` on any field leaves it unchanged, so a dialog can submit only
 * the fields it actually shows. The current selection passes through
 * untouched: these are section-wide settings, not a text edit. */
export function setHeaderFooterOptions(
  model: DocumentModel,
  selection: ModelSelection,
  sectionIndex: number,
  patch: { titlePage?: boolean; evenAndOddHeaders?: boolean; pageNumberStart?: number | null },
): CommandResult {
  const sections = model.sections.map((section, index) => {
    if (index !== sectionIndex) {
      return section;
    }
    const next: Section = { ...section };
    if (patch.titlePage !== undefined) {
      if (patch.titlePage) {
        next.titlePage = true;
      } else {
        delete next.titlePage;
      }
    }
    if (patch.evenAndOddHeaders !== undefined) {
      if (patch.evenAndOddHeaders) {
        next.evenAndOddHeaders = true;
      } else {
        delete next.evenAndOddHeaders;
      }
    }
    if (patch.pageNumberStart !== undefined) {
      if (patch.pageNumberStart === null) {
        delete next.pageNumberStart;
      } else {
        next.pageNumberStart = patch.pageNumberStart;
      }
    }
    delete next.docxSectPr;
    return next;
  });
  return { model: { ...model, sections }, selection };
}

/** Addresses a section's header/footer content as a throwaway single-section
 * document (section 0 holding exactly the header/footer's own blocks) so
 * every existing body-editing primitive (insertText, deleteBackward,
 * toggleFormat, table/image insertion, …) works unmodified against
 * header/footer content — SPEC-029 reuses them rather than duplicating a
 * second copy of every editing command. Absent content starts from one
 * empty paragraph, the same seed a blank document's own section starts
 * from. */
export function withHeaderFooterModel(
  model: DocumentModel,
  sectionIndex: number,
  part: HeaderFooterPart,
  kind: HeaderFooterKind,
): DocumentModel {
  const content = sectionHeaderFooter(model, sectionIndex, part, kind) ?? createEmptyHeaderFooterContent();
  return { ...model, sections: [{ blocks: content.blocks }] };
}

/** Folds a `withHeaderFooterModel` command result back into the real
 * document: the synthetic section's blocks become the section's
 * header/footer content for `kind`, while every other model-level change
 * the inner command made (new `media` entries from an inserted image,
 * updated `styles`, …) is kept from `result`. */
export function commitHeaderFooterModel(
  before: DocumentModel,
  sectionIndex: number,
  part: HeaderFooterPart,
  kind: HeaderFooterKind,
  result: DocumentModel,
): DocumentModel {
  const blocks = result.sections[0].blocks;
  const key: "headers" | "footers" = part === "header" ? "headers" : "footers";
  const sections = before.sections.map((section, index) => {
    if (index !== sectionIndex) {
      return section;
    }
    return { ...section, [key]: { ...section[key], [kind]: { blocks } } };
  });
  return { ...result, sections };
}

/** The header/footer kinds a section actually renders (SPEC-029 AC-9,
 * AC-20): `"default"` always, plus `"first"`/`"even"` when the section's
 * own toggles are on — the "Insert page number" control writes into every
 * one of them so the field shows up on every page (AC-20), not only the
 * ones already edited. */
function applicableHeaderFooterKinds(section: Section): HeaderFooterKind[] {
  const kinds: HeaderFooterKind[] = ["default"];
  if (section.titlePage) {
    kinds.push("first");
  }
  if (section.evenAndOddHeaders) {
    kinds.push("even");
  }
  return kinds;
}

/** The "Insert page number" control's settings (SPEC-029 AC-14..19):
 * `part`/`position` place it (header or footer, left/centre/right — the
 * paragraph's alignment); `format` is its numbering style; `pageXofY`
 * inserts "Page X of Y" (a PAGE field, literal " of ", a NUMPAGES field)
 * instead of a bare page number. */
export interface PageNumberInsertOptions {
  part: HeaderFooterPart;
  position: ParagraphAlignment;
  format: PageNumberFormat;
  pageXofY: boolean;
}

function pageNumberRuns(options: PageNumberInsertOptions): Run[] {
  if (options.pageXofY) {
    return [
      createFieldRun("page", options.format),
      { text: " of ", formatting: {} },
      createFieldRun("numpages", options.format),
    ];
  }
  return [createFieldRun("page", options.format)];
}

/** Inserts a page-number paragraph into a section's header or footer
 * (SPEC-029 AC-14..19): appended after any existing content so a repeated
 * insert never clobbers text already there. Creates the header/footer (and
 * every kind the section's own first-page/odd-even toggles call for) on
 * demand — AC's "does the control create one automatically" ambiguity
 * resolved in favour of "yes", the friendlier default. */
export function insertPageNumber(
  model: DocumentModel,
  sectionIndex: number,
  options: PageNumberInsertOptions,
): DocumentModel {
  const section = model.sections[sectionIndex];
  if (!section) {
    return model;
  }
  const paragraph: Paragraph = {
    type: "paragraph",
    runs: pageNumberRuns(options),
    formatting: { alignment: options.position },
  };
  const key: "headers" | "footers" = options.part === "header" ? "headers" : "footers";
  const updated = { ...section[key] };
  for (const kind of applicableHeaderFooterKinds(section)) {
    const existing = updated[kind];
    updated[kind] = existing ? { blocks: [...existing.blocks, paragraph] } : { blocks: [paragraph] };
  }
  const sections = model.sections.map((entry, index) =>
    index === sectionIndex ? { ...entry, [key]: updated } : entry,
  );
  return { ...model, sections };
}

/** `insertPageNumber` wrapped as a CommandResult (selection passes through
 * unchanged — inserting a page number never moves the body caret) so it
 * plugs into DocumentView's applyEdit the same way every other
 * dialog-mediated command does. */
export function insertPageNumberCommand(
  model: DocumentModel,
  selection: ModelSelection,
  sectionIndex: number,
  options: PageNumberInsertOptions,
): CommandResult {
  return { model: insertPageNumber(model, sectionIndex, options), selection };
}
