// Whether the sheet on show's annotations can be edited (SPEC-135, AC-11).
// An .xlsx sheet whose comments, VML or threaded part broke the annotation
// subset — or whose `<hyperlinks>` did — reads with none of them modelled,
// and the writer keeps those parts byte-verbatim; a note, comment or link
// added there could never reach the file. Rather than show it and lose it on
// save, the commands refuse it up front: `notesReason` (notes and comments)
// and `linksReason` (links) hold the disabled items' title, "" when editing
// applies. The check is xlsx/annotations.ts' xlsxAnnotationEditing, loaded
// on demand with an .xlsx workbook (the view never loads the xlsx code
// otherwise) and cached per retained sheet part on the model's `opaque`,
// which the retained-references pass replaces (never mutates) when it
// rewrites a part. Until it has loaded, editing applies.
import { computed, shallowRef, watch, type Ref, type ShallowRef } from "vue";
import type { OpaquePart } from "../../lib/core/packageParts";
import { LINKS_KEPT_AS_READ_TITLE, NOTES_KEPT_AS_READ_TITLE } from "../../lib/spreadsheet/annotations";
import type { Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";

type EditingCheck = (model: Pick<SpreadsheetModel, "opaque" | "persons">, sheet: Pick<Sheet, "partName">) => { notes: boolean; hyperlinks: boolean };

const EDITABLE = { notes: true, hyperlinks: true } as const;

export interface AnnotationEditingDeps {
  book: { workbook: ShallowRef<SpreadsheetModel | null>; activeSheet: Readonly<Ref<Sheet | null>> };
  /** The check — default: xlsxAnnotationEditing, imported on demand. */
  load?: () => Promise<EditingCheck>;
}

export function useAnnotationEditing(deps: AnnotationEditingDeps) {
  const load = deps.load ?? (async () => (await import("../../lib/spreadsheet/xlsx/annotations")).xlsxAnnotationEditing);
  const check = shallowRef<EditingCheck | null>(null);
  let loading = false;
  watch(
    () => deps.book.workbook.value?.format,
    (format) => {
      if (format === "xlsx" && check.value === null && !loading) {
        loading = true;
        void load().then((loaded) => {
          check.value = loaded;
        });
      }
    },
    { immediate: true },
  );

  const cache = new WeakMap<readonly OpaquePart[], Map<string, { notes: boolean; hyperlinks: boolean }>>();
  const editing = computed(() => {
    const m = deps.book.workbook.value;
    const sheet = deps.book.activeSheet.value;
    const run = check.value;
    if (!m || m.format !== "xlsx" || !run || sheet?.kind !== "grid" || sheet.partName === undefined) {
      return EDITABLE;
    }
    let byPart = cache.get(m.opaque);
    if (!byPart) {
      byPart = new Map();
      cache.set(m.opaque, byPart);
    }
    let result = byPart.get(sheet.partName);
    if (!result) {
      result = run(m, sheet);
      byPart.set(sheet.partName, result);
    }
    return result;
  });

  const notesReason = computed(() => (editing.value.notes ? "" : NOTES_KEPT_AS_READ_TITLE));
  const linksReason = computed(() => (editing.value.hyperlinks ? "" : LINKS_KEPT_AS_READ_TITLE));
  return { notesReason, linksReason };
}

export type AnnotationEditing = ReturnType<typeof useAnnotationEditing>;
