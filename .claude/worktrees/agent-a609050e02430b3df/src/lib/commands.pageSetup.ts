// --- page setup (SPEC-027) ---------------------------------------------
import { orderedRange, type CommandResult, type ModelSelection } from "./commands.editing";
import { snapCm } from "./commands.paragraph";
import { type Margins, type PageSize } from "./core/pageSetup";
import { type MeasurementUnit } from "./core/units";
import { type DocumentModel, type PageBorders, type Section } from "./documentModel";
import { type ModelColor } from "./core/theme";

/** Drops a stale preserved `w:sectPr` (SPEC-027): any command that changes
 * the page size or margins must not let the writer re-emit the source
 * .docx's original section properties over top of the new values — the
 * same rule tables.ts's clearPreservedTable/clearPreservedCell apply to
 * `w:tblPr`/`w:tcPr`. Absent on documents without one (.odt sources, blank
 * documents), which is also the common case, so callers only pay for the
 * copy when there is something to clear. */
function clearPreservedSectPr(model: DocumentModel): DocumentModel {
  if (model.sourceXml?.docxSectPr === undefined) {
    return model;
  }
  const sourceXml = { ...model.sourceXml };
  delete sourceXml.docxSectPr;
  return { ...model, sourceXml };
}

/** A section rewritten with new page geometry, its preserved sectPr dropped
 * so the writer regenerates it from the new values (SPEC-028) — the section
 * -level sibling of clearPreservedSectPr. */
function withSectionGeometry(
  section: Section,
  geometry: { pageSize?: PageSize; margins?: Margins },
): Section {
  const updated: Section = { ...section, ...geometry };
  delete updated.docxSectPr;
  return updated;
}

/** Sets the page's margins from a ruler marker drag — the horizontal ruler's
 * left/right (SPEC-014 AC-23) and the vertical ruler's top/bottom (SPEC-064
 * AC-2a). When the caret's section carries its own margins (SPEC-028), the
 * drag edits that section; otherwise it edits the document defaults, exactly
 * as before sections could diverge. */
export function setPageMargins(
  model: DocumentModel,
  selection: ModelSelection,
  patch: { leftCm?: number; rightCm?: number; topCm?: number; bottomCm?: number },
): CommandResult {
  const { start } = orderedRange(model, selection);
  const section = model.sections[start.section];
  const margins = { ...(section.margins ?? model.defaultStyles.margins) };
  if (patch.leftCm !== undefined) {
    margins.leftCm = snapCm(Math.max(0, patch.leftCm));
  }
  if (patch.rightCm !== undefined) {
    margins.rightCm = snapCm(Math.max(0, patch.rightCm));
  }
  if (patch.topCm !== undefined) {
    margins.topCm = snapCm(Math.max(0, patch.topCm));
  }
  if (patch.bottomCm !== undefined) {
    margins.bottomCm = snapCm(Math.max(0, patch.bottomCm));
  }
  let next: DocumentModel;
  if (section.margins) {
    next = {
      ...model,
      sections: model.sections.map((entry, index) =>
        index === start.section ? withSectionGeometry(entry, { margins }) : entry,
      ),
    };
    if (start.section === model.sections.length - 1) {
      next = clearPreservedSectPr(next);
    }
  } else {
    next = clearPreservedSectPr({
      ...model,
      defaultStyles: { ...model.defaultStyles, margins },
    });
  }
  const range = orderedRange(next, selection);
  return { model: next, selection: { anchor: range.start, focus: range.end } };
}

/** Sets the caret section's header distance from the top edge and/or footer
 * distance from the bottom edge, from a vertical-ruler band-boundary drag
 * (SPEC-064 AC-3a/AC-3b). Writes the per-section field (SPEC-029) directly —
 * these distances are a display/layout preference on the section, not part of
 * the page margins, so the drag never disturbs the margin markers. Negative
 * distances are clamped to zero. */
export function setHeaderFooterDistance(
  model: DocumentModel,
  selection: ModelSelection,
  patch: { headerDistanceCm?: number; footerDistanceCm?: number },
): CommandResult {
  const target = orderedRange(model, selection).start.section;
  const next: DocumentModel = {
    ...model,
    sections: model.sections.map((section, index) => {
      if (index !== target) {
        return section;
      }
      const updated: Section = { ...section };
      if (patch.headerDistanceCm !== undefined) {
        updated.headerDistanceCm = snapCm(Math.max(0, patch.headerDistanceCm));
      }
      if (patch.footerDistanceCm !== undefined) {
        updated.footerDistanceCm = snapCm(Math.max(0, patch.footerDistanceCm));
      }
      return updated;
    }),
  };
  const range = orderedRange(next, selection);
  return { model: next, selection: { anchor: range.start, focus: range.end } };
}

/** What a Page setup apply targets (SPEC-028 AC-20/21): the caret's section
 * or the whole document. */
export type PageSetupScope = "thisSection" | "wholeDocument";

/** Typed pass-through for a scope coming from a `.vue` template's plain
 * string (see documentModel.asCharacterFormatKey for the espree constraint). */
export function asPageSetupScope(value: string): PageSetupScope {
  return value as PageSetupScope;
}

/** The Page setup dialog's OK (SPEC-027 AC-1..7, SPEC-028 AC-20/21): paper
 * size, orientation and the four margins are already resolved to a concrete
 * `PageSize`/`Margins` by the dialog (paper-size lookup and orientation
 * swap happen there — see lib/pageSetup.ts); this command just snaps them
 * onto the twip grid and commits. `applyTo` picks the target: the caret's
 * section (`"thisSection"`), or the document defaults (`"wholeDocument"`,
 * which also drops every section's own override so the whole document
 * genuinely follows the new values). `measurementUnit`, when given,
 * replaces the document's preference (AC-7). */
export function setPageSetup(
  model: DocumentModel,
  selection: ModelSelection,
  patch: { pageSize: PageSize; margins: Margins; measurementUnit?: MeasurementUnit },
  applyTo: PageSetupScope = "wholeDocument",
): CommandResult {
  const pageSize: PageSize = {
    widthCm: snapCm(patch.pageSize.widthCm),
    heightCm: snapCm(patch.pageSize.heightCm),
  };
  const margins: Margins = {
    topCm: snapCm(patch.margins.topCm),
    rightCm: snapCm(patch.margins.rightCm),
    bottomCm: snapCm(patch.margins.bottomCm),
    leftCm: snapCm(patch.margins.leftCm),
  };
  let next: DocumentModel;
  if (applyTo === "thisSection") {
    const target = orderedRange(model, selection).start.section;
    next = {
      ...model,
      sections: model.sections.map((section, index) =>
        index === target ? withSectionGeometry(section, { pageSize, margins }) : section,
      ),
      measurementUnit: patch.measurementUnit ?? model.measurementUnit,
    };
    if (target === model.sections.length - 1) {
      next = clearPreservedSectPr(next);
    }
  } else {
    next = clearPreservedSectPr({
      ...model,
      defaultStyles: { ...model.defaultStyles, pageSize, margins },
      sections: model.sections.map((section) => {
        if (
          section.pageSize === undefined &&
          section.margins === undefined &&
          section.docxSectPr === undefined
        ) {
          return section;
        }
        const updated = { ...section };
        delete updated.pageSize;
        delete updated.margins;
        delete updated.docxSectPr;
        return updated;
      }),
      measurementUnit: patch.measurementUnit ?? model.measurementUnit,
    });
  }
  const { start, end } = orderedRange(next, selection);
  return { model: next, selection: { anchor: start, focus: end } };
}

// --- page borders and background (SPEC-074 AC-19..23) ------------------------

/** The Page setup dialog's Borders tab (SPEC-074 AC-19..23). Both fields
 * follow the `null`-removes convention: `pageBorders: null` clears the
 * section's border, `pageBackground: null` returns it to paper white.
 *
 * `applyTo` picks the scope the same way setPageSetup does — the caret's
 * section, or every section. Unlike page size and margins, there are no
 * document-level defaults for these: "whole document" means writing the same
 * value onto every section, which is also what makes a `.docx` save lossless
 * (OOXML holds one page background for the whole document — see AC-32). */
export function setPageBorders(
  model: DocumentModel,
  selection: ModelSelection,
  patch: { pageBorders?: PageBorders | null; pageBackground?: ModelColor | null },
  applyTo: PageSetupScope = "thisSection",
): CommandResult {
  const target = orderedRange(model, selection).start.section;
  const apply = (section: Section): Section => {
    const updated: Section = { ...section };
    if (patch.pageBorders !== undefined) {
      const borders = patch.pageBorders === null ? undefined : snapPageBorders(patch.pageBorders);
      if (borders === undefined) {
        delete updated.pageBorders;
      } else {
        updated.pageBorders = borders;
      }
    }
    if (patch.pageBackground !== undefined) {
      if (patch.pageBackground === null) {
        delete updated.pageBackground;
      } else {
        updated.pageBackground = patch.pageBackground;
      }
    }
    // The preserved sectPr carries the source's own w:pgBorders, so it must
    // go the moment the modelled value changes — the same rule
    // withSectionGeometry applies to page size and margins.
    delete updated.docxSectPr;
    return updated;
  };
  let next: DocumentModel = {
    ...model,
    sections: model.sections.map((section, index) =>
      applyTo === "wholeDocument" || index === target ? apply(section) : section,
    ),
  };
  if (applyTo === "wholeDocument" || target === model.sections.length - 1) {
    next = clearPreservedSectPr(next);
  }
  const { start, end } = orderedRange(next, selection);
  return { model: next, selection: { anchor: start, focus: end } };
}

/** Border widths on the eighth-of-a-point grid `@w:sz` carries exactly, and
 * the offset on the whole-point grid `@w:space` does — so a Borders tab edit
 * survives a save + re-open unchanged. The offset is clamped to OOXML's own
 * 0..31pt range rather than silently written out of schema. */
function snapPageBorders(borders: PageBorders): PageBorders {
  const snapped: PageBorders = {};
  for (const edge of ["top", "bottom", "left", "right"] as const) {
    const side = borders[edge];
    if (side) {
      snapped[edge] = { ...side, widthPt: Math.round(side.widthPt * 8) / 8 };
    }
  }
  if (!snapped.top && !snapped.bottom && !snapped.left && !snapped.right) {
    return snapped;
  }
  if (borders.offsetPt !== undefined) {
    snapped.offsetPt = Math.min(31, Math.max(0, Math.round(borders.offsetPt)));
  }
  snapped.measureFrom = borders.measureFrom ?? "page";
  if (borders.appliesTo !== undefined && borders.appliesTo !== "allPages") {
    snapped.appliesTo = borders.appliesTo;
  }
  return snapped;
}
