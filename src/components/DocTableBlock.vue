<script setup lang="ts">
import { markersForBlocks } from "../lib/lists";
import { asDocumentTheme, asTrackChangesDisplayMode, isMergeChipField } from "../lib/documentModel";
import { asListDefinitions } from "../lib/core/lists";
import { chipLabel } from "../lib/mergeChips";
import { cachedFieldText, isReadOnlyFieldTitle } from "../lib/fields";
import {
  displayRunFormatting,
  headingTag,
  paragraphStyle,
  runKey,
  runStyle,
  tableCellDisplayStyle,
  tableColumnWidth,
  tableRowStyle,
  tableStyle,
  trackChangeBlockHidden,
  trackChangeRunHidden,
  trackChangeRunStyle,
} from "../lib/renderFormatting";
import { asTable, visibleRowCells } from "../lib/tables";
import { asCellRange, isCellInRange } from "../lib/core/tables";

// One table block on the canvas (SPEC-022; formatting and merged cells
// SPEC-023). Cell paragraphs render exactly like top-level ones (same
// classes, styles and marker attributes) and carry the full cell address in
// data attributes — that is the contract lib/domCaret resolves carets
// through. Runtime prop declarations, not the defineProps<{...}>()
// generics — see AppShell.vue's comment on the same espree constraint.
const props = defineProps({
  /** The Table block (documentModel.Table). */
  block: { type: Object, required: true },
  section: { type: Number, required: true },
  blockIndex: { type: Number, required: true },
  /** The active rectangular cell selection (lib/tables.CellRange), if any —
   * drives the per-cell highlight (AC-15, AC-16). */
  cellRange: { type: Object, default: null },
  /** The model's multilevel list definitions (SPEC-032) — cell paragraphs
   * derive defined-list markers and indents from them. */
  listDefinitions: { type: Array, default: () => [] },
  /** The review pane's display mode (SPEC-039 AC-9) — drives whether cell
   * runs show inline tracked-change markup. */
  displayMode: { type: String, default: "all" },
  /** The document's theme (SPEC-077) — cell shading, paragraph paint and run
   * colours/fonts resolve theme references against it. */
  theme: { type: Object, default: null },
});

/** A cell's colspan/rowspan attribute value: undefined (the HTML default,
 * 1) below its own span field's default of 1. */
function spanAttr(span = 1) {
  return span > 1 ? span : undefined;
}
</script>

<template>
  <table
    class="doc-table"
    data-testid="doc-table"
    :class="{ 'cell-selecting': cellRange && cellRange.section === section && cellRange.block === blockIndex }"
    :style="tableStyle(asTable(block))"
  >
    <colgroup>
      <col
        v-for="(widthCm, colIndex) in block.columnWidthsCm"
        :key="colIndex"
        :style="{ width: tableColumnWidth(asTable(block), Number(colIndex)) }"
      >
    </colgroup>
    <tbody>
      <tr
        v-for="(row, rowIndex) in block.rows"
        v-show="!trackChangeBlockHidden(row.changes, asTrackChangesDisplayMode(displayMode))"
        :key="rowIndex"
        :style="tableRowStyle(row)"
      >
        <td
          v-for="entry in visibleRowCells(row)"
          :key="entry.colIndex"
          class="doc-table-cell"
          :class="{
            'cell-selected': isCellInRange(asCellRange(cellRange), section, blockIndex, Number(rowIndex), entry.colIndex),
          }"
          :style="tableCellDisplayStyle(entry.cell, isCellInRange(asCellRange(cellRange), section, blockIndex, Number(rowIndex), entry.colIndex), asDocumentTheme(theme))"
          :colspan="spanAttr(entry.cell.colSpan)"
          :rowspan="spanAttr(entry.cell.rowSpan)"
          :data-row="rowIndex"
          :data-col="entry.colIndex"
          :data-testid="`table-cell-${rowIndex}-${entry.colIndex}`"
        >
          <component
            :is="headingTag(paragraph.headingLevel)"
            v-for="(paragraph, cellBlockIndex) in entry.cell.blocks"
            :key="cellBlockIndex"
            class="doc-paragraph"
            :style="paragraphStyle(paragraph, asDocumentTheme(theme), asListDefinitions(listDefinitions), asTrackChangesDisplayMode(displayMode))"
            :data-section="section"
            :data-block="blockIndex"
            :data-row="rowIndex"
            :data-col="entry.colIndex"
            :data-cell-block="cellBlockIndex"
            :data-list-marker="markersForBlocks(entry.cell.blocks, asListDefinitions(listDefinitions))[cellBlockIndex] || undefined"
            :data-list-level="paragraph.list ? paragraph.list.level : undefined"
          >
            <template v-if="paragraph.runs.length > 0">
              <template
                v-for="(run, runIndex) in paragraph.runs"
                :key="runKey(run, runIndex)"
              >
                <!-- A mail-merge chip (SPEC-046 AC-13): the same paint-time
                     substitution usePaginationView applies to body runs —
                     the canonical {{...}} label rendered as one caret atom
                     standing for the one-character model placeholder — so a
                     chip in a cell reads and deletes exactly like the same
                     chip in the body. First in precedence, as in the body. -->
                <span
                  v-if="run.field && isMergeChipField(run.field)"
                  class="doc-merge-chip"
                  contenteditable="false"
                  data-testid="merge-chip"
                  :data-caret-atom="run.text"
                  :style="[runStyle(run.formatting, asDocumentTheme(theme)), trackChangeRunStyle(run, asTrackChangesDisplayMode(displayMode))]"
                >{{ chipLabel(run.field) }}</span>
                <!-- A SPEC-053 field in a table cell (AC-33): a caret atom
                     showing its saved value. A field is never recomputed on
                     open, so `cached` is exactly what the cell displays. -->
                <span
                  v-else-if="run.field"
                  :class="isReadOnlyFieldTitle(run.field) ? 'doc-field doc-field-readonly' : 'doc-field'"
                  contenteditable="false"
                  data-testid="field"
                  :data-caret-atom="run.text"
                  :title="isReadOnlyFieldTitle(run.field) || undefined"
                  :style="[runStyle(run.formatting, asDocumentTheme(theme)), trackChangeRunStyle(run, asTrackChangesDisplayMode(displayMode))]"
                >{{ cachedFieldText(run.field) }}</span>
                <!-- SPEC-039 AC-9: a run hidden by the current display mode
                     (a pending deletion in "No markup", a pending insertion
                     in "Original") renders no text, staying a caret atom
                     carrying its full model text — see
                     lib/renderFormatting.ts's trackChangeRunHidden. -->
                <span
                  v-else-if="trackChangeRunHidden(run, asTrackChangesDisplayMode(displayMode))"
                  contenteditable="false"
                  :data-caret-atom="run.text"
                  :style="[runStyle(run.formatting, asDocumentTheme(theme)), trackChangeRunStyle(run, asTrackChangesDisplayMode(displayMode))]"
                />
                <span
                  v-else
                  :style="[runStyle(displayRunFormatting(run, asTrackChangesDisplayMode(displayMode)), asDocumentTheme(theme)), trackChangeRunStyle(run, asTrackChangesDisplayMode(displayMode))]"
                >{{ run.text }}</span>
              </template>
            </template>
            <br v-else>
          </component>
        </td>
      </tr>
    </tbody>
  </table>
</template>

<style scoped>
.doc-table {
  table-layout: fixed;
  border-collapse: collapse;
  margin: 2px 0;
}

.doc-table-cell {
  border: 1px solid #b6b6b6;
  padding: 0.097cm;
  vertical-align: top;
  overflow-wrap: break-word;
}

/* While a cell-range drag is active the browser's own text selection is
   suppressed — the rectangular highlight is the selection's presentation
   (AC-15, AC-16). */
.doc-table.cell-selecting {
  user-select: none;
}

.doc-table-cell.cell-selected {
  background: rgb(66 133 244 / 0.22);
}

/* The same marker rule as the canvas's top-level paragraphs (SPEC-018) —
   scoped styles don't cross component boundaries, so it is restated for
   cell paragraphs. */
.doc-paragraph {
  margin: 0;
  font-weight: normal;
  font-size: inherit;
}

.doc-paragraph[data-list-marker]::before {
  content: attr(data-list-marker);
  display: inline-block;
  min-width: 0.635cm;
  margin-inline-start: -0.635cm;
  white-space: nowrap;
}

/* The mail-merge chip pill (SPEC-046) — restated from DocBlockEntries.vue
   for the same reason as .doc-paragraph above: scoped styles don't cross
   the component boundary, and cells render their own runs. */
.doc-merge-chip {
  background: #e8f0fe;
  border-radius: 3px;
  padding: 0 2px;
  color: #1a73e8;
}

/* A SPEC-053 field in a cell (AC-33), restated from DocBlockEntries.vue for
   the same scoped-style reason as the chip pill above. */
.doc-field {
  cursor: default;
  user-select: none;
}

.doc-field-readonly {
  cursor: help;
}
</style>
