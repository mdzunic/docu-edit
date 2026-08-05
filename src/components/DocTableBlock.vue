<script setup lang="ts">
import { markersForBlocks } from "../lib/lists";
import { headingTag, paragraphStyle, runKey, runStyle } from "../lib/renderFormatting";
import { asCellRange, isCellInRange } from "../lib/tables";

// One table block on the canvas (SPEC-022). Cell paragraphs render exactly
// like top-level ones (same classes, styles and marker attributes) and
// carry the full cell address in data attributes — that is the contract
// lib/domCaret resolves carets through. Runtime prop declarations, not the
// defineProps<{...}>() generics — see AppShell.vue's comment on the same
// espree constraint.
const props = defineProps({
  /** The Table block (documentModel.Table). */
  block: { type: Object, required: true },
  section: { type: Number, required: true },
  blockIndex: { type: Number, required: true },
  /** The active rectangular cell selection (lib/tables.CellRange), if any —
   * drives the per-cell highlight (AC-15, AC-16). */
  cellRange: { type: Object, default: null },
});

/** The table's overall width: the column widths' sum. */
function tableWidthCm() {
  let total = 0;
  for (const width of props.block.columnWidthsCm) {
    total += width;
  }
  return total;
}

/** Derived list markers for one cell's own block list (SPEC-018 semantics,
 * scoped to the cell — SPEC-022). */
function cellMarkers(cell = { blocks: [] }) {
  return markersForBlocks(cell.blocks);
}
</script>

<template>
  <table
    class="doc-table"
    data-testid="doc-table"
    :class="{ 'cell-selecting': cellRange && cellRange.section === section && cellRange.block === blockIndex }"
    :style="{ width: `${tableWidthCm()}cm` }"
  >
    <colgroup>
      <col
        v-for="(widthCm, colIndex) in block.columnWidthsCm"
        :key="colIndex"
        :style="{ width: `${widthCm}cm` }"
      >
    </colgroup>
    <tbody>
      <tr
        v-for="(row, rowIndex) in block.rows"
        :key="rowIndex"
      >
        <td
          v-for="(cell, colIndex) in row.cells"
          :key="colIndex"
          class="doc-table-cell"
          :class="{ 'cell-selected': isCellInRange(asCellRange(cellRange), section, blockIndex, rowIndex, colIndex) }"
          :data-row="rowIndex"
          :data-col="colIndex"
          :data-testid="`table-cell-${rowIndex}-${colIndex}`"
        >
          <component
            :is="headingTag(paragraph.headingLevel)"
            v-for="(paragraph, cellBlockIndex) in cell.blocks"
            :key="cellBlockIndex"
            class="doc-paragraph"
            :style="paragraphStyle(paragraph)"
            :data-section="section"
            :data-block="blockIndex"
            :data-row="rowIndex"
            :data-col="colIndex"
            :data-cell-block="cellBlockIndex"
            :data-list-marker="cellMarkers(cell)[cellBlockIndex] || undefined"
            :data-list-level="paragraph.list ? paragraph.list.level : undefined"
          >
            <template v-if="paragraph.runs.length > 0">
              <span
                v-for="(run, runIndex) in paragraph.runs"
                :key="runKey(run, runIndex)"
                :style="runStyle(run.formatting)"
              >{{ run.text }}</span>
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
</style>
