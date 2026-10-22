<script setup lang="ts">
// The table style gallery (SPEC-133 "Table styles"; the design bar's
// `sheet-table-styles`): Excel's Light, Medium and Dark sections, "None"
// first in Light as in Excel, then the 60 built-ins. Each tile previews its
// style on a sample table (a header row over four banded data rows) through
// the grid's own path — tableCellLook for every cell, cellCss/fillCss for its
// paint, resolveEdges + edgeElementCss for its lines — against the workbook
// theme, so a tile never drifts from what picking it paints. Emits `pick`
// with the style's name, or undefined for None. Runtime prop declarations
// per the espree constraint (see AppShell.vue).
import { computed } from "vue";
import { asDocumentTheme } from "../../lib/core/theme";
import { cellCss, edgeElementCss, fillCss, resolveEdges } from "../../lib/spreadsheet/cellRender";
import { buildGridMetrics } from "../../lib/spreadsheet/geometry";
import { buildMergeIndex } from "../../lib/spreadsheet/merges";
import { BUILT_IN_TABLE_STYLE_NAMES, sampleTable, tableCellLook } from "../../lib/spreadsheet/tableStyles";

const props = defineProps({
  /** The workbook's DocumentTheme the previews resolve against; null = the default theme. */
  theme: { type: Object, default: null },
  /** The table's current style name; absent = None. */
  current: { type: String, default: undefined },
});

const emit = defineEmits(["pick"]);

const theme = computed(() => asDocumentTheme(props.theme));

/** The sample every tile previews: A1:D5 in 16 × 12 px cells. */
const SAMPLE = { top: 0, left: 0, bottom: 4, right: 3 };
const SHEET = { ...sampleTable(undefined).sheet, defaultColWidthPx: 16, defaultRowHeightPt: 9 };
const METRICS = buildGridMetrics(SHEET);
const MERGES = buildMergeIndex([]);
/** Room around the sample for its outer lines (a double line is 3 px). */
const INSET = 2;
const ROWS = [0, 1, 2, 3, 4];
const COLS = [0, 1, 2, 3];

const SECTIONS = [
  { label: "Light", prefix: "TableStyleLight" },
  { label: "Medium", prefix: "TableStyleMedium" },
  { label: "Dark", prefix: "TableStyleDark" },
];

/** A tile's preview of style `name` ("" = None): each cell's paint and a
 * dash in its text colour (thicker when bold), and the resolved lines. */
function preview(name = "") {
  const table = sampleTable(name, SAMPLE.bottom + 1, SAMPLE.right + 1).table;
  const cells = ROWS.flatMap((row) =>
    COLS.map((col) => {
      const look = tableCellLook(SHEET, table, row, col, theme.value);
      const text = cellCss(look ?? {}, theme.value).text;
      return {
        key: `${row}:${col}`,
        row,
        col,
        box: {
          ...fillCss(look?.fill, theme.value),
          left: `${INSET + METRICS.colLeft(col)}px`,
          top: `${INSET + METRICS.rowTop(row)}px`,
          width: `${METRICS.colWidth(col)}px`,
          height: `${METRICS.rowHeight(row)}px`,
        },
        dash: { "background-color": text.color, height: text["font-weight"] === "bold" ? "2px" : "1px" },
      };
    }),
  );
  const lines = resolveEdges(SAMPLE, SHEET, (row, col) => tableCellLook(SHEET, table, row, col, theme.value) ?? {}, METRICS, MERGES, theme.value).map(
    (segment, index) => ({ key: `${segment.orientation}:${segment.x}:${segment.y}:${index}`, style: edgeElementCss(segment, -INSET, -INSET) }),
  );
  return { cells, lines };
}

const sections = computed(() =>
  SECTIONS.map((section) => ({
    label: section.label,
    tiles: [
      ...(section.label === "Light" ? [{ key: "none", name: undefined, label: "None", title: "None", ...preview("") }] : []),
      ...BUILT_IN_TABLE_STYLE_NAMES.filter((name) => name.startsWith(section.prefix)).map((name) => ({
        key: name,
        name,
        label: name,
        title: name.replace(/^TableStyle(\D+)(\d+)$/, "$1 $2"),
        ...preview(name),
      })),
    ],
  })),
);
</script>

<template>
  <div
    class="table-style-gallery"
    data-testid="sheet-table-style-gallery"
    role="listbox"
    aria-label="Table styles"
  >
    <div
      v-for="section in sections"
      :key="section.label"
      class="section"
      role="group"
      :aria-label="section.label"
    >
      <span
        class="section-label"
        aria-hidden="true"
      >{{ section.label }}</span>
      <div class="tiles">
        <button
          v-for="tile in section.tiles"
          :key="tile.key"
          type="button"
          role="option"
          class="tile"
          :data-testid="tile.name === undefined ? 'sheet-table-style-none' : `sheet-table-style-${tile.name}`"
          :aria-label="tile.label"
          :title="tile.title"
          :aria-selected="props.current === tile.name"
          @mousedown.prevent
          @click="emit('pick', tile.name)"
        >
          <span
            class="preview"
            aria-hidden="true"
          >
            <span
              v-for="cell in tile.cells"
              :key="cell.key"
              class="preview-cell"
              :data-row="cell.row"
              :data-col="cell.col"
              :style="cell.box"
            >
              <span
                class="preview-text"
                :style="cell.dash"
              />
            </span>
            <span
              v-for="line in tile.lines"
              :key="line.key"
              class="preview-line"
              :style="line.style"
            />
          </span>
        </button>
      </div>
    </div>
  </div>
</template>

<style scoped>
.table-style-gallery {
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
  padding: var(--space-2);
  max-height: min(28rem, 70vh);
  overflow-y: auto;
}

.section-label {
  display: block;
  margin-bottom: var(--space-1);
  font-size: var(--text-sm);
  font-weight: var(--weight-button);
  color: var(--text-secondary);
}

.tiles {
  display: grid;
  grid-template-columns: repeat(7, max-content);
  gap: var(--space-1);
}

.tile {
  padding: var(--space-1);
  line-height: 0;
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-subtle);
  border-radius: var(--radius-sm);
  cursor: pointer;
}

.tile:hover {
  outline: calc(2 * var(--border-width)) solid var(--accent);
}

.tile[aria-selected="true"] {
  border-color: var(--accent);
  box-shadow: inset 0 0 0 var(--border-width) var(--accent);
}

/* The sample is paper: white under its cells, whatever the app theme. */
.preview {
  position: relative;
  display: block;
  width: 68px;
  height: 64px;
  overflow: hidden;
  background: #ffffff;
}

.preview-cell {
  position: absolute;
  display: flex;
  align-items: center;
  justify-content: center;
}

.preview-text {
  width: 60%;
}

.preview-line {
  position: absolute;
}
</style>
