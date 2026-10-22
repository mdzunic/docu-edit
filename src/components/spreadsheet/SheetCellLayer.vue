<script setup lang="ts">
// The sheet paper (SPEC-117 "Rendering", "Dark mode"): the cell layer of the
// grid, in layers bottom to top — gridlines (one line per visible row and
// column boundary, absent when the sheet hides them) → cells (fill and text,
// merged ranges as one element, overflowing text raised over its empty
// neighbours) → borders (one element per resolved edge). SPEC-059's rule
// applies: the page is not a surface, so its white, its #d4d4d4 gridlines
// and every cell colour (the file's own, from the model) are literal in
// every app theme, and the dim-page filter reaches it through `.sheet-paper`
// (App.vue). This file is the one tokens.spec.ts exempts.
//
// `bands` is useSheetCells' BandItem[] — the frozen rows, then the scrolled
// rows — each split into the frozen-column and scrolled-column regions. A
// cell's family is the resolved name; it is wrapped in SPEC-048's fontStack
// here. Each cell element's id, role, data-ref and ARIA come from
// cellAttributes — none on a merge's paint-only copy in a further frozen
// pane. A number format's `*x` fill (SPEC-118) is a flex row: the text
// before it, a run of the fill character taking the remaining width (clipped,
// hidden from assistive technology), the text after it. SPEC-120: a stale
// cell carries data-stale and its title (`staleTitle`). SPEC-128: each of
// the autofilter's header cells draws its `filter-button` (`filterButton`,
// useSortFilter's filterButtonAt) — right-aligned, 16 px, so it virtualises
// with the grid — whose press never reaches the grid's pointer handling and
// emits filter-button(col, anchor) with the viewport point under it. SPEC-129:
// a cell with a conditional-format result carries data-cf="true" (its style
// arrives already layered — dxf, number format, a scale's fill); its data
// bar (`sheet-cf-databar`, cfBarsOf) sits behind the text inside the cell,
// 2 px in, with a dashed axis line when it has one; its icon (CfIcon.vue,
// 16 px on the paper, so it scales with the zoom like everything on it)
// sits at the left edge, the text aligned in the width that remains; a
// winning bar or icon with showValue false hides the text. SPEC-130: the
// active cell is described by the data-validation input message while it
// shows (`describedBy`, cellAttributes). SPEC-135: a cell with a note draws
// `sheet-note-indicator` — a 6 px red triangle in its top-right corner — and
// one with a thread `sheet-thread-indicator`, a purple corner marker dimmed
// when the thread is resolved (`annotationMarks`, useCellNotes' marksAt);
// neither restyles the cell, and both are hidden from assistive technology
// (the active cell's aria-describedby carries the text). A linked cell
// (`linkAt`, useCellLinks' isLinked — a stored hyperlink or a top-level
// HYPERLINK() formula) carries data-link="true" and its text sits in a
// `.sheet-cell-link` span with the pointer cursor — the area a Ctrl/Cmd+click
// follows the link over; its colour and underline are the cell's own style,
// never restyled here. Runtime prop declarations per the espree constraint
// (see AppShell.vue).
import { computed } from "vue";
import { fontStack } from "../../lib/fonts/substitution";
import { annotationMarksOf, asBandItems, asDescribedCell, asStaleTitle, cellAttributes, cfBarsOf, filterButtonsOf } from "../../views/spreadsheet/sheetLayerProps";
import { isLinkedCell } from "../../views/spreadsheet/cellLinks";
import AppIcon from "../AppIcon.vue";
import CfIcon from "./CfIcon.vue";

const props = defineProps({
  /** The bands to draw (useSheetCells' BandItem[]). */
  bands: { type: Array, default: () => [] },
  /** Whether the sheet shows gridlines (`showGridLines`). */
  gridLines: { type: Boolean, default: true },
  /** The frozen columns' width, unscaled px (0 without a freeze). */
  frozenWidth: { type: Number, default: 0 },
  /** A stale cell's title by position (SPEC-120; useSheetGrid's). */
  staleTitle: { type: Function, default: undefined },
  /** The header row's filter buttons by position (SPEC-128). */
  filterButton: { type: Function, default: undefined },
  /** SPEC-130: { ref, id } — the cell the input message describes; null = none. */
  describedBy: { type: Object, default: undefined },
  /** SPEC-135: a cell's note and thread marks by position (useCellNotes' marksAt). */
  annotationMarks: { type: Function, default: undefined },
  /** SPEC-135: whether a cell is linked, by position (useCellLinks' isLinked). */
  linkAt: { type: Function, default: undefined },
});

const emit = defineEmits(["filter-button"]);

const items = computed(() => asBandItems(props.bands));
const staleTitleOf = computed(() => asStaleTitle(props.staleTitle));
const describedCell = computed(() => asDescribedCell(props.describedBy));

/** The text box's style, its family wrapped in the substitution stack. */
function textStyle(style = {}) {
  const family = Reflect.get(style, "font-family");
  return typeof family === "string" && family.length > 0 ? { ...style, "font-family": fontStack(family) } : style;
}

function px(value = 0) {
  return `${value}px`;
}

/** A filter button's press: its column and the point under it. */
function pressFilter(col = 0, event = new MouseEvent("click")) {
  const box = event.currentTarget instanceof Element ? event.currentTarget.getBoundingClientRect() : null;
  emit("filter-button", col, { left: box ? box.left : 0, top: box ? box.bottom : 0 });
}
</script>

<template>
  <div class="sheet-paper">
    <div
      v-for="band in items"
      :key="band.kind"
      class="sheet-band"
      :style="{ top: px(band.top), height: px(band.height) }"
    >
      <template v-if="props.gridLines">
        <div
          class="sheet-gridlines"
          :style="{ left: '0px', width: px(props.frozenWidth) }"
        >
          <div
            v-for="y in band.gridRows"
            :key="`h${y}`"
            class="sheet-gridline is-horizontal"
            :style="{ top: px(y) }"
          />
          <div
            v-for="x in band.gridColsFrozen"
            :key="`v${x}`"
            class="sheet-gridline is-vertical"
            :style="{ left: px(x) }"
          />
        </div>
        <div
          class="sheet-gridlines"
          :style="{ left: px(props.frozenWidth), right: '0px' }"
        >
          <div
            v-for="y in band.gridRows"
            :key="`h${y}`"
            class="sheet-gridline is-horizontal"
            :style="{ top: px(y) }"
          />
          <div
            v-for="x in band.gridColsScrolled"
            :key="`v${x}`"
            class="sheet-gridline is-vertical"
            :style="{ left: px(x) }"
          />
        </div>
      </template>
      <div
        v-for="row in band.rows"
        :key="row.index"
        class="grid-row"
        role="row"
        :aria-rowindex="row.index + 2"
        :style="{ top: px(row.top), height: px(row.height) }"
      >
        <div
          v-for="region in [{ key: 'frozen', cells: row.frozen, left: 0, width: props.frozenWidth }, { key: 'scrolled', cells: row.scrolled, left: props.frozenWidth, width: null }]"
          :key="region.key"
          class="sheet-region"
          :style="region.width === null ? { left: px(region.left), right: '0px' } : { left: px(region.left), width: px(region.width) }"
        >
          <div
            v-for="cell in region.cells"
            :key="cell.key"
            class="sheet-cell"
            :class="[cell.kindClass, { 'is-merged': cell.merged, 'is-raised': cell.raised, 'is-paint-only': cell.paintOnly }]"
            v-bind="cellAttributes(cell, staleTitleOf, describedCell)"
            :data-link="isLinkedCell(cell, props.linkAt) ? 'true' : undefined"
            :style="cell.boxStyle"
          >
            <div
              v-for="bar in cfBarsOf(cell)"
              :key="`cf-bar-${bar.widthPct}`"
              class="sheet-cf-bar-area"
              aria-hidden="true"
            >
              <div
                class="sheet-cf-databar"
                data-testid="sheet-cf-databar"
                aria-hidden="true"
                :data-width-pct="bar.widthPct"
                :style="bar.barStyle"
              />
              <div
                v-if="bar.axisStyle"
                class="sheet-cf-axis"
                :style="bar.axisStyle"
              />
            </div>
            <span
              v-if="cell.cf?.icon && !cell.paintOnly"
              class="sheet-cf-icon-box"
            >
              <CfIcon
                :set="cell.cf.icon.set"
                :index="cell.cf.icon.index"
              />
            </span>
            <div
              v-if="!cell.cf?.hideValue"
              class="sheet-cell-text"
              :style="textStyle(cell.textStyle)"
            >
              <span
                class="sheet-cell-content"
                :class="{ 'is-filled': cell.fill }"
                :style="cell.contentStyle"
              ><template v-if="cell.runs"><span
                v-for="(run, i) in cell.runs"
                :key="i"
                :class="{ 'sheet-cell-link': isLinkedCell(cell, props.linkAt) }"
                :style="textStyle(run.style)"
              >{{ run.text }}</span></template><template v-else-if="cell.fill"><span :class="{ 'sheet-cell-link': isLinkedCell(cell, props.linkAt) }">{{ cell.fill.before }}</span><span
                class="sheet-cell-fill"
                aria-hidden="true"
              >{{ cell.fill.run }}</span><span :class="{ 'sheet-cell-link': isLinkedCell(cell, props.linkAt) }">{{ cell.fill.after }}</span></template><span
                v-else-if="isLinkedCell(cell, props.linkAt)"
                class="sheet-cell-link"
              >{{ cell.text }}</span><template v-else>{{ cell.text }}</template></span>
            </div>
            <svg
              v-if="cell.diagonal"
              class="sheet-cell-diagonal"
              aria-hidden="true"
              :viewBox="`0 0 ${cell.width} ${cell.height}`"
              preserveAspectRatio="none"
            >
              <line
                v-if="cell.diagonal.down"
                x1="0"
                y1="0"
                :x2="cell.width"
                :y2="cell.height"
                :stroke="cell.diagonal.color"
                :stroke-width="cell.diagonal.width"
                :stroke-dasharray="cell.diagonal.dasharray || undefined"
                :stroke-opacity="cell.diagonal.opacity ?? undefined"
              />
              <line
                v-if="cell.diagonal.up"
                x1="0"
                :y1="cell.height"
                :x2="cell.width"
                y2="0"
                :stroke="cell.diagonal.color"
                :stroke-width="cell.diagonal.width"
                :stroke-dasharray="cell.diagonal.dasharray || undefined"
                :stroke-opacity="cell.diagonal.opacity ?? undefined"
              />
            </svg>
            <button
              v-for="button in filterButtonsOf(cell, props.filterButton)"
              :key="button.state"
              type="button"
              class="sheet-filter-button"
              data-testid="filter-button"
              aria-haspopup="dialog"
              :aria-expanded="button.expanded ? 'true' : 'false'"
              :aria-label="button.label"
              :title="button.label"
              :data-state="button.state"
              @pointerdown.stop
              @mousedown.stop
              @dblclick.stop
              @click.stop="pressFilter(cell.col, $event)"
            >
              <AppIcon :name="button.icon" />
            </button>
            <span
              v-for="mark in annotationMarksOf(cell, props.annotationMarks)"
              :key="mark.kind"
              :class="['sheet-annotation-mark', `is-${mark.kind}`, { 'is-resolved': mark.resolved }]"
              :data-testid="mark.testid"
              :data-resolved="mark.resolved ? 'true' : undefined"
              aria-hidden="true"
            />
          </div>
        </div>
      </div>
      <div
        class="sheet-borders"
        :style="{ left: '0px', width: px(props.frozenWidth) }"
      >
        <div
          v-for="border in band.bordersFrozen"
          :key="border.key"
          class="sheet-border"
          :style="border.style"
        />
      </div>
      <div
        class="sheet-borders"
        :style="{ left: px(props.frozenWidth), right: '0px' }"
      >
        <div
          v-for="border in band.bordersScrolled"
          :key="border.key"
          class="sheet-border"
          :style="border.style"
        />
      </div>
    </div>
  </div>
</template>

<style scoped>
/* The paper: literal white with #d4d4d4 gridlines in every app theme — the
   file's colours are painted on it, and App.vue's dim-page rule dims it as it
   dims a text page. */
.sheet-paper {
  position: absolute;
  overflow: hidden;
  background: #ffffff;
  color: #000000;
}

.sheet-band {
  position: absolute;
  left: 0;
  right: 0;
  overflow: hidden;
}

.sheet-gridlines,
.sheet-borders {
  position: absolute;
  top: 0;
  bottom: 0;
  overflow: hidden;
  pointer-events: none;
}

.sheet-borders {
  z-index: 2;
}

.sheet-gridline {
  position: absolute;
  background: #d4d4d4;
}

.sheet-gridline.is-horizontal {
  left: 0;
  right: 0;
  height: 1px;
}

.sheet-gridline.is-vertical {
  top: 0;
  bottom: 0;
  width: 1px;
}

.grid-row {
  position: absolute;
  left: 0;
  right: 0;
}

/* A column region of one row: clips horizontally at the freeze line, never
   vertically (a merged range runs on below its anchor row). */
.sheet-region {
  position: absolute;
  top: 0;
  height: 100%;
  clip-path: inset(-100000px 0 -100000px 0);
}

.sheet-cell {
  position: absolute;
  display: flex;
  flex-direction: column;
  justify-content: flex-end;
  box-sizing: border-box;
  font-family: Calibri, Carlito, sans-serif;
  font-size: 11pt;
  line-height: 1.2;
}

.sheet-cell.is-raised {
  z-index: 1;
}

.sheet-cell-text {
  position: absolute;
  top: 0;
  bottom: 0;
  display: flex;
  flex-direction: column;
  justify-content: inherit;
  box-sizing: border-box;
  overflow: hidden;
  white-space: pre;
}

.sheet-cell-content {
  display: block;
}

/* SPEC-135: a linked cell's text — the pointer cursor over the glyphs only. */
.sheet-cell-link {
  cursor: pointer;
}

/* A format's fill: the run takes whatever width the text leaves. */
.sheet-cell-content.is-filled {
  display: flex;
}

.sheet-cell-fill {
  flex: 1 1 0;
  min-width: 0;
  overflow: hidden;
}

/* SPEC-129: a data bar's area — inside the cell, 2 px in — behind the text
   (earlier in the cell's paint order), and its icon at the left edge. */
.sheet-cf-bar-area {
  position: absolute;
  inset: 2px;
  pointer-events: none;
}

.sheet-cf-databar {
  position: absolute;
  top: 0;
  bottom: 0;
  box-sizing: border-box;
}

.sheet-cf-axis {
  position: absolute;
  top: -2px;
  bottom: -2px;
  width: 0;
}

.sheet-cf-icon-box {
  position: absolute;
  left: 1px;
  top: 50%;
  transform: translateY(-50%);
  pointer-events: none;
}

.sheet-cell-diagonal {
  position: absolute;
  inset: 0;
  width: 100%;
  height: 100%;
  pointer-events: none;
}

/* SPEC-128: a header cell's filter dropdown, on the paper (literal colours,
   as every paper colour), right-aligned at the cell's foot. */
.sheet-filter-button {
  position: absolute;
  right: 1px;
  bottom: 1px;
  z-index: 1;
  display: flex;
  align-items: center;
  justify-content: center;
  box-sizing: border-box;
  width: 16px;
  height: 16px;
  padding: 0;
  color: #444444;
  background: #f3f3f3;
  border: 1px solid #a6a6a6;
  border-radius: 2px;
  cursor: pointer;
}

.sheet-filter-button:hover {
  background: #e1e1e1;
}

.sheet-filter-button[data-state="filtered"],
.sheet-filter-button[data-state="sorted-asc"],
.sheet-filter-button[data-state="sorted-desc"] {
  color: #1f5fbf;
}

.sheet-filter-button:focus-visible {
  outline: 2px solid #1f5fbf;
  outline-offset: 1px;
}

.sheet-filter-button :deep(.app-icon) {
  width: 14px;
  height: 14px;
}

/* SPEC-135: a note's 6 px red triangle in the top-right corner, and a
   thread's purple marker beside it (dimmed when resolved) — over the text,
   never taking the pointer. */
.sheet-annotation-mark {
  position: absolute;
  top: 0;
  right: 0;
  z-index: 1;
  width: 0;
  height: 0;
  pointer-events: none;
  border-style: solid;
  border-width: 0 6px 6px 0;
  border-color: transparent #ff0000 transparent transparent;
}

.sheet-annotation-mark.is-thread {
  border-width: 0 8px 8px 0;
  border-color: transparent #7030a0 transparent transparent;
}

.sheet-annotation-mark.is-note + .sheet-annotation-mark.is-thread {
  right: 7px;
}

.sheet-annotation-mark.is-resolved {
  opacity: 0.4;
}

.sheet-border {
  position: absolute;
  box-sizing: content-box;
  width: 0;
  height: 0;
}
</style>
