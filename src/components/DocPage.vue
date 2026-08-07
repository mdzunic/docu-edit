<script setup lang="ts">
// One rendered page of the paginated canvas (SPEC-028): a fixed-size page
// box whose content area renders the block entries the pagination layout
// assigned to it. A separate component so the entry v-for can carry v-memo
// (Vue only honours v-memo on the v-for element itself — nesting the entry
// loop inside DocumentView's page loop would silently disable it, and the
// memo is what keeps typing responsive on very large documents).
//
// Entries arrive fully precomputed (tag, style objects, run render lists,
// slice windows) from DocumentView's pagedViews computed, so this component
// holds no document logic of its own. Runtime prop/emit declarations, not
// the generic form — this script is parsed by espree (see AppShell.vue).
//
// Header/footer bands (SPEC-029): two more entry lists, rendered as
// absolutely-positioned bands inside the same page box rather than folded
// into the body's content flow — a header/footer never affects pagination
// (lib/pagination.ts is untouched by SPEC-029), it just repeats within the
// page's own top/bottom margin. `headerEditable`/`footerEditable` mark the
// one page currently being edited (only one at a time, see DocumentView's
// activeHeaderFooter) — every other page renders the identical content
// read-only, which is how "edits apply to every page in the section at
// once" (AC-10) falls out for free: they all read the same model content.
import { computed } from "vue";
import DocBlockEntries from "./DocBlockEntries.vue";
import DocPageBlock from "./DocPageBlock.vue";
import { asPageColumns, asRenderEntries, noRenderEntries } from "../lib/pageRender";

const props = defineProps({
  pageStyle: { type: Object, required: true },
  contentStyle: { type: Object, required: true },
  /** Precomputed render entries (lib/pageRender.RenderEntry) — the typed
   * default seed gives the template its element type (see StylePicker). */
  entries: { type: Array, default: () => noRenderEntries() },
  /** This page's multi-column geometry (SPEC-030), or null for a single
   * column. When set, entries are grouped into positioned column boxes. */
  columns: { type: Object, default: null },
  /** A virtualized page (SPEC-028): the box renders with its exact layout
   * geometry but none of its content — the pagination controller measures
   * its blocks off-DOM instead. */
  virtual: { type: Boolean, default: false },
  cellRange: { type: Object, default: null },
  headerEntries: { type: Array, default: () => noRenderEntries() },
  footerEntries: { type: Array, default: () => noRenderEntries() },
  headerStyle: { type: Object, default: null },
  footerStyle: { type: Object, default: null },
  headerEditable: { type: Boolean, default: false },
  footerEditable: { type: Boolean, default: false },
  /** Body content dims while a header/footer is being edited (AC-1, AC-2). */
  dimmed: { type: Boolean, default: false },
});

defineEmits([
  "image-select",
  "dblclick-header",
  "dblclick-footer",
  "dblclick-body",
  "hf-beforeinput",
  "hf-keydown",
  "hf-compositionstart",
  "hf-compositionend",
]);

const renderEntries = computed(() => asRenderEntries(props.entries));

// Multi-column layout (SPEC-030): entries carry the column they flow into, so
// each column box renders its own slice of the list. The single-column path
// below is left untouched — a page with no columns renders exactly as before.
const columnBoxes = computed(() => {
  const cols = asPageColumns(props.columns);
  if (!cols || cols.count <= 1) {
    return [];
  }
  return cols.columns.map((column, index) => ({
    index,
    style: { left: `${column.offsetCm}cm`, width: `${column.widthCm}cm` },
    entries: renderEntries.value.filter((entry) => (entry.column ?? 0) === index),
  }));
});

/** Left offset (cm) of each "line between columns" rule (AC-4): centred in the
 * gutter to the right of every column but the last. */
const columnRules = computed(() => {
  const cols = asPageColumns(props.columns);
  if (!cols || cols.count <= 1 || !cols.separator) {
    return [];
  }
  return cols.columns.slice(0, -1).map((column) => column.offsetCm + column.widthCm + cols.gapCm / 2);
});

const hasColumns = computed(() => columnBoxes.value.length > 0);
</script>

<template>
  <div
    class="page"
    data-testid="document-page"
    :data-virtual="virtual ? '1' : undefined"
    :style="pageStyle"
  >
    <div
      class="page-header-band"
      data-testid="page-header-band"
      :style="headerStyle"
      :contenteditable="headerEditable"
      :data-hf-editable="headerEditable ? '1' : undefined"
      spellcheck="false"
      @dblclick="$emit('dblclick-header')"
      @beforeinput.stop="headerEditable ? $emit('hf-beforeinput', $event) : undefined"
      @keydown.stop="headerEditable ? $emit('hf-keydown', $event) : undefined"
      @compositionstart.stop="headerEditable ? $emit('hf-compositionstart') : undefined"
      @compositionend.stop="headerEditable ? $emit('hf-compositionend', $event) : undefined"
    >
      <DocBlockEntries :entries="headerEntries" />
    </div>
    <div
      class="page-content"
      :style="contentStyle"
      @dblclick="$emit('dblclick-body')"
    >
      <div
        class="page-margin"
        data-testid="page-margin"
        :class="{ 'has-columns': hasColumns, 'page-margin-dimmed': dimmed }"
      >
        <!-- Multi-column pages (SPEC-030): each column is a positioned box the
             engine's per-column flow filled; "line between" rules sit in the
             gutters. -->
        <template v-if="hasColumns">
          <div
            v-for="col in columnBoxes"
            :key="col.index"
            class="doc-column"
            data-testid="doc-column"
            :data-column="col.index"
            :style="col.style"
          >
            <!-- v-memo cannot nest inside the column v-for, so multi-column
                 pages re-render their blocks each patch — acceptable, they are
                 far fewer than a long single-column document's pages. -->
            <DocPageBlock
              v-for="entry in col.entries"
              :key="entry.key"
              :entry="entry"
              :cell-range="cellRange"
              @image-select="(section, blockIndex) => $emit('image-select', section, blockIndex)"
            />
          </div>
          <div
            v-for="(left, ruleIndex) in columnRules"
            :key="`rule-${ruleIndex}`"
            class="doc-column-rule"
            aria-hidden="true"
            :style="{ left: `${left}cm` }"
          />
        </template>
        <!-- Single-column page: unchanged from SPEC-028. -->
        <template v-else>
          <DocBlockEntries
            :entries="entries"
            :cell-range="cellRange"
            @image-select="(section, block) => $emit('image-select', section, block)"
          />
        </template>
      </div>
    </div>
    <div
      class="page-footer-band"
      data-testid="page-footer-band"
      :style="footerStyle"
      :contenteditable="footerEditable"
      :data-hf-editable="footerEditable ? '1' : undefined"
      spellcheck="false"
      @dblclick="$emit('dblclick-footer')"
      @beforeinput.stop="footerEditable ? $emit('hf-beforeinput', $event) : undefined"
      @keydown.stop="footerEditable ? $emit('hf-keydown', $event) : undefined"
      @compositionstart.stop="footerEditable ? $emit('hf-compositionstart') : undefined"
      @compositionend.stop="footerEditable ? $emit('hf-compositionend', $event) : undefined"
    >
      <DocBlockEntries :entries="footerEntries" />
    </div>
  </div>
</template>

<style scoped>
/* A page-shaped surface sized by the layout's own geometry (inline style,
   real centimetres). 12px of margin is AC-2's visible gap between pages;
   overflow clips provisional overhang while measurement settles. */
.page {
  position: relative;
  margin: 0 auto 12px;
  background: #fff;
  box-shadow: 0 1px 3px rgb(0 0 0 / 0.3);
  box-sizing: border-box;
  overflow: hidden;
}

.page:last-child {
  margin-bottom: 0;
}

/* The content area: padding carries the page margins (inline style), so the
   blocks flow in normal document flow exactly where the engine placed
   them. */
.page-content {
  box-sizing: border-box;
  height: 100%;
}

/* Body content dims while a header/footer is being edited (SPEC-029
 * AC-1, AC-2) — opacity, not visibility, so the dimmed text stays a
 * legible reading reference the way Word's own dimming does.
 * `pointer-events: none` sits on the margin box (the entries' own
 * container), not `.page-content` itself, so a double-click still lands on
 * `.page-content`'s own dblclick-body listener above it in the DOM and
 * exits editing (AC-8) instead of being swallowed by the dimmed area. */
.page-margin-dimmed {
  opacity: 0.35;
  pointer-events: none;
}

/* Header/footer bands (SPEC-029): positioned within the page's own
   top/bottom margin — inline top/left/right/bottom come from
   DocumentView's headerStyle/footerStyle, computed from the section's
   margins and a fixed header/footer distance from the page edge. */
.page-header-band,
.page-footer-band {
  position: absolute;
  overflow: hidden;
  box-sizing: border-box;
  display: flex;
  flex-direction: column;
}

/* Header text sits at the bottom of the top margin, just above the body
   text's own top edge; footer text sits at the top of the bottom margin,
   just below the body text's own bottom edge — matching where Word
   positions header/footer content within the margin band. */
.page-header-band {
  justify-content: flex-end;
  padding-bottom: 0.2cm;
}

.page-footer-band {
  justify-content: flex-start;
  padding-top: 0.2cm;
}

.page-header-band[data-hf-editable],
.page-footer-band[data-hf-editable] {
  outline: 1px dashed #1a73e8;
  outline-offset: 2px;
}

/* Marks where the page margins end and content begins — an outline, not a
   border, so the rendered content geometry matches the engine's
   margin-box math exactly. */
.page-margin {
  box-sizing: border-box;
  height: 100%;
  outline: 1px dashed #c8c8c8;
  outline-offset: -1px;
}

/* Multi-column page (SPEC-030): the column boxes are absolutely positioned
   within the content area at the widths/offsets the engine resolved, so each
   column is an independent normal-flow column exactly where it was measured.
   overflow:hidden contains anything wider than the column (AC-15). */
.page-margin.has-columns {
  position: relative;
}

.doc-column {
  position: absolute;
  top: 0;
  bottom: 0;
  box-sizing: border-box;
  overflow: hidden;
}

/* "Line between columns" (AC-4): a hairline rule centred in the gutter,
   spanning the content height. */
.doc-column-rule {
  position: absolute;
  top: 0;
  bottom: 0;
  width: 0;
  border-left: 1px solid #9aa0a6;
  pointer-events: none;
}

/* DocPageBlock (used for the multi-column path above) has no <style> of its
   own — the SPEC-030 branch relies on Vue scoping the parent's data-v
   attribute onto a child component's root element, so these rules (kept in
   sync with DocBlockEntries.vue, which the single-column path below uses
   instead) still have to live here too. */
.doc-paragraph {
  margin: 0;
  font-weight: normal;
  font-size: inherit;
}

/* The list marker (SPEC-018): painted from the derived data-list-marker
   attribute, never part of the editable text — caret offsets and the
   model/DOM text equality check stay untouched. */
.doc-paragraph[data-list-marker]::before {
  content: attr(data-list-marker);
  display: inline-block;
  min-width: 0.635cm;
  margin-inline-start: -0.635cm;
  white-space: nowrap;
}

/* A manual break block (SPEC-028): always in the DOM (it is content the
   caret can delete), invisible until formatting marks are shown, then a
   labelled dashed rule (AC-11, AC-16). The label is a pseudo-element so
   the canvas text stays byte-identical to the model. */
.doc-break {
  display: block;
  height: 0;
  overflow: hidden;
  user-select: none;
}

.marks-on .doc-break {
  height: auto;
  overflow: visible;
  margin: 2px 0;
  border-top: 1px dashed #9aa0a6;
  color: #5f6368;
  font-size: 0.7rem;
  line-height: 1.3;
  text-align: center;
}

.marks-on .doc-break::before {
  content: attr(data-marker-label);
}

/* A section break's marker (SPEC-028 AC-16): rendered only while marks are
   on, and absolutely positioned with zero flow height so showing marks
   never re-paginates the section boundary itself. */
.doc-section-marker {
  display: block;
  position: relative;
  height: 0;
  overflow: visible;
  user-select: none;
}

.marks-on .doc-section-marker::before {
  content: attr(data-marker-label);
  position: absolute;
  top: 0;
  right: 0;
  left: 0;
  border-top: 1px dashed #9aa0a6;
  color: #5f6368;
  font-size: 0.7rem;
  line-height: 1.3;
  text-align: center;
}

/* Space/tab marks (AC-14, AC-15): the real characters render (selection,
   caret math and copy all see the model text); the visible mark paints
   over them from a pseudo-element. */
.mark-space {
  position: relative;
}

.mark-space::before {
  content: "·";
  position: absolute;
  left: 50%;
  transform: translateX(-50%);
  color: #80868b;
  pointer-events: none;
}

.mark-tab {
  position: relative;
}

.mark-tab::before {
  content: "→";
  position: absolute;
  left: 0;
  color: #80868b;
  pointer-events: none;
}

/* The pilcrow (AC-13) — suppressed on a slice that continues onto the next
   page: the paragraph has not ended there. */
.marks-on .doc-paragraph:not([data-slice-more])::after {
  content: "¶";
  color: #80868b;
}
</style>
