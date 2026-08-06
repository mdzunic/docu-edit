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
import { computed } from "vue";
import DocImageBlock from "./DocImageBlock.vue";
import DocTableBlock from "./DocTableBlock.vue";
import { asRenderEntries, noRenderEntries } from "../lib/pageRender";

const props = defineProps({
  pageStyle: { type: Object, required: true },
  contentStyle: { type: Object, required: true },
  /** Precomputed render entries (lib/pageRender.RenderEntry) — the typed
   * default seed gives the template its element type (see StylePicker). */
  entries: { type: Array, default: () => noRenderEntries() },
  /** A virtualized page (SPEC-028): the box renders with its exact layout
   * geometry but none of its content — the pagination controller measures
   * its blocks off-DOM instead. */
  virtual: { type: Boolean, default: false },
  cellRange: { type: Object, default: null },
});

defineEmits(["image-select"]);

const renderEntries = computed(() => asRenderEntries(props.entries));

function componentFor(entry = noRenderEntries()[0]) {
  if (entry.kind === "table") {
    return DocTableBlock;
  }
  if (entry.kind === "image") {
    return DocImageBlock;
  }
  return entry.tag || "div";
}
</script>

<template>
  <div
    class="page"
    data-testid="document-page"
    :data-virtual="virtual ? '1' : undefined"
    :style="pageStyle"
  >
    <div
      class="page-content"
      :style="contentStyle"
    >
      <div
        class="page-margin"
        data-testid="page-margin"
      >
        <component
          :is="componentFor(entry)"
          v-for="entry in renderEntries"
          :key="entry.key"
          v-memo="entry.memo"
          :class="entry.className"
          :style="entry.styleObj"
          :contenteditable="entry.uneditable ? 'false' : undefined"
          :data-section="entry.dataSection"
          :data-block="entry.dataBlock"
          :data-slice-start="entry.sliceStart"
          :data-slice-more="entry.sliceMore"
          :data-break-kind="entry.breakKind"
          :data-section-marker="entry.sectionMarker"
          :data-marker-label="entry.markerLabel"
          :data-list-marker="entry.listMarker"
          :data-list-level="entry.listLevel"
          :block="entry.blockProp"
          :section="entry.sectionProp"
          :block-index="entry.blockIndexProp"
          :cell-range="entry.kind === 'table' ? cellRange : undefined"
          :src="entry.imageSrc"
          :selected="entry.imageSelected"
          @select="entry.kind === 'image' ? $emit('image-select', entry.sectionProp, entry.blockIndexProp) : undefined"
        >
          <template v-if="entry.runs && entry.runs.length > 0">
            <template
              v-for="run in entry.runs"
              :key="run.key"
            >
              <!-- The plain (no-marks) variants keep a single text child, so
                   the span's firstChild stays the run's own text node — DOM
                   shapes older specs' tests and tooling rely on. -->
              <a
                v-if="run.href && !run.segments"
                data-doc-link=""
                :href="run.href"
                :target="run.target"
                :rel="run.rel"
                :style="run.styleObj"
              >{{ run.text }}</a>
              <a
                v-else-if="run.href"
                data-doc-link=""
                :href="run.href"
                :target="run.target"
                :rel="run.rel"
                :style="run.styleObj"
              ><template
                v-for="(segment, segmentIndex) in run.segments"
                :key="segmentIndex"
              ><span
                v-if="segment.kind !== 'text'"
                :class="segment.kind === 'tab' ? 'mark-tab' : 'mark-space'"
              >{{ segment.text }}</span><template v-else>{{ segment.text }}</template></template></a>
              <span
                v-else-if="run.segments"
                :style="run.styleObj"
              ><template
                v-for="(segment, segmentIndex) in run.segments"
                :key="segmentIndex"
              ><span
                v-if="segment.kind !== 'text'"
                :class="segment.kind === 'tab' ? 'mark-tab' : 'mark-space'"
              >{{ segment.text }}</span><template v-else>{{ segment.text }}</template></template></span>
              <span
                v-else
                :style="run.styleObj"
              >{{ run.text }}</span>
            </template>
          </template>
          <br v-else-if="entry.needsBr">
        </component>
      </div>
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

/* Marks where the page margins end and content begins — an outline, not a
   border, so the rendered content geometry matches the engine's
   margin-box math exactly. */
.page-margin {
  box-sizing: border-box;
  height: 100%;
  outline: 1px dashed #c8c8c8;
  outline-offset: -1px;
}

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
