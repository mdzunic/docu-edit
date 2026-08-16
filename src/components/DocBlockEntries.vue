<script setup lang="ts">
// The shared block-entry renderer (SPEC-028, split out for SPEC-029): the
// `v-for` DocPage's body content renders, factored out so header/footer
// bands (SPEC-029) can render the exact same paragraph/table/image markup
// without a second copy of it. Entries arrive fully precomputed (tag, style
// objects, run render lists) — this component holds no document logic of
// its own, same as DocPage before the split. Runtime prop/emit
// declarations, not the generic form — this script is parsed by espree.
import { computed } from "vue";
import DocImageBlock from "./DocImageBlock.vue";
import DocShapeBlock from "./DocShapeBlock.vue";
import DocChartBlock from "./DocChartBlock.vue";
import DocTableBlock from "./DocTableBlock.vue";
import DocTocBlock from "./DocTocBlock.vue";
import DocHeadingToggle from "./DocHeadingToggle.vue";
import { asRenderEntries, noRenderEntries } from "../lib/pageRender";

const props = defineProps({
  entries: { type: Array, default: () => noRenderEntries() },
  cellRange: { type: Object, default: null },
});

defineEmits([
  "image-select",
  "shape-select",
  "shape-edit-text",
  "shape-text-beforeinput",
  "shape-text-keydown",
  "shape-text-compositionstart",
  "shape-text-compositionend",
  "chart-select",
  "chart-edit-data",
  "toc-action",
  "outline-toggle",
]);

const renderEntries = computed(() => asRenderEntries(props.entries));

function componentFor(entry = noRenderEntries()[0]) {
  if (entry.kind === "table") {
    return DocTableBlock;
  }
  if (entry.kind === "image") {
    return DocImageBlock;
  }
  if (entry.kind === "shape") {
    return DocShapeBlock;
  }
  if (entry.kind === "chart") {
    return DocChartBlock;
  }
  if (entry.kind === "toc") {
    return DocTocBlock;
  }
  return entry.tag || "div";
}
</script>

<template>
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
    :data-tracked-change="entry.hasTrackedChange ? '1' : undefined"
    :block="entry.blockProp"
    :section="entry.sectionProp"
    :block-index="entry.blockIndexProp"
    :list-definitions="entry.listDefinitions"
    :display-mode="entry.trackChangesDisplayMode"
    :cell-range="entry.kind === 'table' ? cellRange : undefined"
    :src="entry.imageSrc"
    :selected="entry.kind === 'shape' ? entry.shapeSelected : entry.kind === 'chart' ? entry.chartSelected : entry.imageSelected"
    :drawing-index="entry.drawingIndex"
    :editing-text="entry.kind === 'shape' ? entry.shapeEditing : undefined"
    :stale-headings="entry.kind === 'toc' ? entry.tocStaleHeadings : undefined"
    :stale-pagination="entry.kind === 'toc' ? entry.tocStalePagination : undefined"
    @select="entry.kind === 'image' ? $emit('image-select', entry.sectionProp, entry.blockIndexProp) : entry.kind === 'shape' ? $emit('shape-select', entry.sectionProp, entry.blockIndexProp, $event) : entry.kind === 'chart' ? $emit('chart-select', entry.sectionProp, entry.blockIndexProp, $event) : undefined"
    @edit-text="entry.kind === 'shape' ? $emit('shape-edit-text', entry.sectionProp, entry.blockIndexProp) : undefined"
    @edit-data="entry.kind === 'chart' ? $emit('chart-edit-data', entry.sectionProp, entry.blockIndexProp) : undefined"
    @text-beforeinput="entry.kind === 'shape' ? $emit('shape-text-beforeinput', $event) : undefined"
    @text-keydown="entry.kind === 'shape' ? $emit('shape-text-keydown', $event) : undefined"
    @text-compositionstart="entry.kind === 'shape' ? $emit('shape-text-compositionstart') : undefined"
    @text-compositionend="entry.kind === 'shape' ? $emit('shape-text-compositionend', $event) : undefined"
    @action="entry.kind === 'toc' ? $emit('toc-action', $event) : undefined"
  >
    <DocHeadingToggle
      v-if="entry.outlineToggle"
      :collapsed="entry.outlineCollapsed"
      @toggle="$emit('outline-toggle', entry.dataSection, entry.dataBlock)"
    />
    <template v-if="entry.runs && entry.runs.length > 0">
      <template
        v-for="run in entry.runs"
        :key="run.key"
      >
        <!-- The plain (no-marks) variants keep a single text child, so
             the span's firstChild stays the run's own text node — DOM
             shapes older specs' tests and tooling rely on. -->
        <!-- A note reference mark (SPEC-031): renders its computed number
             but stands for the one-character model placeholder carried in
             data-caret-atom (see lib/domCaret.ts's caret-atom mapping). -->
        <span
          v-if="run.atomText"
          :class="run.atomClass || 'doc-note-ref'"
          contenteditable="false"
          :data-caret-atom="run.atomText"
          :data-note-id="run.noteId"
          :data-ref-target="run.refTarget"
          :data-testid="run.atomClass === 'doc-merge-chip' ? 'merge-chip' : run.atomClass && run.atomClass.indexOf('doc-field') === 0 ? 'field' : undefined"
          :title="run.title"
          :style="run.styleObj"
          :lang="run.lang"
        >{{ run.text }}</span>
        <!-- A tracked-change run hidden by the current display mode
             (SPEC-039 AC-9): renders no visible text, but still carries its
             full model text via data-caret-atom so caret offsets land where
             they would if the text were showing. -->
        <span
          v-else-if="run.caretAtom"
          contenteditable="false"
          :data-caret-atom="run.caretAtom"
          :style="run.styleObj"
          :lang="run.lang"
        />
        <a
          v-else-if="run.href && !run.segments"
          data-doc-link=""
          :href="run.href"
          :target="run.target"
          :rel="run.rel"
          :style="run.styleObj"
          :lang="run.lang"
        >{{ run.text }}</a>
        <a
          v-else-if="run.href"
          data-doc-link=""
          :href="run.href"
          :target="run.target"
          :rel="run.rel"
          :style="run.styleObj"
          :lang="run.lang"
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
          :lang="run.lang"
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
          :lang="run.lang"
        >{{ run.text }}</span>
      </template>
    </template>
    <br v-else-if="entry.needsBr">
  </component>
</template>

<style scoped>
/* These rules moved here from DocPage.vue when its per-entry markup did
 * (SPEC-029): Vue's scoped CSS only scopes the rightmost compound selector,
 * so `.marks-on .doc-paragraph::after` still matches across the
 * DocumentView -> DocPage boundary (`.marks-on` lives on DocumentView's own
 * root) but only when `.doc-paragraph` itself is an element THIS component
 * owns — it stopped being one of DocPage's own elements the moment the
 * entries `v-for` moved here, so these rules have to live wherever the
 * entries actually render now. */
.doc-paragraph {
  margin: 0;
  font-weight: normal;
  font-size: inherit;
}

/* SPEC-050: a Markdown thematic break (`---`) — an ordinary empty editable
   paragraph in the "Horizontal Rule" style, drawn with a rule so it reads as
   one instead of a blank line. A pseudo-element (not a real `<hr>`) so the
   paragraph keeps the same editable/caret DOM shape as any other. */
.doc-horizontal-rule {
  min-height: 0.1cm;
  position: relative;
}
.doc-horizontal-rule::before {
  content: "";
  position: absolute;
  left: 0;
  right: 0;
  top: 50%;
  border-top: 1px solid #9aa0a6;
}

/* SPEC-039 AC-5: a change bar in the margin for any paragraph carrying a
   tracked change — a plain vertical rule outside the paragraph's own box
   (a box-shadow, not a pseudo-element, so it never collides with the list
   marker's ::before), shown in "All markup"/"Simple markup" only. */
.doc-paragraph[data-tracked-change] {
  box-shadow: -0.3cm 0 0 0 #1a73e8;
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

/* A note reference mark (SPEC-031): clickable (AC-15), visually a link-blue
   superscript number — the superscript itself comes from the run's own
   vertAlign formatting in the inline style. */
.doc-note-ref {
  color: #1a73e8;
  cursor: pointer;
  user-select: none;
}

.doc-note-ref:hover {
  text-decoration: underline;
}

/* A mail-merge chip (SPEC-046): a pill so it reads as one atomic unit
   distinct from surrounding text, matching the "chip" language the spec
   itself uses. */
.doc-merge-chip {
  background: #e8f0fe;
  border-radius: 3px;
  padding: 0 2px;
  color: #1a73e8;
}

/* A SPEC-053 field: a caret atom whose displayed value is computed. Reads as
   ordinary text (no pill) until shaded. */
.doc-field {
  cursor: default;
  user-select: none;
}

/* Field shading (SPEC-053 AC-10, AC-11): grey background when the caret sits
   on the field or "View → Field shading" is on. Never prints or exports
   (AC-12) — the print/preview rule below drops it. */
.doc-field-shaded {
  background: #d9d9d9;
}

/* An unsupported field type or unsupported date picture (SPEC-053 AC-42):
   shown read-only, its `title` carries the explanatory message. */
.doc-field-readonly {
  cursor: help;
}

/* AC-12: field shading is a screen-only affordance. In print preview and on
   the printed/exported page the grey background is removed. */
@media print {
  .doc-field-shaded {
    background: transparent !important;
  }
}
</style>
