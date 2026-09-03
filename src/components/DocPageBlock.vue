<script setup lang="ts">
// One rendered block entry of a page (SPEC-028, extracted in SPEC-030): the
// `<component :is>` element plus its run children. Pulled out of DocPage.vue
// so the single-column flow and the multi-column column boxes render the
// exact same markup from one place — the component's single root element IS
// the block element, so the produced DOM is byte-identical to the previous
// inline template. Runtime prop/emit declarations (espree, see AppShell.vue).
import { computed, inject } from "vue";
import DocImageBlock from "./DocImageBlock.vue";
import DocShapeBlock from "./DocShapeBlock.vue";
import DocChartBlock from "./DocChartBlock.vue";
import DocTableBlock from "./DocTableBlock.vue";
import DocTocBlock from "./DocTocBlock.vue";
import DocTofBlock from "./DocTofBlock.vue";
import DocHeadingToggle from "./DocHeadingToggle.vue";
import EquationView from "./EquationView.vue";
import { asRenderEntry, noRenderEntries } from "../lib/pageRender";
import { noopEquationActivate } from "../lib/documentModel";

const props = defineProps({
  /** One lib/pageRender.RenderEntry — typed via the default seed. */
  entry: { type: Object, default: () => noRenderEntries()[0] },
  cellRange: { type: Object, default: null },
});

// SPEC-056: clicking an equation opens the editor. Provided by DocumentView.
const activateEquation = inject("activateEquation", noopEquationActivate);

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
  "tof-action",
  "outline-toggle",
]);

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
  if (entry.kind === "tof") {
    return DocTofBlock;
  }
  return entry.tag || "div";
}

// Reactive alias so the template reads like the original inline version and
// still updates when the parent hands this instance a new entry (v-memo may
// reuse the instance across relayouts).
const entry = computed(() => asRenderEntry(props.entry));
</script>

<template>
  <component
    :is="componentFor(entry)"
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
    :list-definitions="entry.listDefinitions"
    :styles="entry.styles"
    :row-start="entry.rowStart"
    :row-end="entry.rowEnd"
    :data-table-section="entry.dataTableSection"
    :data-table-block="entry.dataTableBlock"
    :data-row-start="entry.rowStart"
    :display-mode="entry.trackChangesDisplayMode"
    :cell-range="entry.kind === 'table' ? cellRange : undefined"
    :src="entry.imageSrc"
    :selected="entry.kind === 'shape' ? entry.shapeSelected : entry.kind === 'chart' ? entry.chartSelected : entry.imageSelected"
    :drawing-index="entry.drawingIndex"
    :editing-text="entry.kind === 'shape' ? entry.shapeEditing : undefined"
    :stale-headings="entry.kind === 'toc' ? entry.tocStaleHeadings : undefined"
    :stale-entries="entry.kind === 'tof' ? entry.tofStaleEntries : undefined"
    :stale-pagination="entry.kind === 'toc' ? entry.tocStalePagination : entry.kind === 'tof' ? entry.tofStalePagination : undefined"
    @select="entry.kind === 'image' ? $emit('image-select', entry.sectionProp, entry.blockIndexProp) : entry.kind === 'shape' ? $emit('shape-select', entry.sectionProp, entry.blockIndexProp, $event) : entry.kind === 'chart' ? $emit('chart-select', entry.sectionProp, entry.blockIndexProp, $event) : undefined"
    @edit-text="entry.kind === 'shape' ? $emit('shape-edit-text', entry.sectionProp, entry.blockIndexProp) : undefined"
    @edit-data="entry.kind === 'chart' ? $emit('chart-edit-data', entry.sectionProp, entry.blockIndexProp) : undefined"
    @text-beforeinput="entry.kind === 'shape' ? $emit('shape-text-beforeinput', $event) : undefined"
    @text-keydown="entry.kind === 'shape' ? $emit('shape-text-keydown', $event) : undefined"
    @text-compositionstart="entry.kind === 'shape' ? $emit('shape-text-compositionstart') : undefined"
    @text-compositionend="entry.kind === 'shape' ? $emit('shape-text-compositionend', $event) : undefined"
    @action="entry.kind === 'toc' ? $emit('toc-action', $event) : entry.kind === 'tof' ? $emit('tof-action', $event) : undefined"
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
        <!-- An equation (SPEC-056): laid-out equation as a caret atom, kept
             in sync with DocBlockEntries.vue's single-column branch. -->
        <span
          v-if="run.equation"
          class="doc-equation"
          contenteditable="false"
          :data-caret-atom="run.atomText"
          :style="run.styleObj"
          @click="activateEquation(run.equation)"
        ><EquationView :equation="run.equation" /></span>
        <!-- The plain (no-marks) variants keep a single text child, so the
             span's firstChild stays the run's own text node — DOM shapes
             older specs' tests and tooling rely on. -->
        <a
          v-else-if="run.href && !run.segments"
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
</template>
