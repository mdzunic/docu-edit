<script setup lang="ts">
// One rendered block entry of a page (SPEC-028, extracted in SPEC-030): the
// `<component :is>` element plus its run children. Pulled out of DocPage.vue
// so the single-column flow and the multi-column column boxes render the
// exact same markup from one place — the component's single root element IS
// the block element, so the produced DOM is byte-identical to the previous
// inline template. Runtime prop/emit declarations (espree, see AppShell.vue).
import { computed } from "vue";
import DocImageBlock from "./DocImageBlock.vue";
import DocTableBlock from "./DocTableBlock.vue";
import DocTocBlock from "./DocTocBlock.vue";
import { asRenderEntry, noRenderEntries } from "../lib/pageRender";

const props = defineProps({
  /** One lib/pageRender.RenderEntry — typed via the default seed. */
  entry: { type: Object, default: () => noRenderEntries()[0] },
  cellRange: { type: Object, default: null },
});

defineEmits(["image-select", "toc-action"]);

function componentFor(entry = noRenderEntries()[0]) {
  if (entry.kind === "table") {
    return DocTableBlock;
  }
  if (entry.kind === "image") {
    return DocImageBlock;
  }
  if (entry.kind === "toc") {
    return DocTocBlock;
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
    :cell-range="entry.kind === 'table' ? cellRange : undefined"
    :src="entry.imageSrc"
    :selected="entry.imageSelected"
    :stale-headings="entry.kind === 'toc' ? entry.tocStaleHeadings : undefined"
    :stale-pagination="entry.kind === 'toc' ? entry.tocStalePagination : undefined"
    @select="entry.kind === 'image' ? $emit('image-select', entry.sectionProp, entry.blockIndexProp) : undefined"
    @action="entry.kind === 'toc' ? $emit('toc-action', $event) : undefined"
  >
    <template v-if="entry.runs && entry.runs.length > 0">
      <template
        v-for="run in entry.runs"
        :key="run.key"
      >
        <!-- The plain (no-marks) variants keep a single text child, so the
             span's firstChild stays the run's own text node — DOM shapes
             older specs' tests and tooling rely on. -->
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
</template>
