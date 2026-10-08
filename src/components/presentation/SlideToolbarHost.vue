<script setup lang="ts">
// The ribbon strips' bindings (SPEC-097's PresentationView extraction, AC-58):
// renders the Home strip's PresentationToolbar or the Insert strip's
// SlideInsertBar with every prop and emit wired to the stage-edit controller
// (usePresentationEditing), the text-depth controller (useSlideTextDepth) and,
// for Insert, the chart controller — the two prop/emit blocks that used to sit
// in the view, so PresentationView keeps one `<SlideToolbarHost strip=… />`
// per ribbon slot. Pure wiring: no state of its own beyond the derived
// `hasGroup`. Runtime prop declarations (espree lints .vue scripts).
import { computed } from "vue";
import PresentationToolbar from "./PresentationToolbar.vue";
import SlideInsertBar from "./SlideInsertBar.vue";

const props = defineProps({
  /** Which ribbon strip this host fills: "home" or "insert". */
  strip: { type: String, default: "home" },
  /** The usePresentationEditing() controller (selection, editing, arrange,
   * insertion, manipulation, recent colours, fill/stroke/colour commands). */
  stageEdit: { type: Object, required: true },
  /** The useSlideTextDepth() controller (lists, links). */
  textDepth: { type: Object, required: true },
  /** The deck theme (SPEC-094), for the colour pickers' theme rows. */
  theme: { type: Object, default: null },
  /** The useSlideCharts() controller (Insert ▸ Chart). */
  charts: { type: Object, default: null },
  /** The current slide (Insert: whether a selected frame is a group). */
  slide: { type: Object, default: null },
});

const e = props.stageEdit;
const t = props.textDepth;

const hasGroup = computed(() =>
  e.selection.selectedFrameIds.value.some(
    (id = "") => props.slide?.frames.find((frame = Object()) => frame.id === id)?.kind === "group",
  ),
);

function setAlignReference(value = "selection") {
  e.arrange.alignReference.value = value;
}
function pickChart(chartType = "") {
  props.charts?.pickChartType(chartType);
}
</script>

<template>
  <PresentationToolbar
    v-if="strip === 'home'"
    :state="e.editing.toolbarState.value ?? undefined"
    :enabled="e.editing.toolbarEnabled.value"
    :recent-colors="e.recentColors.value"
    :theme="theme"
    :list-geometry="t.listGeometry.value"
    :list-error="t.listError.value"
    :image-accept="t.imageAccept"
    :link-enabled="t.canLinkText.value"
    @bold="e.editing.toggleBold()"
    @italic="e.editing.toggleItalic()"
    @underline="e.editing.toggleUnderline()"
    @font="e.editing.setFontFamily($event)"
    @size="e.editing.setFontSize($event)"
    @color="e.setColor($event)"
    @align="e.editing.setAlign($event)"
    @bullet="e.editing.toggleBulletCmd()"
    @indent="e.editing.indent(1)"
    @outdent="e.editing.indent(-1)"
    @numbering="t.toggleNumbering()"
    @list-style="t.applyListStyle($event)"
    @bullet-char="t.setBulletChar($event)"
    @bullet-image="t.setBulletImage($event)"
    @start-at="t.setStartAt($event)"
    @list-indent="t.setListIndent($event)"
    @list-hanging="t.setListHanging($event)"
    @link="t.openLinkDialog()"
  />
  <SlideInsertBar
    v-else
    :selected-frame="e.selection.anchorFrame.value ?? undefined"
    :selected-count="e.arrange.selectedCount.value"
    :has-group="hasGroup"
    :can-distribute="e.arrange.canDistribute.value"
    :distribute-reason="e.arrange.distributeReason.value"
    :align-reference="e.arrange.alignReference.value"
    :theme="theme"
    :armed-kind="e.insertion.armed.value?.kind ?? ''"
    :can-link="t.linkableFrame.value !== null"
    @insert-text-box="e.insertion.armTextBox()"
    @insert-shape="e.insertion.armShape($event)"
    @insert-image="e.insertion.insertImageFile($event)"
    @insert-media="e.insertion.insertMediaFile($event)"
    @insert-table="(rows, cols) => e.insertion.insertTable(rows, cols)"
    @insert-chart="pickChart($event)"
    @arrange="e.manipulation.arrange($event)"
    @align="e.arrange.align($event)"
    @distribute="e.arrange.distribute($event)"
    @align-reference="setAlignReference($event)"
    @group="e.arrange.group()"
    @ungroup="e.arrange.ungroup()"
    @position-toggle="e.arrange.openPosition()"
    @set-fill="e.setFill($event)"
    @set-stroke="e.setStroke($event)"
    @link="t.openFrameLinkDialog()"
  />
</template>
