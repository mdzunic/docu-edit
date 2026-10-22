<script setup lang="ts">
// The transient drawing-surface overlays that sit inside DocumentView's
// `.pages` container, between the shape handles and the document content:
// snap guides (SPEC-070), the anchor marker (SPEC-051), the connector layer
// and its toolbar (SPEC-072), the two in-progress draw previews and the
// point-edit overlay (SPEC-071). All are purely presentational or delegate
// straight back to the composable they are handed, so they were extracted out
// of the DocumentView backbone to keep it within its line budget — the same
// move SPEC-064 made for PageMarkupOverlays.
//
// The elements are declared as fragment roots (no wrapper) so they stay direct
// children of `.pages` and their `position: absolute` still resolves against
// it, in the paint order they had inline.
//
// Runtime prop declarations per the .vue espree constraint (see DocPage.vue);
// the composable objects arrive whole (`connectors`, `pointEdit`) so the
// wiring in DocumentView stays one binding each, as PointEditOverlay already
// does with its controller.
import ConnectorOverlay from "./ConnectorOverlay.vue";
import ConnectorToolbar from "./ConnectorToolbar.vue";
import DrawingGuidesOverlay from "./DrawingGuidesOverlay.vue";
import FreeformDrawPreview from "./FreeformDrawPreview.vue";
import PointEditOverlay from "./PointEditOverlay.vue";
import ShapeDrawPreview from "./ShapeDrawPreview.vue";

defineProps({
  /** SnapGuide[] — empty while no drag is snapping. */
  snapGuides: { type: Array, default: () => [] },
  /** { left, top } in `.pages` px, absent when no shape is selected. */
  anchorMarker: { type: Object, default: undefined },
  /** useConnectors' whole return value. */
  connectors: { type: Object, required: true },
  /** The shape a drag is about to insert, with its box; absent between drags. */
  drawPreview: { type: Object, default: undefined },
  /** The in-progress freeform path's `d`, or "" when not drawing one. */
  freeformD: { type: String, default: "" },
  /** useShapes' point-edit controller. */
  pointEdit: { type: Object, required: true },
});
</script>

<template>
  <DrawingGuidesOverlay
    v-if="snapGuides.length"
    :guides="snapGuides"
  />
  <div
    v-if="anchorMarker"
    class="shape-anchor-marker print-exclude"
    data-testid="shape-anchor-marker"
    :style="{ left: `${anchorMarker.left}px`, top: `${anchorMarker.top}px` }"
  />
  <ConnectorOverlay
    :renders="connectors.connectorRenders.value"
    :hover-sites="connectors.hoverSites.value || undefined"
    :draw-from="connectors.drawFrom.value || undefined"
    :draw-pointer="connectors.drawPointer.value || undefined"
    @select="connectors.selectConnector"
    @endpoint-drag="connectors.startEndpointDrag"
    @adjust-drag="connectors.startAdjustDrag"
  />
  <ConnectorToolbar
    v-if="connectors.selectedConnector.value"
    v-bind="connectors.connectorToolbarProps.value"
    :style="connectors.connectorToolbarStyle.value"
    @set-kind="connectors.setKind"
    @set-outline="connectors.setOutline"
    @set-arrowheads="connectors.setArrowheads"
    @set-wrap="connectors.setWrap"
    @set-alt-text="connectors.setAltText"
    @set-z-order="connectors.setZOrder"
    @delete="connectors.removeSelectedConnector"
  />
  <ShapeDrawPreview
    v-if="drawPreview"
    class="print-exclude"
    :shape="drawPreview.shape"
    :left="drawPreview.left"
    :top="drawPreview.top"
    :width="drawPreview.width"
    :height="drawPreview.height"
  />
  <FreeformDrawPreview
    v-if="freeformD"
    :d="freeformD"
  />
  <PointEditOverlay
    v-if="pointEdit.frame.value"
    class="print-exclude"
    :controller="pointEdit"
  />
</template>

<style scoped>
/* The anchor marker (SPEC-051 AC-27): a small flag at the start of the
   paragraph a selected shape is anchored to, moving with it on re-anchor
   (AC-28) — same overlay convention as the resize/rotation handles, one
   layer below them (z-index) since it never needs to receive a pointer. */
.shape-anchor-marker {
  position: absolute;
  z-index: 14;
  pointer-events: none;
  width: 10px;
  height: 10px;
  margin-left: -5px;
  margin-top: -2px;
  background: var(--accent);
  clip-path: polygon(0 0, 100% 50%, 0 100%);
}
</style>
