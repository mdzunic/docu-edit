<script setup lang="ts">
// SPEC-070 / SPEC-034: the painted half of the connectors on ONE page.
//
// Connectors used to be drawn only by ConnectorOverlay, a single SVG stretched
// over the whole `.pages` stack and marked `print-exclude` — so they were
// screen-only chrome: absent from the print preview, absent from printed
// sheets, and invisible to the PDF extractor, which reads one realized `.page`
// element at a time. This layer lives INSIDE its page box instead, which puts
// every connector on the right sheet, in the preview, and within reach of
// domExtract. The overlay keeps the interactive chrome (hit paths, handles,
// site hints), which stays screen-only.
//
// Geometry arrives fully computed in page-local px (useConnectors measures it
// and lib/core/connectorPaint turns it into paint data). Runtime prop
// declarations, not the generic form — this script is parsed by espree (see
// AppShell.vue).
import { computed } from "vue";
import { asConnectorPaint, noConnectorPaint } from "../lib/core/connectorPaint";

const props = defineProps({
  renders: { type: Array, default: () => noConnectorPaint() },
});

const paints = computed(() => asConnectorPaint(props.renders));
</script>

<template>
  <svg
    v-if="paints.length > 0"
    class="doc-connector-layer"
    data-testid="connector-paint-layer"
    data-canvas-ignore="1"
    aria-hidden="true"
    focusable="false"
  >
    <g
      v-for="paint in paints"
      :key="paint.key"
    >
      <path
        class="doc-connector-line"
        :d="paint.d"
        fill="none"
        :stroke="paint.color"
        :stroke-width="paint.widthPx"
        :stroke-dasharray="paint.dash || undefined"
      />
      <path
        v-for="(head, headIndex) in paint.heads"
        :key="headIndex"
        class="doc-connector-arrow"
        :d="head.d"
        :fill="head.filled ? paint.color : 'none'"
        :stroke="head.filled ? 'none' : paint.color"
        :stroke-width="head.filled ? undefined : paint.widthPx"
      />
    </g>
  </svg>
</template>

<style scoped>
/* Covers its page box exactly, so the paint data's page-local px are the
   SVG's own user units. `overflow: visible` lets a head that reaches a
   hair past an endpoint show; the page's own `overflow: hidden` still clips
   anything that leaves the sheet. It never takes the pointer — selection and
   dragging belong to ConnectorOverlay.

   The z-index keeps the behaviour the old `.pages`-level overlay had: a
   connector paints above the shapes it joins. renderFormatting gives a
   floating drawing `100000 + zOrder`, and zOrder is its index among the
   document's drawings, so this clears any realistic document — a connector's
   own z-order against a shape's is left for the day the spec asks for it
   (the model already carries it, and reorderShapeZ already sets it). */
.doc-connector-layer {
  position: absolute;
  inset: 0;
  width: 100%;
  height: 100%;
  overflow: visible;
  z-index: 1000000;
  pointer-events: none;
}
</style>
