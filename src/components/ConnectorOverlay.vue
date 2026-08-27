<script setup lang="ts">
// SPEC-070: the page-level connector overlay — the painted lines, the
// endpoint/midpoint drag handles, the two-click draw rubber band and the
// connection-site hints. All geometry arrives in unscaled page px (the
// `.shape-handles` frame) from useConnectors; this component only paints it
// and reports the handle gestures back. Runtime prop/emit declarations, not
// the generic form — this script is parsed by espree (see AppShell.vue).
import { computed } from "vue";
import {
  asConnectorRenders,
  asDrawFrom,
  asDrawPoint,
  asSiteHint,
  noConnectorRenders,
  noSiteHint,
  noDrawPoint,
  noDrawFrom,
} from "../lib/connectorOverlay";

const props = defineProps({
  renders: { type: Array, default: () => noConnectorRenders() },
  hoverSites: { type: Object, default: () => noSiteHint() },
  drawFrom: { type: Object, default: () => noDrawFrom() },
  drawPointer: { type: Object, default: () => noDrawPoint() },
});

// Typed views of the runtime-declared props (see connectorOverlay.ts).
const renderList = computed(() => asConnectorRenders(props.renders));
const siteHint = computed(() => asSiteHint(props.hoverSites));
const drawFromPoint = computed(() => asDrawFrom(props.drawFrom));
const drawToPoint = computed(() => asDrawPoint(props.drawPointer));

defineEmits(["select", "endpoint-drag", "adjust-drag"]);

// Outline width is in points, the overlay in unscaled page px (1pt = 96/72px);
// dash patterns mirror the shape outline's.
function strokeWidthPx(widthPt = 1) {
  return Math.max(1, (widthPt || 1) * (96 / 72));
}
function dashArray(dash = "solid") {
  if (dash === "dash") {
    return "6 4";
  }
  if (dash === "dot") {
    return "1 3";
  }
  return undefined;
}
function previewPath(from = noDrawFrom(), to = noDrawPoint()) {
  if (!from || !to) {
    return "";
  }
  return `M ${from.point.x} ${from.point.y} L ${to.x} ${to.y}`;
}
</script>

<template>
  <svg
    v-if="renderList.length || drawFromPoint || siteHint"
    class="connector-overlay print-exclude"
    data-testid="connector-overlay"
  >
    <defs>
      <marker
        id="conn-arrow-end"
        viewBox="0 0 10 10"
        refX="9"
        refY="5"
        markerWidth="7"
        markerHeight="7"
        orient="auto-start-reverse"
      >
        <path d="M0 0 L10 5 L0 10 z" fill="context-stroke" />
      </marker>
      <marker
        id="conn-arrow-start"
        viewBox="0 0 10 10"
        refX="9"
        refY="5"
        markerWidth="7"
        markerHeight="7"
        orient="auto-start-reverse"
      >
        <path d="M0 0 L10 5 L0 10 z" fill="context-stroke" />
      </marker>
    </defs>
    <g
      v-for="render in renderList"
      :key="render.key"
    >
      <path
        class="connector-line"
        :d="render.d"
        fill="none"
        :stroke="render.outline ? render.outline.color : '#2e5395'"
        :stroke-width="strokeWidthPx(render.outline ? render.outline.widthPt : 1)"
        :stroke-dasharray="dashArray(render.outline ? render.outline.dash : 'solid')"
        :marker-end="render.endArrow && render.endArrow !== 'none' ? 'url(#conn-arrow-end)' : undefined"
        :marker-start="render.startArrow && render.startArrow !== 'none' ? 'url(#conn-arrow-start)' : undefined"
      />
      <path
        class="connector-hit"
        :d="render.d"
        fill="none"
        :data-testid="`connector-${render.location.section}-${render.location.block}`"
        @mousedown.stop="$emit('select', render.location)"
      />
      <template v-if="render.selected">
        <circle
          class="connector-endpoint-handle"
          data-testid="connector-handle-start"
          :cx="render.a.x"
          :cy="render.a.y"
          r="5"
          @mousedown.stop="$emit('endpoint-drag', render.location, 'start', $event)"
        />
        <circle
          class="connector-endpoint-handle"
          data-testid="connector-handle-end"
          :cx="render.b.x"
          :cy="render.b.y"
          r="5"
          @mousedown.stop="$emit('endpoint-drag', render.location, 'end', $event)"
        />
        <circle
          v-if="render.mid"
          class="connector-mid-handle"
          data-testid="connector-handle-mid"
          :cx="render.mid.x"
          :cy="render.mid.y"
          r="5"
          @mousedown.stop="$emit('adjust-drag', render.location, $event)"
        />
      </template>
    </g>
    <path
      v-if="drawFromPoint"
      class="connector-preview"
      fill="none"
      :d="previewPath(drawFromPoint, drawToPoint)"
    />
    <circle
      v-for="(site, siteIndex) in (siteHint ? siteHint.sites : [])"
      :key="`site-${siteIndex}`"
      class="connector-site"
      data-testid="connector-site"
      :cx="site.x"
      :cy="site.y"
      r="4"
    />
  </svg>
</template>

<style scoped>
.connector-overlay {
  position: absolute;
  inset: 0;
  width: 100%;
  height: 100%;
  overflow: visible;
  z-index: 31;
  pointer-events: none;
}

.connector-line {
  pointer-events: none;
}

.connector-hit {
  stroke: transparent;
  stroke-width: 12px;
  pointer-events: stroke;
  cursor: pointer;
}

.connector-endpoint-handle,
.connector-mid-handle {
  fill: var(--surface-base);
  stroke: var(--accent);
  stroke-width: 1.5px;
  pointer-events: auto;
  cursor: grab;
}

.connector-mid-handle {
  fill: var(--accent);
}

.connector-preview {
  stroke: var(--accent);
  stroke-width: 1.5px;
  stroke-dasharray: 4 3;
  pointer-events: none;
}

.connector-site {
  fill: var(--surface-base);
  stroke: var(--accent);
  stroke-width: 1.5px;
  pointer-events: none;
}
</style>
