<script setup lang="ts">
// SPEC-070: the page-level connector *chrome* — the transparent hit paths,
// the endpoint/midpoint drag handles, the two-click draw rubber band and the
// connection-site hints. All of it is screen-only (`print-exclude`), which is
// why the connector LINE is no longer painted here: it lived in this overlay,
// so it never reached the print preview, a printed sheet or the PDF. The line
// and its arrowheads now paint inside each page box (ConnectorPaintLayer);
// this overlay is purely the interaction surface stretched over the whole
// `.pages` stack.
//
// All geometry arrives in unscaled page px (the `.shape-handles` frame) from
// useConnectors; this component only places it and reports the handle
// gestures back. Runtime prop/emit declarations, not the generic form — this
// script is parsed by espree (see AppShell.vue).
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
    <g
      v-for="render in renderList"
      :key="render.key"
    >
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
