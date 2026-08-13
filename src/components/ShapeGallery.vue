<script setup lang="ts">
// "Insert -> Shape"'s gallery (SPEC-051 AC-1, AC-2): exactly the Shape
// catalogue's geometries, grouped as Lines/Basic shapes/Arrows/Callouts.
// Picking one arms draw mode (AC-3) — this component only reports the
// pick, DocumentView.vue's useShapes owns what happens next. Runtime
// prop/emit declarations per the espree constraint (see AppShell.vue).
import { SHAPE_CATALOGUE } from "../lib/documentModel";

defineEmits(["pick"]);

const GROUPS = ["Lines", "Basic shapes", "Arrows", "Callouts"];

function itemsFor(group = "") {
  return SHAPE_CATALOGUE.filter((item) => item.group === group);
}
</script>

<template>
  <div
    class="shape-gallery"
    data-testid="shape-gallery"
    role="menu"
    aria-label="Shape gallery"
  >
    <div
      v-for="group in GROUPS"
      :key="group"
      class="shape-gallery-group"
    >
      <div class="shape-gallery-group-label">
        {{ group }}
      </div>
      <div class="shape-gallery-items">
        <button
          v-for="item in itemsFor(group)"
          :key="item.geometry"
          type="button"
          class="shape-gallery-item"
          :data-testid="`shape-gallery-${item.geometry}`"
          :title="item.label"
          @mousedown.prevent
          @click="$emit('pick', item.geometry)"
        >
          {{ item.label }}
        </button>
      </div>
    </div>
  </div>
</template>

<style scoped>
.shape-gallery {
  display: flex;
  flex-direction: column;
  gap: 0.1875rem;
  padding: 0.1875rem;
  min-width: 220px;
}

.shape-gallery-group-label {
  font-size: 0.5625rem;
  font-weight: 600;
  text-transform: uppercase;
  color: #5f6368;
  margin-bottom: 0.0625rem;
}

.shape-gallery-items {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 0.125rem;
}

.shape-gallery-item {
  font: inherit;
  font-size: 0.6875rem;
  line-height: 1.1;
  padding: 0.0625rem 0.3rem;
  border: 1px solid #dadce0;
  border-radius: 4px;
  background: #fff;
  color: #202124;
  cursor: pointer;
  text-align: left;
  white-space: nowrap;
}

.shape-gallery-item:hover {
  background: #f1f3f4;
}
</style>
