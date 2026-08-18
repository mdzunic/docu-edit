<script setup lang="ts">
// "Insert -> Shape"'s gallery (SPEC-051 AC-1, AC-2): exactly the Shape
// catalogue's geometries, grouped as Lines/Basic shapes/Arrows/Callouts.
// Picking one arms draw mode (AC-3) — this component only reports the
// pick, DocumentView.vue's useShapes owns what happens next.
//
// Each geometry shows as its own outline (shapeIcons.ts draws it from the
// same vertex lists the page canvas paints with), not its name: a fixed-size
// icon makes every cell the same width, so the rows line up with each other
// and the gallery sits on the menu's own 1rem text gutter. Names that no
// longer show stay reachable as the button's tooltip and accessible name.
//
// The buttons deliberately carry no menuitem role: MenuBar styles that role
// as a full-width command row and puts it in the Up/Down item ring, which
// a 14-cell grid has no business joining (same reasoning as the Insert
// menu's table-size grid — see MenuBar.vue's onPopupClick comment).
//
// Runtime prop/emit declarations per the espree constraint (see AppShell.vue).
import { asShapeGeometry, LINE_GEOMETRIES, SHAPE_CATALOGUE } from "../lib/core/shapes";
import { SHAPE_ICON_SIZE, shapeIconPath } from "../lib/core/shapeIcons";

defineEmits(["pick"]);

const GROUPS = ["Lines", "Basic shapes", "Arrows", "Callouts"];
const ICON_VIEWBOX = `0 0 ${SHAPE_ICON_SIZE} ${SHAPE_ICON_SIZE}`;

function itemsFor(group = "") {
  return SHAPE_CATALOGUE.filter((item) => item.group === group);
}

function iconPath(geometry = "rect") {
  return shapeIconPath(asShapeGeometry(geometry));
}

/** A line geometry's icon is an open segment (plus a filled arrowhead), so
 * it strokes rather than fills like the closed outlines do. */
function iconClass(geometry = "rect") {
  return LINE_GEOMETRIES.has(asShapeGeometry(geometry)) ? "shape-gallery-icon-line" : "shape-gallery-icon-area";
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
          :aria-label="item.label"
          @mousedown.prevent
          @click="$emit('pick', item.geometry)"
        >
          <svg
            class="shape-gallery-icon"
            :viewBox="ICON_VIEWBOX"
            aria-hidden="true"
            focusable="false"
          >
            <path
              :class="iconClass(item.geometry)"
              :d="iconPath(item.geometry)"
            />
          </svg>
        </button>
      </div>
    </div>
  </div>
</template>

<style scoped>
.shape-gallery {
  display: flex;
  flex-direction: column;
  gap: var(--space-1);
  /* The 1rem side padding is MenuBar's own command gutter, so the icon grid
     starts on the same vertical line as the menu's labels above and below. */
  padding: var(--space-1) var(--space-5) var(--space-2);
}

/* Each group is a label beside its own row rather than above it: the Insert
   menu is by far the longest in the bar and has to fit a 720px viewport
   whole (SPEC-028 AC-8/AC-9 asserts the break commands below this gallery
   are inside it), and the menu has horizontal room to spare where it has
   none vertically. */
.shape-gallery-group {
  display: flex;
  align-items: center;
  gap: var(--space-3);
}

.shape-gallery-group-label {
  flex: none;
  width: 4rem;
  font-size: var(--text-xs);
  font-weight: var(--weight-bold);
  line-height: var(--line-tight);
  letter-spacing: 0.04em;
  text-transform: uppercase;
  color: var(--text-secondary);
}

.shape-gallery-items {
  flex: 1;
  display: grid;
  /* Fixed track count with equal shares: every cell holds the same
     fixed-size icon, so the columns line up across all four groups. */
  grid-template-columns: repeat(6, 1fr);
  gap: var(--space-2);
}

/* A fixed row height (rather than a square cell) keeps the four groups
   from adding up to more menu than the names they replaced did — the
   Insert menu is already the longest one in the bar. */
.shape-gallery-item {
  display: flex;
  align-items: center;
  justify-content: center;
  height: 1.5rem;
  padding: var(--space-1);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  cursor: pointer;
}

.shape-gallery-item:hover {
  background: var(--surface-sunken);
  border-color: var(--accent);
}

/* Focus keeps the accent border AND a full 2px --focus-ring (AC-45/46). */
.shape-gallery-item:focus-visible {
  background: var(--surface-sunken);
  border-color: var(--accent);
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: var(--space-1);
}

.shape-gallery-icon {
  display: block;
  width: 1.25rem;
  height: 1.25rem;
}

.shape-gallery-icon-area {
  fill: var(--surface-selected);
  stroke: var(--text-primary);
  stroke-width: 1.25;
  stroke-linejoin: round;
}

.shape-gallery-icon-line {
  fill: var(--text-primary);
  stroke: var(--text-primary);
  stroke-width: 1.5;
  stroke-linecap: round;
  stroke-linejoin: round;
}
</style>
