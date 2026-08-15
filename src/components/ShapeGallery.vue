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
import { asShapeGeometry, LINE_GEOMETRIES, SHAPE_CATALOGUE } from "../lib/documentModel";
import { SHAPE_ICON_SIZE, shapeIconPath } from "../lib/shapeIcons";

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
  gap: 0.1875rem;
  /* The 1rem side padding is MenuBar's own command gutter, so the icon grid
     starts on the same vertical line as the menu's labels above and below. */
  padding: 0.125rem 1rem 0.25rem;
}

/* Each group is a label beside its own row rather than above it: the Insert
   menu is by far the longest in the bar and has to fit a 720px viewport
   whole (SPEC-028 AC-8/AC-9 asserts the break commands below this gallery
   are inside it), and the menu has horizontal room to spare where it has
   none vertically. */
.shape-gallery-group {
  display: flex;
  align-items: center;
  gap: 0.5rem;
}

.shape-gallery-group-label {
  flex: none;
  width: 4rem;
  font-size: 0.625rem;
  font-weight: 600;
  line-height: 1.15;
  letter-spacing: 0.04em;
  text-transform: uppercase;
  color: #80868b;
}

.shape-gallery-items {
  flex: 1;
  display: grid;
  /* Fixed track count with equal shares: every cell holds the same
     fixed-size icon, so the columns line up across all four groups. */
  grid-template-columns: repeat(6, 1fr);
  gap: 0.25rem;
}

/* A fixed row height (rather than a square cell) keeps the four groups
   from adding up to more menu than the names they replaced did — the
   Insert menu is already the longest one in the bar. */
.shape-gallery-item {
  display: flex;
  align-items: center;
  justify-content: center;
  height: 1.5rem;
  padding: 0.125rem;
  border: 1px solid #dadce0;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
}

.shape-gallery-item:hover,
.shape-gallery-item:focus-visible {
  background: #f1f3f4;
  border-color: #1a73e8;
  outline: none;
}

.shape-gallery-icon {
  display: block;
  width: 1.25rem;
  height: 1.25rem;
}

.shape-gallery-icon-area {
  fill: #e8f0fe;
  stroke: #3c4043;
  stroke-width: 1.25;
  stroke-linejoin: round;
}

.shape-gallery-icon-line {
  fill: #3c4043;
  stroke: #3c4043;
  stroke-width: 1.5;
  stroke-linecap: round;
  stroke-linejoin: round;
}
</style>
