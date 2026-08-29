<script setup lang="ts">
// "Insert -> Shape"'s gallery (SPEC-051 AC-1/AC-2, SPEC-069 AC-2): the whole
// expanded catalogue, grouped as in "The expanded catalogue" and made
// scrollable, with a search field that filters by label and alias. Picking
// one arms draw mode (AC-3) — this component only reports the pick,
// DocumentView.vue's useShapes owns what happens next.
//
// Each geometry shows as its own outline (shapeIcons.ts draws it from the
// same vertex lists the page canvas paints with), not its name: a fixed-size
// icon makes every cell the same width. Names that no longer show stay
// reachable as the button's tooltip, accessible name and search term.
//
// The buttons deliberately carry no menuitem role (see MenuBar.vue). Runtime
// prop/emit declarations per the espree constraint (see AppShell.vue).
import { computed, ref } from "vue";
import { asShapeGeometry, LINE_GEOMETRIES } from "../lib/core/shapes";
import { SHAPE_DEFINITIONS } from "../lib/core/shapeCatalogue";
import { SHAPE_ICON_SIZE, shapeIconPath } from "../lib/core/shapeIcons";

defineEmits(["pick", "pick-connector", "pick-freeform"]);

// SPEC-071: the three freeform drawing tools, shown in their own "Freeform"
// group. Like connectors these are not a ShapeGeometry (they produce a
// path-bearing shape), so they are listed explicitly.
const FREEFORM_TOOLS = [
  { tool: "freeform", label: "Freeform", d: "M2 20 L8 4 L14 16 L22 6" },
  { tool: "curve", label: "Curve", d: "M2 18 C6 2 16 2 22 18" },
  { tool: "scribble", label: "Scribble", d: "M2 12 C4 4 8 20 12 12 C16 4 20 20 22 12" },
];

// SPEC-070: the three connector kinds shown in their own "Connectors" group.
// Connectors are a distinct block type (not a ShapeGeometry), so they are
// listed explicitly rather than derived from SHAPE_DEFINITIONS.
const CONNECTOR_KINDS = [
  { kind: "straight", label: "Straight connector", d: "M2 2 L22 22" },
  { kind: "elbow", label: "Elbow connector", d: "M2 2 L2 12 L22 12 L22 22" },
  { kind: "curved", label: "Curved connector", d: "M2 2 C2 16 22 8 22 22" },
];

// Groups in "The expanded catalogue" order. No type annotation here (the
// .vue script's espree parser bars TS syntax — see the header); ShapeGroup
// is enforced on the definitions in shapeCatalogue.ts instead.
const GROUPS = ["Lines", "Basic shapes", "Block arrows", "Flowchart", "Stars and banners", "Callouts"];
const ICON_VIEWBOX = `0 0 ${SHAPE_ICON_SIZE} ${SHAPE_ICON_SIZE}`;

const query = ref("");

/** The catalogue split by group and filtered by the search box: a geometry
 * matches when its label, key or any alias contains the query (AC-2b);
 * groups with no match are dropped. */
const filteredGroups = computed(() => {
  const term = query.value.trim().toLowerCase();
  return GROUPS.map((group) => ({
    group,
    items: SHAPE_DEFINITIONS.filter(
      (def) =>
        def.group === group &&
        (!term || [def.label, def.geometry, ...(def.aliases ?? [])].join(" ").toLowerCase().includes(term)),
    ),
  })).filter((entry) => entry.items.length > 0);
});

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
    <div class="shape-gallery-search">
      <input
        v-model="query"
        type="search"
        class="shape-gallery-search-input"
        data-testid="shape-gallery-search"
        placeholder="Search shapes"
        aria-label="Search shapes"
        @mousedown.stop
        @keydown.stop
        @click.stop
      >
    </div>
    <div class="shape-gallery-scroll">
      <div
        v-for="entry in filteredGroups"
        :key="entry.group"
        class="shape-gallery-group"
      >
        <div class="shape-gallery-group-label">
          {{ entry.group }}
        </div>
        <div class="shape-gallery-items">
          <button
            v-for="item in entry.items"
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
      <div
        v-if="!query.trim()"
        class="shape-gallery-group"
        data-testid="connector-gallery-group"
      >
        <div class="shape-gallery-group-label">
          Connectors
        </div>
        <div class="shape-gallery-items">
          <button
            v-for="conn in CONNECTOR_KINDS"
            :key="conn.kind"
            type="button"
            class="shape-gallery-item"
            :data-testid="`connector-gallery-${conn.kind}`"
            :title="conn.label"
            :aria-label="conn.label"
            @mousedown.prevent
            @click="$emit('pick-connector', conn.kind)"
          >
            <svg
              class="shape-gallery-icon"
              :viewBox="ICON_VIEWBOX"
              aria-hidden="true"
              focusable="false"
            >
              <path
                class="shape-gallery-icon-line"
                fill="none"
                :d="conn.d"
              />
            </svg>
          </button>
        </div>
      </div>
      <div
        v-if="!query.trim()"
        class="shape-gallery-group"
        data-testid="freeform-gallery-group"
      >
        <div class="shape-gallery-group-label">
          Freeform
        </div>
        <div class="shape-gallery-items">
          <button
            v-for="tool in FREEFORM_TOOLS"
            :key="tool.tool"
            type="button"
            class="shape-gallery-item"
            :data-testid="`freeform-gallery-${tool.tool}`"
            :title="tool.label"
            :aria-label="tool.label"
            @mousedown.prevent
            @click="$emit('pick-freeform', tool.tool)"
          >
            <svg
              class="shape-gallery-icon"
              :viewBox="ICON_VIEWBOX"
              aria-hidden="true"
              focusable="false"
            >
              <path
                class="shape-gallery-icon-line"
                fill="none"
                :d="tool.d"
              />
            </svg>
          </button>
        </div>
      </div>
      <div
        v-if="filteredGroups.length === 0"
        class="shape-gallery-empty"
        data-testid="shape-gallery-empty"
      >
        No shapes match “{{ query }}”
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

.shape-gallery-search {
  padding-bottom: var(--space-1);
}

.shape-gallery-search-input {
  width: 100%;
  box-sizing: border-box;
  padding: var(--space-1) var(--space-2);
  font-size: var(--text-sm);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  color: var(--text-primary);
}

.shape-gallery-search-input:focus-visible {
  border-color: var(--accent);
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: 0;
}

/* The whole catalogue no longer fits the Insert menu vertically, so the
   groups scroll inside a bounded box while the menu itself stays whole. The
   cap is tight on purpose: SPEC-028 AC-8/AC-9 asserts all four break commands
   below this gallery lay out inside a 720px viewport, and everything above
   them (the table grid, this gallery, the chart gallery) has to leave room.
   The rem cap wins over 40vh at the 720px height the e2e drives, so the box
   stays small there; 40vh only takes over on a viewport too short for the
   rem cap to fit. */
.shape-gallery-scroll {
  display: flex;
  flex-direction: column;
  gap: var(--space-1);
  max-height: min(40vh, 8.5rem);
  overflow-y: auto;
}

/* Each group is a label beside its own row rather than above it: the Insert
   menu has horizontal room to spare where it has none vertically. */
.shape-gallery-group {
  display: flex;
  align-items: flex-start;
  gap: var(--space-3);
}

.shape-gallery-group-label {
  flex: none;
  width: 4rem;
  padding-top: var(--space-1);
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
     fixed-size icon, so the columns line up across all the groups. */
  grid-template-columns: repeat(6, 1fr);
  gap: var(--space-2);
}

/* A fixed row height (rather than a square cell) keeps the groups compact. */
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

.shape-gallery-empty {
  padding: var(--space-2) 0;
  font-size: var(--text-sm);
  color: var(--text-secondary);
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
