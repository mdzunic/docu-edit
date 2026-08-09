<script setup lang="ts">
import { computed, ref, watch } from "vue";
import {
  OUTLINE_EMPTY_TEXT,
  asKeySet,
  asOutlineNodes,
  noAnchor,
  noDropPosition,
  visibleOutlineRows,
} from "../lib/outline";

// The document outline navigator pane (SPEC-037). A single-tabstop tree
// (`aria-activedescendant`) inside a `<nav>` landmark (AC-20): rows are
// indented by depth (AC-2), the caret's heading is highlighted (AC-3),
// clicking or pressing Enter navigates (AC-5, AC-6, AC-19), arrow keys move
// the selection (AC-18), rows with children collapse/expand (AC-7, AC-8), and
// a row can be dragged among its siblings to restructure the document (AC-9).
// Runtime prop/emit declarations, not the generic form — espree parses .vue
// scripts and rejects TS-only syntax (see AppShell.vue).
const props = defineProps({
  /** The heading tree (built by lib/outline.buildOutline). */
  outline: { type: Array, default: () => [] },
  /** Anchors whose children are collapsed in the pane (a Set of strings). */
  collapsed: { type: Object, default: () => new Set() },
  /** The anchor of the heading containing the caret, highlighted (AC-3). */
  activeAnchor: { type: String, default: null },
});

const emit = defineEmits(["navigate", "toggle", "move"]);

const outlineNodes = computed(() => asOutlineNodes(props.outline));
const rows = computed(() => visibleOutlineRows(outlineNodes.value, asKeySet(props.collapsed)));

// Each anchor's parent anchor, so a drag can tell siblings apart (only
// same-level siblings may reorder — AC-10). Mirrors lib/outline's own rule;
// the committed move re-checks authoritatively via moveHeading.
const parentOf = computed(() => {
  const map = new Map();
  const walk = (nodes = asOutlineNodes([]), parent = noAnchor()) => {
    for (const node of nodes) {
      map.set(node.anchor, parent);
      walk(node.children, node.anchor);
    }
  };
  walk(outlineNodes.value, noAnchor());
  return map;
});

const levelOf = computed(() => {
  const map = new Map();
  const walk = (nodes = asOutlineNodes([])) => {
    for (const node of nodes) {
      map.set(node.anchor, node.level);
      walk(node.children);
    }
  };
  walk(outlineNodes.value);
  return map;
});

// --- keyboard selection (single roving active descendant) ------------------
const focusedAnchor = ref(noAnchor());

function itemId(anchor = "") {
  return `outline-item-${anchor}`;
}

// Keep the focused row valid as the visible set changes (collapse, edits): if
// it vanished, fall back to the active heading or the first row.
watch(
  [rows, () => props.activeAnchor],
  () => {
    const visible = rows.value.map((row) => row.node.anchor);
    if (focusedAnchor.value && visible.includes(focusedAnchor.value)) {
      return;
    }
    if (props.activeAnchor && visible.includes(props.activeAnchor)) {
      focusedAnchor.value = props.activeAnchor;
    } else {
      focusedAnchor.value = visible[0] ?? null;
    }
  },
  { immediate: true },
);

function focusedIndex() {
  return rows.value.findIndex((row) => row.node.anchor === focusedAnchor.value);
}

function moveFocus(delta = 0) {
  if (rows.value.length === 0) {
    return;
  }
  const current = focusedIndex();
  const next = current < 0 ? 0 : Math.min(rows.value.length - 1, Math.max(0, current + delta));
  focusedAnchor.value = rows.value[next].node.anchor;
}

function handleKeydown(event = new KeyboardEvent("keydown")) {
  const index = focusedIndex();
  const row = index >= 0 ? rows.value[index] : null;
  switch (event.key) {
    case "ArrowDown":
      event.preventDefault();
      moveFocus(1);
      break;
    case "ArrowUp":
      event.preventDefault();
      moveFocus(-1);
      break;
    case "Home":
      event.preventDefault();
      if (rows.value.length > 0) {
        focusedAnchor.value = rows.value[0].node.anchor;
      }
      break;
    case "End":
      event.preventDefault();
      if (rows.value.length > 0) {
        focusedAnchor.value = rows.value[rows.value.length - 1].node.anchor;
      }
      break;
    case "ArrowRight":
      if (row?.hasChildren && row.collapsed) {
        event.preventDefault();
        emit("toggle", row.node.anchor);
      }
      break;
    case "ArrowLeft":
      if (row?.hasChildren && !row.collapsed) {
        event.preventDefault();
        emit("toggle", row.node.anchor);
      }
      break;
    case "Enter":
    case " ":
      if (focusedAnchor.value) {
        event.preventDefault();
        emit("navigate", focusedAnchor.value);
      }
      break;
    default:
      break;
  }
}

function onRowClick(anchor = "") {
  focusedAnchor.value = anchor;
  emit("navigate", anchor);
}

function onToggle(anchor = "") {
  emit("toggle", anchor);
}

// --- drag to reorder among siblings (AC-9, AC-10, AC-11) -------------------
const draggingAnchor = ref(noAnchor());
const dropTargetAnchor = ref(noAnchor());
const dropPosition = ref(noDropPosition());

function canDrop(targetAnchor = "", position = noDropPosition()) {
  const source = draggingAnchor.value;
  if (!source || position === null || source === targetAnchor) {
    return false;
  }
  // Same level and same parent — the sibling rule (AC-10, AC-11).
  return (
    levelOf.value.get(source) === levelOf.value.get(targetAnchor) &&
    parentOf.value.get(source) === parentOf.value.get(targetAnchor)
  );
}

function onDragStart(anchor = "", event = new DragEvent("dragstart")) {
  draggingAnchor.value = anchor;
  if (event.dataTransfer) {
    event.dataTransfer.effectAllowed = "move";
    // Some browsers require data to be set for a drag to start.
    event.dataTransfer.setData("text/plain", anchor);
  }
}

// The drop lands "before" in the top third and "after" in the bottom third;
// the middle third is a drop ONTO the heading itself, which is rejected (AC-11).
function zoneFor(event = new DragEvent("dragover")) {
  const target = event.currentTarget;
  if (!(target instanceof HTMLElement)) {
    return null;
  }
  const rect = target.getBoundingClientRect();
  const ratio = rect.height === 0 ? 0.5 : (event.clientY - rect.top) / rect.height;
  if (ratio < 0.34) {
    return "before";
  }
  if (ratio > 0.66) {
    return "after";
  }
  return null;
}

function onDragOver(anchor = "", event = new DragEvent("dragover")) {
  const position = zoneFor(event);
  if (canDrop(anchor, position)) {
    event.preventDefault();
    dropTargetAnchor.value = anchor;
    dropPosition.value = position;
    if (event.dataTransfer) {
      event.dataTransfer.dropEffect = "move";
    }
  } else {
    dropTargetAnchor.value = null;
    dropPosition.value = null;
  }
}

function onDragLeave(anchor = "") {
  if (dropTargetAnchor.value === anchor) {
    dropTargetAnchor.value = null;
    dropPosition.value = null;
  }
}

function onDrop(anchor = "", event = new DragEvent("drop")) {
  const position = zoneFor(event);
  if (canDrop(anchor, position) && draggingAnchor.value) {
    event.preventDefault();
    emit("move", { source: draggingAnchor.value, target: anchor, position });
  }
  onDragEnd();
}

function onDragEnd() {
  draggingAnchor.value = null;
  dropTargetAnchor.value = null;
  dropPosition.value = null;
}
</script>

<template>
  <nav
    class="outline-pane"
    aria-label="Document outline"
    data-testid="outline-pane"
  >
    <div class="outline-heading">
      Outline
    </div>
    <p
      v-if="rows.length === 0"
      class="outline-empty"
      data-testid="outline-empty"
    >
      {{ OUTLINE_EMPTY_TEXT }}
    </p>
    <div
      v-else
      class="outline-tree"
      role="tree"
      aria-label="Document outline"
      tabindex="0"
      data-testid="outline-tree"
      :aria-activedescendant="focusedAnchor ? itemId(focusedAnchor) : undefined"
      @keydown="handleKeydown"
    >
      <div
        v-for="row in rows"
        :id="itemId(row.node.anchor)"
        :key="row.node.anchor"
        class="outline-item"
        :class="{
          'is-active': row.node.anchor === activeAnchor,
          'is-focused': row.node.anchor === focusedAnchor,
          'drop-before': dropTargetAnchor === row.node.anchor && dropPosition === 'before',
          'drop-after': dropTargetAnchor === row.node.anchor && dropPosition === 'after',
        }"
        role="treeitem"
        :aria-level="row.depth + 1"
        :aria-selected="row.node.anchor === focusedAnchor ? 'true' : 'false'"
        :aria-current="row.node.anchor === activeAnchor ? 'true' : undefined"
        :aria-expanded="row.hasChildren ? (row.collapsed ? 'false' : 'true') : undefined"
        :style="{ paddingLeft: `${row.depth * 1 + 0.5}rem` }"
        :data-testid="`outline-item-${row.node.anchor}`"
        draggable="true"
        @click="onRowClick(row.node.anchor)"
        @dragstart="onDragStart(row.node.anchor, $event)"
        @dragover="onDragOver(row.node.anchor, $event)"
        @dragleave="onDragLeave(row.node.anchor)"
        @drop="onDrop(row.node.anchor, $event)"
        @dragend="onDragEnd"
      >
        <button
          v-if="row.hasChildren"
          type="button"
          class="outline-twisty"
          tabindex="-1"
          :aria-label="row.collapsed ? 'Expand' : 'Collapse'"
          :data-testid="`outline-twisty-${row.node.anchor}`"
          @click.stop="onToggle(row.node.anchor)"
          @mousedown.stop
        >
          {{ row.collapsed ? "▸" : "▾" }}
        </button>
        <span
          v-else
          class="outline-twisty-spacer"
          aria-hidden="true"
        />
        <span class="outline-label">{{ row.node.text || "(untitled heading)" }}</span>
      </div>
    </div>
  </nav>
</template>

<style scoped>
.outline-pane {
  flex: none;
  align-self: flex-start;
  position: sticky;
  top: 0;
  width: 16rem;
  max-height: 100vh;
  display: flex;
  flex-direction: column;
  min-height: 0;
  overflow-y: auto;
  background: #f8f9fa;
  border-right: 1px solid #dadce0;
  box-sizing: border-box;
}

.outline-heading {
  flex: none;
  padding: 0.5rem 0.75rem;
  font-size: 0.75rem;
  font-weight: 600;
  letter-spacing: 0.04em;
  text-transform: uppercase;
  color: #5f6368;
}

.outline-empty {
  margin: 0;
  padding: 0.25rem 0.75rem 0.75rem;
  font-size: 0.8125rem;
  font-style: italic;
  color: #5f6368;
}

.outline-tree:focus {
  outline: none;
}

.outline-item {
  position: relative;
  display: flex;
  align-items: center;
  gap: 0.25rem;
  padding-top: 0.2rem;
  padding-bottom: 0.2rem;
  padding-right: 0.5rem;
  font-size: 0.8125rem;
  color: #202124;
  cursor: pointer;
  user-select: none;
}

.outline-item:hover {
  background: #f1f3f4;
}

.outline-item.is-active {
  background: #e8f0fe;
  color: #1a73e8;
  font-weight: 600;
}

.outline-tree:focus .outline-item.is-focused {
  box-shadow: inset 2px 0 0 #1a73e8;
}

.outline-item.drop-before::before,
.outline-item.drop-after::after {
  content: "";
  position: absolute;
  left: 0;
  right: 0;
  height: 2px;
  background: #1a73e8;
}

.outline-item.drop-before::before {
  top: 0;
}

.outline-item.drop-after::after {
  bottom: 0;
}

.outline-twisty {
  flex: none;
  width: 1.1rem;
  height: 1.1rem;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  border: 0;
  background: transparent;
  color: #5f6368;
  font-size: 0.7rem;
  line-height: 1;
  cursor: pointer;
  border-radius: 2px;
}

.outline-twisty:hover {
  background: #e0e0e0;
}

.outline-twisty-spacer {
  flex: none;
  width: 1.1rem;
}

.outline-label {
  overflow: hidden;
  white-space: nowrap;
  text-overflow: ellipsis;
}
</style>
