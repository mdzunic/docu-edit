<script setup lang="ts">
import { computed, onBeforeUnmount, onMounted, ref, useTemplateRef } from "vue";
import AppIcon from "./AppIcon.vue";
import { asSidePanelTabs } from "../views/document/useSidePanel";

// The consolidated side panel chrome (SPEC-065): one docked, tabbed,
// resizable region hosting every auxiliary pane. It draws the header (active
// title, a pane-supplied #header-actions slot, a close button), the tab rail,
// the drag-to-resize handle and the tabpanel body; DocumentView renders the
// pane bodies into the default slot. All state (open tabs, active tab, side,
// width, collapse) is owned by useSidePanel — this component only reflects it
// and emits intent. Runtime prop/emit declarations and default-valued handler
// params, not the generic/annotated form — espree parses .vue scripts and
// rejects TS-only syntax (see CommentsPane.vue / OutlinePane.vue).
const props = defineProps({
  side: { type: String, default: "right" },
  widthPx: { type: Number, default: 320 },
  collapsed: { type: Boolean, default: false },
  tabs: { type: Array, default: () => [] },
  activePane: { type: String, default: null },
});

const emit = defineEmits([
  "activate",
  "close",
  "toggle-side",
  "toggle-collapse",
  "update:widthPx",
  "commit-width",
  "nudge-width",
]);

const PANEL_ID = "side-panel-tabpanel";
function tabDomId(id = "") {
  return `side-panel-tab-${id}`;
}

// The runtime Array prop carries no element type (the espree constraint bars
// PropType<...> here) — this cast gives the template typed tabs.
const tabList = computed(() => asSidePanelTabs(props.tabs));

const activeTitle = computed(() => {
  const active = tabList.value.find((tab) => tab.id === props.activePane);
  return active ? active.title : "";
});

// --- responsive overlay (AC-25/26): below 900px the panel overlays the
// document with a scrim instead of displacing it. -----------------------------
const viewportWidth = ref(typeof window === "undefined" ? 1024 : window.innerWidth);
const overlay = computed(() => viewportWidth.value < 900);
function onViewportResize() {
  viewportWidth.value = window.innerWidth;
}
function onWindowKeydown(event = new KeyboardEvent("keydown")) {
  // Escape closes the overlay panel (AC-25b).
  if (overlay.value && event.key === "Escape") {
    emit("close");
  }
}
onMounted(() => {
  window.addEventListener("resize", onViewportResize);
  window.addEventListener("keydown", onWindowKeydown);
});
onBeforeUnmount(() => {
  window.removeEventListener("resize", onViewportResize);
  window.removeEventListener("keydown", onWindowKeydown);
});

const panelStyle = computed(() => (props.collapsed ? {} : { width: `${props.widthPx}px` }));

// --- tab rail keyboard navigation (AC-19) ------------------------------------
const tablistRef = useTemplateRef("tablistRef");
function focusTabAt(index = 0) {
  const buttons = tablistRef.value?.querySelectorAll('[role="tab"]');
  const button = buttons ? buttons[index] : null;
  if (button instanceof HTMLElement) {
    button.focus();
  }
}
function onTabKeydown(event = new KeyboardEvent("keydown"), index = 0) {
  const count = tabList.value.length;
  if (count === 0) {
    return;
  }
  let target = -1;
  switch (event.key) {
    case "ArrowRight":
    case "ArrowDown":
      target = (index + 1) % count;
      break;
    case "ArrowLeft":
    case "ArrowUp":
      target = (index - 1 + count) % count;
      break;
    case "Home":
      target = 0;
      break;
    case "End":
      target = count - 1;
      break;
    case "Enter":
    case " ":
      event.preventDefault();
      emit("activate", tabList.value[index].id);
      return;
    default:
      return;
  }
  event.preventDefault();
  focusTabAt(target);
}

// --- resize handle (AC-8, AC-20) ---------------------------------------------
let dragging = false;
function onHandlePointerdown(event = new PointerEvent("pointerdown")) {
  dragging = true;
  if (event.target instanceof HTMLElement) {
    event.target.setPointerCapture(event.pointerId);
  }
  event.preventDefault();
}
function onHandlePointermove(event = new PointerEvent("pointermove")) {
  if (!dragging) {
    return;
  }
  // The handle sits on the panel's inner edge (next to the document). For a
  // right dock that edge is the panel's left, so the width is the distance
  // from the pointer to the window's right; mirror it for a left dock.
  const width = props.side === "right" ? window.innerWidth - event.clientX : event.clientX;
  emit("update:widthPx", width);
}
function onHandlePointerup(event = new PointerEvent("pointerup")) {
  if (!dragging) {
    return;
  }
  dragging = false;
  if (event.target instanceof HTMLElement) {
    event.target.releasePointerCapture(event.pointerId);
  }
  emit("commit-width");
}
function onHandleKeydown(event = new KeyboardEvent("keydown")) {
  const STEP = 8;
  const widenKey = props.side === "right" ? "ArrowLeft" : "ArrowRight";
  const narrowKey = props.side === "right" ? "ArrowRight" : "ArrowLeft";
  let delta = 0;
  if (event.key === widenKey || event.key === "ArrowUp") {
    delta = STEP;
  } else if (event.key === narrowKey || event.key === "ArrowDown") {
    delta = -STEP;
  } else {
    return;
  }
  event.preventDefault();
  emit("nudge-width", delta);
}
</script>

<template>
  <div
    class="side-panel-root"
    :class="[`dock-${side}`, { overlay, collapsed }]"
  >
    <div
      v-if="overlay"
      class="side-panel-scrim"
      data-testid="side-panel-scrim"
      @click="emit('close')"
    />
    <aside
      class="side-panel print-exclude"
      :class="[`dock-${side}`, { collapsed }]"
      :style="panelStyle"
      role="complementary"
      :aria-label="activeTitle || 'Side panel'"
      data-testid="side-panel"
    >
      <div class="side-panel-inner">
        <header class="side-panel-header">
          <button
            type="button"
            class="side-panel-icon-button"
            :aria-label="side === 'right' ? 'Move panel to left' : 'Move panel to right'"
            data-testid="side-panel-side-toggle"
            @click="emit('toggle-side')"
          >
            {{ side === "right" ? "⇤" : "⇥" }}
          </button>
          <span class="side-panel-title">{{ activeTitle }}</span>
          <div class="side-panel-header-actions">
            <slot name="header-actions" />
          </div>
          <button
            type="button"
            class="side-panel-icon-button side-panel-close"
            aria-label="Close panel"
            data-testid="side-panel-close"
            @click="emit('close')"
          >
            ×
          </button>
        </header>
        <div
          ref="tablistRef"
          class="side-panel-tabs"
          role="tablist"
          aria-label="Panel tabs"
        >
          <button
            v-for="(tab, index) in tabList"
            :id="tabDomId(tab.id)"
            :key="tab.id"
            type="button"
            class="side-panel-tab"
            :class="{ active: tab.id === activePane }"
            role="tab"
            :aria-selected="tab.id === activePane"
            :aria-controls="PANEL_ID"
            :tabindex="tab.id === activePane ? 0 : -1"
            :data-testid="`side-panel-tab-${tab.id}`"
            @click="emit('activate', tab.id)"
            @keydown="onTabKeydown($event, index)"
          >
            <AppIcon
              class="side-panel-tab-icon"
              :name="tab.icon"
            />
            <span class="side-panel-tab-label">{{ tab.title }}</span>
            <span
              v-if="tab.badge > 0"
              class="side-panel-tab-badge"
              :data-testid="`side-panel-badge-${tab.id}`"
            >{{ tab.badge }}</span>
          </button>
        </div>
        <div
          v-show="!collapsed"
          :id="PANEL_ID"
          class="side-panel-body"
          role="tabpanel"
          :aria-labelledby="activePane ? tabDomId(activePane) : undefined"
        >
          <slot />
        </div>
      </div>
      <!-- Last in the DOM so focus runs document -> tab rail -> pane contents
           (AC-21) before reaching this chrome; absolutely positioned, so its
           document order does not affect layout. -->
      <div
        class="side-panel-resize"
        role="separator"
        aria-orientation="vertical"
        aria-label="Resize panel"
        tabindex="0"
        data-testid="side-panel-resize"
        @pointerdown="onHandlePointerdown"
        @pointermove="onHandlePointermove"
        @pointerup="onHandlePointerup"
        @dblclick="emit('toggle-collapse')"
        @keydown="onHandleKeydown"
      />
    </aside>
  </div>
</template>

<style scoped>
/* The panel is a flex column sibling of the document (AC-26 ≥900px: it
   displaces the document). page-wrap stays first in the DOM so focus runs
   document -> panel (AC-21); a left dock only reorders it visually. */
.side-panel-root {
  flex: 0 0 auto;
  display: flex;
  align-self: stretch;
}

.side-panel-root.dock-left {
  order: -1;
}

.side-panel {
  display: flex;
  box-sizing: border-box;
  background: var(--surface-base);
  z-index: 20;
}

/* The visible chrome pins to the top of the scroll viewport while the pages
   scroll beneath, mirroring the old per-pane docks. */
.side-panel-inner {
  position: sticky;
  top: 0;
  display: flex;
  flex-direction: column;
  flex: 1;
  min-width: 0;
  max-height: 100vh;
  overflow: hidden;
}

.side-panel.dock-right {
  border-left: var(--border-width) solid var(--border-default);
}

.side-panel.dock-left {
  border-right: var(--border-width) solid var(--border-default);
}

/* Collapsed: the tab rail only, the body hidden (AC-10). */
.side-panel.collapsed {
  width: auto !important;
}

.side-panel.collapsed .side-panel-header {
  display: none;
}

.side-panel-resize {
  position: absolute;
  top: 0;
  bottom: 0;
  width: 6px;
  cursor: ew-resize;
  z-index: 21;
  touch-action: none;
}

.side-panel.dock-right .side-panel-resize {
  left: -3px;
}

.side-panel.dock-left .side-panel-resize {
  right: -3px;
}

.side-panel-resize:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: -1px;
}

.side-panel-header {
  display: flex;
  align-items: center;
  gap: var(--space-2);
  padding: var(--space-2) var(--space-3);
  border-bottom: var(--border-width) solid var(--border-subtle);
}

.side-panel-title {
  flex: 1;
  min-width: 0;
  font-size: var(--text-md);
  font-weight: var(--weight-bold);
  color: var(--text-primary);
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.side-panel-header-actions {
  display: flex;
  align-items: center;
  gap: var(--space-1);
}

.side-panel-icon-button {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  min-width: 24px;
  height: 24px;
  padding: 0 var(--space-1);
  border: none;
  border-radius: var(--radius-sm);
  background: transparent;
  color: var(--text-secondary);
  font-size: var(--text-lg);
  line-height: 1;
  cursor: pointer;
}

.side-panel-icon-button:hover {
  background: var(--surface-sunken);
  color: var(--text-primary);
}

/* Tab rail: a horizontal row of tabs; wraps if there are many open. */
.side-panel-tabs {
  display: flex;
  flex-wrap: wrap;
  gap: var(--space-1);
  padding: var(--space-2);
  border-bottom: var(--border-width) solid var(--border-subtle);
  background: var(--surface-raised);
}

.side-panel-tab {
  display: inline-flex;
  align-items: center;
  gap: var(--space-2);
  padding: var(--space-2) var(--space-3);
  border: var(--border-width) solid transparent;
  border-radius: var(--radius-sm);
  background: transparent;
  color: var(--text-secondary);
  font-size: var(--text-sm);
  cursor: pointer;
}

.side-panel-tab:hover {
  background: var(--surface-sunken);
  color: var(--text-primary);
}

.side-panel-tab.active {
  background: var(--surface-selected);
  color: var(--text-accent);
  border-color: var(--accent-subtle);
}

.side-panel-tab-icon {
  flex: 0 0 auto;
}

.side-panel-tab-label {
  white-space: nowrap;
}

.side-panel-tab-badge {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  min-width: 16px;
  height: 16px;
  padding: 0 var(--space-1);
  border-radius: var(--radius-pill);
  background: var(--accent);
  color: var(--text-on-brand);
  font-size: var(--text-xs);
  line-height: 1;
}

.side-panel-body {
  flex: 1;
  min-height: 0;
  overflow: auto;
}

/* Below 900px the panel overlays the document (AC-25): fixed to the viewport
   over a scrim, taking no layout space so the document does not reflow. */
.side-panel-root.overlay {
  flex: 0 0 0;
  width: 0;
}

.side-panel-scrim {
  position: fixed;
  inset: 0;
  background: var(--surface-overlay);
  z-index: 39;
}

.side-panel-root.overlay .side-panel {
  position: fixed;
  top: 0;
  bottom: 0;
  z-index: 40;
  box-shadow: var(--shadow-3);
}

.side-panel-root.overlay .side-panel.dock-right {
  right: 0;
}

.side-panel-root.overlay .side-panel.dock-left {
  left: 0;
}

.side-panel-root.overlay .side-panel-inner {
  max-height: 100%;
  height: 100%;
}
</style>
