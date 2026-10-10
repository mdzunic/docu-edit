<script setup lang="ts">
// A sheet tab's context menu (SPEC-121 "UI"): Insert sheet, Rename,
// Duplicate, Delete, Move left/right, Hide, Unhide…, Tab colour — built from
// the shared MenuItem, so an item that cannot apply is disabled with its
// reason as a title and inline, never hidden. Fixed at a viewport point and
// kept on screen (flipped above the point at the foot of the view), over a
// backdrop that dismisses it. Keyboard: focus starts on the first enabled
// item; ArrowDown/ArrowUp (wrapping), Home and End move among the enabled
// ones; Escape and Tab close. Emits the chosen action; the view's
// useSheetCommands acts on it. Runtime prop declarations per the espree
// constraint (see AppShell.vue).
import { computed, nextTick, onMounted, onUnmounted, ref, useTemplateRef, watch } from "vue";
import MenuItem from "../MenuItem.vue";
import { asSheetTabMenuItems } from "../../views/spreadsheet/sheetTabMenu";
import { clampToViewport, enabledMenuItems } from "./sheetTabBar";

// Two root nodes (backdrop + menu): fall-through attributes go on the menu.
defineOptions({ inheritAttrs: false });

const props = defineProps({
  /** `{ action, label, testid, disabled?, reason? }` entries, in menu order. */
  items: { type: Array, required: true },
  /** The viewport point the menu opens at. */
  x: { type: Number, default: 0 },
  y: { type: Number, default: 0 },
});

const emit = defineEmits(["run", "close"]);

const entries = computed(() => asSheetTabMenuItems(props.items));
const menu = useTemplateRef("menu");
const position = ref({ x: props.x, y: props.y });

// Where focus was before the menu took it (read during setup, as BaseDialog
// does): a keyboard-opened menu hands focus back to its tab when it closes.
const previouslyFocused = document.activeElement instanceof HTMLElement ? document.activeElement : null;

function place() {
  const element = menu.value;
  if (!element) {
    return;
  }
  const rect = element.getBoundingClientRect();
  position.value = clampToViewport(
    { x: props.x, y: props.y },
    { width: rect.width, height: rect.height },
    { width: window.innerWidth, height: window.innerHeight },
  );
}

onMounted(async () => {
  place();
  await nextTick();
  const first = enabledMenuItems(menu.value)[0];
  if (first) {
    first.focus();
  } else {
    menu.value?.focus();
  }
});

watch(() => [props.x, props.y], place, { flush: "post" });

onUnmounted(() => {
  const active = document.activeElement;
  if (previouslyFocused?.isConnected && (active === null || active === document.body)) {
    previouslyFocused.focus();
  }
});

function onKeydown(event = new KeyboardEvent("keydown")) {
  if (event.key === "Escape") {
    event.preventDefault();
    event.stopPropagation();
    emit("close");
    return;
  }
  if (event.key === "Tab") {
    event.preventDefault();
    emit("close");
    return;
  }
  const items = enabledMenuItems(menu.value);
  const last = items.length - 1;
  const current = items.findIndex((item) => item === document.activeElement);
  const targets = {
    ArrowDown: current < 0 || current === last ? 0 : current + 1,
    ArrowUp: current <= 0 ? last : current - 1,
    Home: 0,
    End: last,
  };
  const target = Object.entries(targets).find(([key]) => key === event.key)?.[1];
  if (target === undefined) {
    return;
  }
  event.preventDefault();
  items[target]?.focus();
}
</script>

<template>
  <div
    class="sheet-tab-menu-backdrop"
    data-testid="sheet-tab-menu-backdrop"
    @click="emit('close')"
    @contextmenu.prevent="emit('close')"
  />
  <div
    ref="menu"
    v-bind="$attrs"
    class="sheet-tab-menu"
    role="menu"
    aria-label="Sheet options"
    data-testid="sheet-tab-menu"
    tabindex="-1"
    :style="{ left: `${position.x}px`, top: `${position.y}px` }"
    @keydown="onKeydown"
    @contextmenu.prevent
  >
    <template
      v-for="item in entries"
      :key="item.action"
    >
      <hr
        v-if="item.separatorBefore"
        class="menu-separator"
      >
      <MenuItem
        :label="item.label"
        :testid="item.testid"
        :disabled="item.disabled"
        :reason="item.reason"
        :danger="item.danger"
        @select="emit('run', item.action)"
      />
    </template>
  </div>
</template>

<style scoped>
.sheet-tab-menu-backdrop {
  position: fixed;
  inset: 0;
  z-index: 60;
}

/* MenuBar's popup look (the rows are MenuItems, styled through :deep as the
   bar styles its slotted items). */
.sheet-tab-menu {
  position: fixed;
  z-index: 61;
  display: flex;
  flex-direction: column;
  min-width: 15rem;
  padding: var(--space-1) var(--space-2);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
  background: var(--surface-base);
  box-shadow: var(--shadow-3);
}

.sheet-tab-menu :deep([role="menuitem"]) {
  display: flex;
  align-items: center;
  gap: var(--space-3);
  width: 100%;
  min-height: calc(var(--height-control) - var(--space-2));
  box-sizing: border-box;
  padding: 0 var(--space-3);
  border: 0;
  background: transparent;
  color: var(--text-primary);
  font: inherit;
  font-size: var(--text-md);
  text-align: left;
  cursor: pointer;
}

.sheet-tab-menu :deep([role="menuitem"]:hover:not(:disabled)),
.sheet-tab-menu :deep([role="menuitem"]:focus) {
  background: color-mix(in srgb, var(--accent) 12%, transparent);
}

.sheet-tab-menu :deep([role="menuitem"]:disabled) {
  color: var(--text-disabled);
  cursor: default;
}

.sheet-tab-menu :deep(.menu-icon) {
  display: inline-flex;
  flex: none;
  width: 1rem;
}

.menu-separator {
  margin: var(--space-1) 0;
  border: none;
  border-top: var(--border-width) solid var(--border-subtle);
}
</style>
