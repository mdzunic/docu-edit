<script setup lang="ts">
// A slide-management context menu (extracted from SlideFilmstrip, SPEC-099):
// a fixed-position role="menu" of buttons at a viewport point, a backdrop that
// dismisses it, focus on the first item on open, and arrow/Home/End/Escape
// keyboard movement. The filmstrip's slide menu, the slide sorter's card menu
// (the same `slide-menu-*` items) and its section menu all render this one
// component, so the menu behaviour is never forked. Runtime prop/emit
// declarations, not the generic forms (espree lints .vue scripts).
import { computed, nextTick, onMounted, useTemplateRef } from "vue";
import { asMenuItems, focusSelector, menuButtons } from "./slideRender";

const props = defineProps({
  /** `{ action, label, testid }` entries, in menu order. */
  items: { type: Array, required: true },
  x: { type: Number, default: 0 },
  y: { type: Number, default: 0 },
  testid: { type: String, default: "slide-context-menu" },
  backdropTestid: { type: String, default: "slide-menu-backdrop" },
  label: { type: String, default: "Slide options" },
});

const emit = defineEmits(["run", "close"]);

const entries = computed(() => asMenuItems(props.items));

const menuRef = useTemplateRef("menu");

onMounted(() => {
  void nextTick(() => focusSelector(menuRef.value, '[role="menuitem"]'));
});

function onKeydown(event = new KeyboardEvent("keydown")) {
  const items = menuButtons(menuRef.value);
  const active = items.findIndex((item) => item === document.activeElement);
  if (event.key === "ArrowDown") {
    event.preventDefault();
    items[(active + 1) % items.length]?.focus();
  } else if (event.key === "ArrowUp") {
    event.preventDefault();
    items[(active - 1 + items.length) % items.length]?.focus();
  } else if (event.key === "Escape") {
    event.preventDefault();
    event.stopPropagation();
    emit("close", true);
  } else if (event.key === "Home") {
    event.preventDefault();
    items[0]?.focus();
  } else if (event.key === "End") {
    event.preventDefault();
    items[items.length - 1]?.focus();
  }
}
</script>

<template>
  <div
    ref="menu"
    class="slide-context-menu"
    role="menu"
    :aria-label="label"
    :data-testid="testid"
    :style="{ left: `${x}px`, top: `${y}px` }"
    tabindex="-1"
    @keydown="onKeydown"
  >
    <button
      v-for="item in entries"
      :key="item.action"
      type="button"
      class="slide-context-menu-item"
      role="menuitem"
      :data-testid="item.testid"
      :disabled="item.disabled === true"
      @click="emit('run', item.action)"
    >
      {{ item.label }}
    </button>
  </div>
  <div
    class="menu-backdrop"
    :data-testid="backdropTestid"
    @click="emit('close', false)"
    @contextmenu.prevent="emit('close', false)"
  />
</template>

<style scoped>
.slide-context-menu {
  position: fixed;
  z-index: 40;
  display: flex;
  flex-direction: column;
  min-width: 10rem;
  padding: var(--space-1);
  background: var(--surface-raised);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
  box-shadow: var(--shadow-2);
}

.slide-context-menu-item {
  display: block;
  width: 100%;
  text-align: left;
  padding: var(--space-2) var(--space-3);
  border: 0;
  border-radius: var(--radius-sm);
  background: transparent;
  color: var(--text-primary);
  cursor: pointer;
}

.slide-context-menu-item:hover {
  background: var(--surface-sunken);
}

.slide-context-menu-item:disabled {
  color: var(--text-secondary);
  cursor: default;
}

.slide-context-menu-item:focus-visible {
  background: var(--surface-sunken);
  outline: var(--border-width) solid var(--focus-ring);
  outline-offset: calc(-1 * var(--border-width));
}

.menu-backdrop {
  position: fixed;
  inset: 0;
  z-index: 39;
}
</style>
