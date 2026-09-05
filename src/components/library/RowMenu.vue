<script setup lang="ts">
import { computed, nextTick, useTemplateRef, watch } from "vue";
import { asMenuItems } from "../../lib/libraryMenu";

// The per-row/card overflow menu (SPEC-067). Reused by list rows, grid cards
// and the trash view, configured by its `items`, so every one keeps the same
// keyboard behaviour and the SPEC-003 data-testids (row-menu, row-menu-list,
// menu-rename, menu-duplicate, menu-delete) are preserved wherever they apply.
// The parent owns which menu is open (one at a time); this component focuses
// the first item when opened and supports arrow-key navigation so the menu is
// reachable without the pointer (AC-62). Runtime prop/emit declarations — the
// .vue ESLint parser rejects TS-only syntax (see AppShell.vue). Each item is
// { id, label, testid, danger? }.
const props = defineProps({
  open: { type: Boolean, default: false },
  label: { type: String, required: true },
  items: { type: Array, default: () => [] },
  buttonTestid: { type: String, default: "row-menu" },
  menuTestid: { type: String, default: "row-menu-list" },
});

const emit = defineEmits(["toggle", "close", "select"]);

const menuItems = computed(() => asMenuItems(props.items));

const menuRef = useTemplateRef("menu");
const buttonRef = useTemplateRef("button");

watch(
  () => props.open,
  async (open) => {
    if (open) {
      await nextTick();
      focusItem(0);
    }
  },
);

function itemButtons() {
  const nodes = menuRef.value?.querySelectorAll("[role=menuitem]") ?? [];
  const buttons = [];
  for (const node of nodes) {
    if (node instanceof HTMLElement) {
      buttons.push(node);
    }
  }
  return buttons;
}

function focusItem(index = 0) {
  const buttons = itemButtons();
  if (buttons.length === 0) {
    return;
  }
  const clamped = (index + buttons.length) % buttons.length;
  buttons[clamped]?.focus();
}

// Activating an item closes the menu, and the item's own button goes with
// it — so focus is returned to the trigger BEFORE the parent is told, not
// after. The order matters: a command like Rename opens a dialog, and that
// dialog records where focus was so it can put it back when it closes. Emit
// first and it records a button that is already being unmounted, so both
// surfaces drop focus and the keyboard user lands back on <body>.
function selectItem(id = "") {
  buttonRef.value?.focus();
  emit("select", id);
}

function onMenuKeydown(event = new KeyboardEvent("keydown")) {
  const buttons = itemButtons();
  const active = document.activeElement;
  const current = active instanceof HTMLElement ? buttons.indexOf(active) : -1;
  if (event.key === "ArrowDown") {
    event.preventDefault();
    focusItem(current + 1);
  } else if (event.key === "ArrowUp") {
    event.preventDefault();
    focusItem(current - 1);
  } else if (event.key === "Escape") {
    event.preventDefault();
    emit("close");
    buttonRef.value?.focus();
  } else if (event.key === "Home") {
    event.preventDefault();
    focusItem(0);
  } else if (event.key === "End") {
    event.preventDefault();
    focusItem(buttons.length - 1);
  }
}
</script>

<template>
  <div
    class="row-menu-wrap"
    @click.stop
    @keydown.stop
  >
    <button
      ref="button"
      type="button"
      class="menu-button"
      :data-testid="props.buttonTestid"
      :aria-label="props.label"
      aria-haspopup="menu"
      :aria-expanded="props.open"
      @click="emit('toggle')"
    >
      ⋯
    </button>
    <div
      v-if="props.open"
      ref="menu"
      class="row-menu"
      role="menu"
      :data-testid="props.menuTestid"
      @keydown="onMenuKeydown"
    >
      <button
        v-for="item in menuItems"
        :key="item.id"
        type="button"
        role="menuitem"
        :class="{ 'is-danger': item.danger }"
        :data-testid="item.testid"
        @click="selectItem(item.id)"
      >
        {{ item.label }}
      </button>
    </div>
  </div>
</template>

<style scoped>
.row-menu-wrap {
  position: relative;
  display: inline-flex;
}

.menu-button {
  font: inherit;
  font-weight: var(--weight-bold);
  line-height: 1;
  padding: var(--space-2) var(--space-3);
  border: none;
  border-radius: var(--radius-sm);
  background: transparent;
  cursor: pointer;
  color: var(--text-secondary);
}

.menu-button:hover {
  background: var(--surface-sunken);
}

.row-menu {
  position: absolute;
  right: 0;
  top: 2rem;
  display: flex;
  flex-direction: column;
  min-width: 10rem;
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  box-shadow: var(--shadow-2);
  z-index: 20;
}

.row-menu button {
  font: inherit;
  text-align: left;
  padding: var(--space-3) var(--space-5);
  border: none;
  background: var(--surface-base);
  cursor: pointer;
  color: var(--text-primary);
  white-space: nowrap;
}

.row-menu button:hover,
.row-menu button:focus-visible {
  background: var(--surface-sunken);
}

.row-menu button.is-danger {
  color: var(--text-danger);
}
</style>
