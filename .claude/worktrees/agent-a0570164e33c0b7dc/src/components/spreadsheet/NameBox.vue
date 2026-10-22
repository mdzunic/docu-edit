<script setup lang="ts">
// The Name Box (SPEC-117's `name-box`; SPEC-133 "The Name Box"): without the
// view's names (`names`, useNames) a read-only input showing the active
// cell's address. With them it is editable — what is typed replaces nothing
// until Enter (useNames' submit: a reference navigates, a name or table
// selects, a new name is defined); Escape restores the address and returns
// to the grid; leaving the box restores it too. Its dropdown
// (`name-box-dropdown`, also Alt+Down in the box) lists the visible names and
// tables (`name-box-option`, a listbox: Up/Down move, Enter picks, Escape
// closes). Runtime prop declarations per the espree constraint (see
// AppShell.vue).
import { computed, nextTick, ref, useTemplateRef } from "vue";
import { asNames } from "../../views/spreadsheet/useNames";

const props = defineProps({
  /** The active cell's address ("B3", "3R × 2C" during a drag). */
  address: { type: String, default: "" },
  /** useNames' return; absent = read-only. */
  names: { type: Object, default: null },
});

const names = computed(() => (props.names ? asNames(props.names) : null));
const shown = computed(() => names.value?.draft.value ?? props.address);
const open = ref(false);
const highlight = ref(0);
const list = useTemplateRef("list");
const toggle = useTemplateRef("toggle");

/** The text of an input event's target. */
function valueOf(event = new Event("input")) {
  return event.target instanceof HTMLInputElement ? event.target.value : "";
}

function selectAll(event = new FocusEvent("focus")) {
  if (event.target instanceof HTMLInputElement) {
    event.target.select();
  }
}

function onKeydown(event = new KeyboardEvent("keydown")) {
  if (event.key === "Enter") {
    event.preventDefault();
    void names.value?.submit();
  } else if (event.key === "Escape") {
    event.preventDefault();
    names.value?.cancel(true);
  } else if (event.key === "ArrowDown" && event.altKey) {
    event.preventDefault();
    void openList();
  }
}

function focusOption() {
  const items = list.value?.querySelectorAll('[data-testid="name-box-option"]');
  const item = items?.[highlight.value];
  if (item instanceof HTMLElement) {
    item.focus();
  }
}

async function openList() {
  open.value = !open.value;
  highlight.value = 0;
  if (open.value) {
    await nextTick();
    focusOption();
  }
}

function closeList(refocus = true) {
  open.value = false;
  if (refocus && toggle.value) {
    toggle.value.focus();
  }
}

function choose(key = "") {
  open.value = false;
  void names.value?.pick(key);
}

function onListKey(event = new KeyboardEvent("keydown")) {
  const count = names.value?.options.value.length ?? 0;
  if (event.key === "ArrowDown" || event.key === "ArrowUp") {
    event.preventDefault();
    highlight.value = count === 0 ? 0 : (highlight.value + (event.key === "ArrowDown" ? 1 : count - 1)) % count;
    focusOption();
  } else if (event.key === "Enter" || event.key === " ") {
    event.preventDefault();
    choose(names.value?.options.value[highlight.value]?.key ?? "");
  } else if (event.key === "Escape" || event.key === "Tab") {
    event.preventDefault();
    closeList();
  }
}

/** Focus leaving the list (a click elsewhere) closes it. */
function onListFocusOut(event = new FocusEvent("focusout")) {
  const next = event.relatedTarget;
  if (!(next instanceof Node) || (!list.value?.contains(next) && next !== toggle.value)) {
    closeList(false);
  }
}
</script>

<template>
  <div class="name-box-wrap">
    <input
      class="name-box"
      data-testid="name-box"
      type="text"
      aria-label="Name Box"
      spellcheck="false"
      autocomplete="off"
      :readonly="!names"
      :value="shown"
      @focus="selectAll"
      @input="names?.input(valueOf($event))"
      @keydown="onKeydown"
      @blur="names?.cancel(false)"
    >
    <button
      v-if="names"
      ref="toggle"
      type="button"
      class="name-box-toggle"
      data-testid="name-box-dropdown"
      aria-label="Names"
      aria-haspopup="listbox"
      :aria-expanded="open ? 'true' : 'false'"
      @click="openList"
    >
      ▾
    </button>
    <ul
      v-if="names && open"
      ref="list"
      class="name-box-list"
      role="listbox"
      aria-label="Names"
      data-testid="name-box-list"
      @keydown="onListKey"
      @focusout="onListFocusOut"
    >
      <li
        v-for="(option, index) in names.options.value"
        :key="option.key"
        role="option"
        tabindex="-1"
        class="name-box-option"
        data-testid="name-box-option"
        :data-kind="option.kind"
        :aria-selected="index === highlight ? 'true' : 'false'"
        @click="choose(option.key)"
      >
        {{ option.name }}
      </li>
      <li
        v-if="names.options.value.length === 0"
        class="name-box-empty"
        role="presentation"
      >
        No names
      </li>
    </ul>
  </div>
</template>

<style scoped>
.name-box-wrap {
  position: relative;
  flex: none;
  display: flex;
  align-items: stretch;
}

.name-box {
  width: 7rem;
  padding: var(--space-1) var(--space-2);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  background: var(--surface-sunken);
  color: inherit;
  font: inherit;
}

.name-box:focus-visible,
.name-box-toggle:focus-visible,
.name-box-option:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: 1px;
}

.name-box-toggle {
  padding: 0 var(--space-1);
  border: none;
  border-radius: var(--radius-sm);
  background: transparent;
  color: var(--text-secondary);
  font: inherit;
  cursor: pointer;
}

.name-box-toggle:hover {
  background: var(--surface-sunken);
  color: var(--text-primary);
}

.name-box-list {
  position: absolute;
  top: 100%;
  left: 0;
  z-index: 40;
  min-width: 100%;
  max-height: 16rem;
  margin: var(--space-1) 0 0;
  padding: var(--space-1) 0;
  overflow: auto;
  list-style: none;
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  box-shadow: var(--shadow-2);
}

.name-box-option,
.name-box-empty {
  padding: var(--space-1) var(--space-3);
  white-space: nowrap;
}

.name-box-option {
  cursor: pointer;
}

.name-box-option:hover,
.name-box-option[aria-selected="true"] {
  background: var(--surface-sunken);
}

.name-box-empty {
  color: var(--text-secondary);
}
</style>
