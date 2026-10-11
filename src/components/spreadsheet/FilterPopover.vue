<script setup lang="ts">
// A header's filter dropdown (SPEC-128 "UI"): `filter-popover`, a
// non-modal role="dialog" under the header's `filter-button`. Top to bottom:
// the column's sorts (`filter-sort-asc` / `filter-sort-desc`, "Sort A to Z"
// or "Sort Smallest to Largest" … by the column's type), `filter-clear`
// (`Clear filter from "Region"`, disabled while the column has no criteria),
// the conditions submenu (`filter-conditions` — "Text Filters", "Number
// Filters" or "Date Filters" — its items `filter-condition` with
// data-condition), the search box (`filter-search`, narrowing the list
// case-insensitively), "(Select All)" (`filter-select-all`, tri-state through
// aria-checked; "(Select All Search Results)" while searching) and one
// `filter-value` checkbox per entry (its text the value alone — the check
// marks are drawn by CSS), "Not all items showing"
// (`filter-not-all`) past FILTER_VALUE_LIST_MAX, then OK (`filter-ok`,
// disabled while nothing is checked) and Cancel (`filter-cancel`).
// OK emits apply(keys) — the checked entries' keys; while searching, only
// the checked matching ones. A column filtered by colour or icon shows only
// its message (`filter-preserved`) and `filter-clear`. Escape emits
// cancel("escape"), Cancel cancel("done"), a press outside it (but on no
// filter button, which toggles it itself) cancel("outside"). Runtime prop
// declarations per the espree constraint (see AppShell.vue).
import { computed, nextTick, onBeforeUnmount, onMounted, ref, useTemplateRef } from "vue";
import { asValueEntries } from "../../views/spreadsheet/sortFilterLabels";
import { PRESERVED_FILTER_MESSAGE } from "../../views/spreadsheet/useFilterPopover";
import UiButton from "../ui/UiButton.vue";

const props = defineProps({
  /** The column's header text ("Region"). */
  title: { type: String, required: true },
  sortAscLabel: { type: String, default: "Sort A to Z" },
  sortDescLabel: { type: String, default: "Sort Z to A" },
  /** "Text Filters" | "Number Filters" | "Date Filters". */
  conditionsLabel: { type: String, default: "Text Filters" },
  /** The submenu's items: { id, label }[]. */
  conditions: { type: Array, default: () => [] },
  /** The value list (ValueEntry[]). */
  entries: { type: Array, default: () => [] },
  truncated: { type: Boolean, default: false },
  hasCriteria: { type: Boolean, default: false },
  /** Filtered by colour or icon: only the message and Clear filter. */
  preserved: { type: Boolean, default: false },
  /** Viewport px: { left, top }. */
  anchor: { type: Object, default: () => ({ left: 0, top: 0 }) },
});

const emit = defineEmits(["sort", "clear", "condition", "apply", "cancel"]);

const root = useTemplateRef("root");
const dialogLabel = computed(() => `Filter "${props.title}"`);
const search = ref("");
const conditionsOpen = ref(false);
const all = computed(() => asValueEntries(props.entries));
const checks = ref(Object.fromEntries(all.value.map((entry) => [entry.key, entry.checked])));

const term = computed(() => search.value.trim().toLowerCase());
const shown = computed(() => (term.value ? all.value.filter((entry) => entry.text.toLowerCase().includes(term.value)) : all.value));
const allState = computed(() => {
  const count = shown.value.filter((entry) => checks.value[entry.key]).length;
  return count === 0 ? "false" : count === shown.value.length ? "true" : "mixed";
});
/** What OK writes: every checked entry, or the checked matches while searching. */
const chosen = computed(() => shown.value.filter((entry) => checks.value[entry.key]).map((entry) => entry.key));
const items = computed(() =>
  (Array.isArray(props.conditions) ? props.conditions : []).map((item) => ({ id: String(Reflect.get(Object(item), "id")), label: String(Reflect.get(Object(item), "label")) })),
);
const position = computed(() => {
  const left = Number(Reflect.get(Object(props.anchor), "left")) || 0;
  const top = Number(Reflect.get(Object(props.anchor), "top")) || 0;
  const width = typeof window === "undefined" ? 1024 : window.innerWidth;
  return { left: `${Math.max(4, Math.min(left, width - 280))}px`, top: `${top + 2}px` };
});

function toggleAll() {
  const next = allState.value !== "true";
  for (const entry of shown.value) {
    checks.value[entry.key] = next;
  }
}

function toggle(key = "") {
  checks.value[key] = !checks.value[key];
}

function apply() {
  if (chosen.value.length > 0) {
    emit("apply", chosen.value);
  }
}

function openConditions(focus = false) {
  conditionsOpen.value = true;
  if (focus) {
    void nextTick(() => {
      const first = root.value ? root.value.querySelector('[data-testid="filter-condition"]') : null;
      if (first instanceof HTMLElement) {
        first.focus();
      }
    });
  }
}

function onConditionsKeydown(event = new KeyboardEvent("keydown")) {
  if (event.key === "ArrowRight" || event.key === "ArrowDown") {
    event.preventDefault();
    openConditions(true);
  }
}

/** Arrow keys move through the submenu; Escape or ArrowLeft closes it. */
function onMenuKeydown(event = new KeyboardEvent("keydown")) {
  const menuItems = root.value ? Array.from(root.value.querySelectorAll('[data-testid="filter-condition"]')) : [];
  const at = menuItems.indexOf(document.activeElement ?? document.body);
  if (event.key === "ArrowDown" || event.key === "ArrowUp") {
    event.preventDefault();
    const next = menuItems[(at + (event.key === "ArrowDown" ? 1 : menuItems.length - 1)) % menuItems.length];
    if (next instanceof HTMLElement) {
      next.focus();
    }
  } else if (event.key === "Escape" || event.key === "ArrowLeft") {
    event.preventDefault();
    event.stopPropagation();
    conditionsOpen.value = false;
    const trigger = root.value ? root.value.querySelector('[data-testid="filter-conditions"]') : null;
    if (trigger instanceof HTMLElement) {
      trigger.focus();
    }
  }
}

function onKeydown(event = new KeyboardEvent("keydown")) {
  if (event.key === "Escape") {
    event.preventDefault();
    event.stopPropagation();
    emit("cancel", "escape");
  }
}

function onDocPointerDown(event = new PointerEvent("pointerdown")) {
  const target = event.target;
  if (!(target instanceof Element) || (root.value && root.value.contains(target)) || target.closest('[data-testid="filter-button"]')) {
    return;
  }
  emit("cancel", "outside");
}

onMounted(() => {
  document.addEventListener("pointerdown", onDocPointerDown, true);
  void nextTick(() => {
    const first = root.value ? (root.value.querySelector('[data-testid="filter-search"]') ?? root.value.querySelector('[data-testid="filter-clear"]')) : null;
    if (first instanceof HTMLElement) {
      first.focus();
    }
  });
});

onBeforeUnmount(() => {
  document.removeEventListener("pointerdown", onDocPointerDown, true);
});
</script>

<template>
  <div
    ref="root"
    class="filter-popover"
    role="dialog"
    :aria-label="dialogLabel"
    data-testid="filter-popover"
    :style="position"
    @keydown="onKeydown"
  >
    <template v-if="props.preserved">
      <p
        class="filter-preserved"
        data-testid="filter-preserved"
      >
        {{ PRESERVED_FILTER_MESSAGE }}
      </p>
      <button
        type="button"
        class="filter-action"
        data-testid="filter-clear"
        @click="emit('clear')"
      >
        Clear filter from "{{ props.title }}"
      </button>
    </template>
    <template v-else>
      <button
        type="button"
        class="filter-action"
        data-testid="filter-sort-asc"
        @click="emit('sort', false)"
      >
        {{ props.sortAscLabel }}
      </button>
      <button
        type="button"
        class="filter-action"
        data-testid="filter-sort-desc"
        @click="emit('sort', true)"
      >
        {{ props.sortDescLabel }}
      </button>
      <hr class="filter-separator">
      <button
        type="button"
        class="filter-action"
        data-testid="filter-clear"
        :disabled="!props.hasCriteria"
        @click="emit('clear')"
      >
        Clear filter from "{{ props.title }}"
      </button>
      <div class="filter-conditions">
        <button
          type="button"
          class="filter-action filter-conditions-trigger"
          data-testid="filter-conditions"
          aria-haspopup="menu"
          :aria-expanded="conditionsOpen ? 'true' : 'false'"
          @click="conditionsOpen ? (conditionsOpen = false) : openConditions(false)"
          @keydown="onConditionsKeydown"
        >
          <span>{{ props.conditionsLabel }}</span>
          <span aria-hidden="true">▸</span>
        </button>
        <div
          v-if="conditionsOpen"
          class="filter-conditions-menu"
          role="menu"
          :aria-label="props.conditionsLabel"
          data-testid="filter-conditions-menu"
          @keydown="onMenuKeydown"
        >
          <button
            v-for="item in items"
            :key="item.id"
            type="button"
            role="menuitem"
            class="filter-action"
            data-testid="filter-condition"
            :data-condition="item.id"
            tabindex="-1"
            @click="emit('condition', item.id)"
          >
            {{ item.label }}
          </button>
        </div>
      </div>
      <input
        v-model="search"
        type="search"
        class="filter-search"
        placeholder="Search"
        aria-label="Search values"
        data-testid="filter-search"
      >
      <div
        class="filter-values"
        role="group"
        aria-label="Values"
      >
        <button
          type="button"
          role="checkbox"
          class="filter-check"
          data-testid="filter-select-all"
          :aria-checked="allState"
          @click="toggleAll"
        >
          <span
            class="filter-mark"
            :class="{ 'is-checked': allState === 'true', 'is-mixed': allState === 'mixed' }"
            aria-hidden="true"
          />
          <span>{{ term ? "(Select All Search Results)" : "(Select All)" }}</span>
        </button>
        <button
          v-for="entry in shown"
          :key="entry.key"
          type="button"
          role="checkbox"
          class="filter-check"
          data-testid="filter-value"
          :aria-checked="checks[entry.key] ? 'true' : 'false'"
          @click="toggle(entry.key)"
        >
          <span
            class="filter-mark"
            :class="{ 'is-checked': checks[entry.key] }"
            aria-hidden="true"
          />
          <span>{{ entry.text }}</span>
        </button>
      </div>
      <p
        v-if="props.truncated"
        class="filter-not-all"
        data-testid="filter-not-all"
      >
        Not all items showing
      </p>
      <div class="filter-buttons">
        <UiButton
          variant="primary"
          data-testid="filter-ok"
          :disabled="chosen.length === 0"
          @click="apply"
        >
          OK
        </UiButton>
        <UiButton
          variant="secondary"
          data-testid="filter-cancel"
          @click="emit('cancel', 'done')"
        >
          Cancel
        </UiButton>
      </div>
    </template>
  </div>
</template>

<style scoped>
.filter-popover {
  position: fixed;
  z-index: 40;
  box-sizing: border-box;
  display: flex;
  flex-direction: column;
  gap: var(--space-1);
  width: 17rem;
  max-height: calc(100vh - 2rem);
  padding: var(--space-2);
  overflow-y: auto;
  color: var(--text-primary);
  font-family: var(--font-ui);
  font-size: var(--text-sm);
  background: var(--surface-raised);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
  box-shadow: var(--shadow-3);
}

.filter-action {
  display: flex;
  align-items: center;
  justify-content: space-between;
  min-height: var(--height-control);
  padding: 0 var(--space-2);
  font: inherit;
  text-align: left;
  color: inherit;
  background: transparent;
  border: none;
  border-radius: var(--radius-sm);
  cursor: pointer;
}

.filter-action:hover:not(:disabled),
.filter-action:focus-visible {
  background: var(--state-hover);
}

.filter-action:disabled {
  color: var(--text-disabled);
  cursor: default;
}

.filter-separator {
  width: 100%;
  margin: var(--space-1) 0;
  border: none;
  border-top: var(--border-width) solid var(--border-subtle);
}

.filter-conditions {
  display: flex;
  flex-direction: column;
}

.filter-conditions-menu {
  display: flex;
  flex-direction: column;
  margin-left: var(--space-3);
  padding-left: var(--space-2);
  border-left: var(--border-width) solid var(--border-subtle);
}

.filter-search {
  height: var(--height-control);
  padding: 0 var(--space-2);
  font: inherit;
  color: var(--text-primary);
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
}

.filter-values {
  display: flex;
  flex-direction: column;
  max-height: 14rem;
  overflow-y: auto;
  border: var(--border-width) solid var(--border-subtle);
  border-radius: var(--radius-sm);
}

.filter-check {
  display: flex;
  align-items: center;
  gap: var(--space-2);
  padding: var(--space-1) var(--space-2);
  font: inherit;
  text-align: left;
  color: inherit;
  white-space: nowrap;
  background: transparent;
  border: none;
  cursor: pointer;
}

.filter-check:hover,
.filter-check:focus-visible {
  background: var(--state-hover);
}

.filter-mark {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  flex: none;
  width: 0.9rem;
  height: 0.9rem;
  font-size: var(--text-xs);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
}

/* The marks are drawn, not text, so an item's text is its value alone. */
.filter-mark.is-checked::after {
  content: "✓";
}

.filter-mark.is-mixed::after {
  content: "–";
}

.filter-not-all,
.filter-preserved {
  margin: 0;
  color: var(--text-secondary);
}

.filter-buttons {
  display: flex;
  justify-content: flex-end;
  gap: var(--space-2);
  margin-top: var(--space-1);
}
</style>
