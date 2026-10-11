<script setup lang="ts">
// SPEC-127's Insert Function dialog (`insert-function-dialog`, on
// BaseDialog): a search field (`insert-function-search`, focused on open), a
// category select (`insert-function-category`: "All", then every category
// holding a registered function, in FUNCTION_CATEGORY_ORDER), the list
// (`insert-function-list` of `insert-function-item`, functionHelp.ts's
// searchFunctions ranking), the detail pane (`insert-function-detail`: the
// signature, the summary and every argument's description) and "Insert"
// (`insert-function-insert`). Up/Down move the selection, Enter and a
// double-click insert. Emits insert(name) / cancel; useFunctionAssist puts
// `=NAME(` into the editor. Runtime prop declarations per the espree
// constraint (see AppShell.vue).
import { computed, ref, watch } from "vue";
import { functionSignature, searchFunctions } from "../../lib/spreadsheet/formula/functionHelp";
import { FUNCTION_CATEGORY_ORDER, getFunction, registeredFunctionNames, type FunctionCategory } from "../../lib/spreadsheet/formula/registry";
import BaseDialog from "../BaseDialog.vue";
import UiButton from "../ui/UiButton.vue";

const emit = defineEmits(["insert", "cancel"]);

const ALL = "All";
const query = ref("");
const category = ref(ALL);
const selected = ref("");

/** "All", then the categories that hold a registered function. */
const categories = computed(() => {
  const held = new Set(registeredFunctionNames().map((name) => getFunction(name)?.category));
  return [ALL, ...FUNCTION_CATEGORY_ORDER.filter((name) => held.has(name))];
});
const names = computed(() => searchFunctions(query.value, category.value === ALL ? null : (category.value as FunctionCategory)));
const spec = computed(() => (selected.value ? getFunction(selected.value) : undefined));

watch(
  names,
  (list) => {
    if (!list.includes(selected.value)) {
      selected.value = list[0] ?? "";
    }
  },
  { immediate: true },
);

function move(step: number): void {
  const list = names.value;
  if (list.length === 0) {
    return;
  }
  const at = list.indexOf(selected.value);
  selected.value = list[Math.max(0, Math.min(list.length - 1, at + step))]!;
}

function insert(): void {
  if (selected.value) {
    emit("insert", selected.value);
  }
}

function onSearchKey(event: KeyboardEvent): void {
  if (event.key === "ArrowDown" || event.key === "ArrowUp") {
    event.preventDefault();
    move(event.key === "ArrowDown" ? 1 : -1);
  } else if (event.key === "Enter") {
    event.preventDefault();
    insert();
  }
}
</script>

<template>
  <BaseDialog
    title="Insert function"
    testid="insert-function-dialog"
    width="34rem"
    @dismiss="emit('cancel')"
  >
    <div class="insert-function">
      <div class="insert-function-filters">
        <input
          v-model="query"
          class="insert-function-search"
          data-testid="insert-function-search"
          type="search"
          aria-label="Search for a function"
          placeholder="Search for a function"
          autofocus
          @keydown="onSearchKey"
        >
        <select
          v-model="category"
          class="insert-function-category"
          data-testid="insert-function-category"
          aria-label="Category"
        >
          <option
            v-for="name in categories"
            :key="name"
            :value="name"
          >
            {{ name }}
          </option>
        </select>
      </div>
      <ul
        class="insert-function-list"
        data-testid="insert-function-list"
        role="listbox"
        aria-label="Functions"
      >
        <li
          v-for="name in names"
          :key="name"
          class="insert-function-item"
          :class="{ 'is-selected': name === selected }"
          data-testid="insert-function-item"
          role="option"
          :data-name="name"
          :aria-selected="name === selected ? 'true' : 'false'"
          @click="selected = name"
          @dblclick="insert"
        >
          {{ name }}
        </li>
      </ul>
      <div
        class="insert-function-detail"
        data-testid="insert-function-detail"
        aria-live="polite"
      >
        <template v-if="spec">
          <p class="insert-function-signature">
            {{ functionSignature(spec) }}
          </p>
          <p>{{ spec.summary }}</p>
          <dl class="insert-function-args">
            <template
              v-for="arg in spec.args"
              :key="arg.name"
            >
              <dt>{{ arg.name }}</dt>
              <dd>{{ arg.description }}</dd>
            </template>
          </dl>
        </template>
        <p v-else>
          No function matches.
        </p>
      </div>
    </div>
    <template #actions>
      <UiButton
        variant="secondary"
        data-testid="insert-function-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </UiButton>
      <UiButton
        variant="primary"
        data-testid="insert-function-insert"
        :disabled="!selected"
        @click="insert"
      >
        Insert
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
.insert-function {
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
  color: var(--text-primary);
  font-size: var(--text-md);
}

.insert-function-filters {
  display: flex;
  gap: var(--space-2);
}

.insert-function-search,
.insert-function-category {
  height: var(--height-control);
  padding: 0 var(--space-2);
  font: inherit;
  color: var(--text-primary);
  background: var(--surface-raised);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
}

.insert-function-search {
  flex: 1 1 auto;
}

.insert-function-list {
  height: 12rem;
  margin: 0;
  padding: 0;
  overflow-y: auto;
  list-style: none;
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
}

.insert-function-item {
  padding: var(--space-1) var(--space-2);
  cursor: pointer;
}

.insert-function-item.is-selected {
  background: var(--accent-subtle, var(--surface-sunken));
  font-weight: var(--weight-button);
}

.insert-function-detail {
  min-height: 6rem;
  font-size: var(--text-sm);
}

.insert-function-signature {
  font-weight: var(--weight-button);
}

.insert-function-args dt {
  font-style: italic;
}

.insert-function-args dd {
  margin: 0 0 var(--space-1) var(--space-3);
  color: var(--text-secondary);
}
</style>
