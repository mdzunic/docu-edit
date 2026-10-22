<script setup lang="ts">
// Go To (SPEC-134), on BaseDialog: the list (`sheet-goto-list` of
// `sheet-goto-item`) — the names and tables the Name Box lists, then the last
// ten Go To origins — the reference field (`sheet-goto-reference`), the error
// (`sheet-goto-error`), "Special…" (`sheet-goto-special`), Cancel and OK
// (`sheet-goto-ok`). A click on an item fills the field; Enter or a
// double-click on one goes there; Enter in the field is OK. Shown while
// useSheetGoTo's `dialogOpen`. Runtime prop declarations per the espree
// constraint (see AppShell.vue).
import { computed } from "vue";
import { asSheetGoTo } from "../../views/spreadsheet/useSheetGoTo";
import BaseDialog from "../BaseDialog.vue";
import UiButton from "../ui/UiButton.vue";

const props = defineProps({
  /** useSheetGoTo's return. */
  goTo: { type: Object, required: true },
});

const g = computed(() => asSheetGoTo(props.goTo));

function onInput(event = new Event("input")) {
  if (event.target instanceof HTMLInputElement) {
    g.value.setReference(event.target.value);
  }
}

/** The field's text is selected on focus, so typing replaces the origin it
 * was pre-filled with. */
function selectField(event = new FocusEvent("focus")) {
  if (event.target instanceof HTMLInputElement) {
    event.target.select();
  }
}

function onFieldKey(event = new KeyboardEvent("keydown")) {
  if (event.key === "Enter") {
    event.preventDefault();
    void g.value.confirm();
  }
}

function onItemKey(key = "", event = new KeyboardEvent("keydown")) {
  if (event.key === "Enter") {
    event.preventDefault();
    void g.value.activate(key);
  } else if (event.key === " ") {
    event.preventDefault();
    g.value.choose(key);
  }
}
</script>

<template>
  <BaseDialog
    v-if="g.dialogOpen.value"
    title="Go To"
    testid="sheet-goto-dialog"
    width="22rem"
    @dismiss="g.cancel()"
  >
    <p
      id="sheet-goto-list-label"
      class="goto-caption"
    >
      Go to
    </p>
    <ul
      class="goto-list"
      role="listbox"
      aria-labelledby="sheet-goto-list-label"
      data-testid="sheet-goto-list"
    >
      <li
        v-for="item in g.items.value"
        :key="item.key"
        class="goto-item"
        :class="{ 'is-chosen': g.chosen.value === item.key, 'is-recent': item.kind === 'recent' }"
        role="option"
        data-testid="sheet-goto-item"
        :data-kind="item.kind"
        :aria-selected="g.chosen.value === item.key ? 'true' : 'false'"
        tabindex="0"
        @click="g.choose(item.key)"
        @dblclick="g.activate(item.key)"
        @keydown="onItemKey(item.key, $event)"
      >
        {{ item.text }}
      </li>
    </ul>
    <label
      class="goto-caption"
      for="sheet-goto-reference"
    >Reference</label>
    <input
      id="sheet-goto-reference"
      class="goto-field"
      data-testid="sheet-goto-reference"
      type="text"
      autocomplete="off"
      spellcheck="false"
      autofocus
      :value="g.reference.value"
      :aria-invalid="g.error.value ? 'true' : undefined"
      :aria-describedby="g.error.value ? 'sheet-goto-error' : undefined"
      @input="onInput"
      @keydown="onFieldKey"
      @focus="selectField"
    >
    <p
      v-if="g.error.value"
      id="sheet-goto-error"
      class="goto-error"
      data-testid="sheet-goto-error"
      role="alert"
    >
      {{ g.error.value }}
    </p>
    <template #actions>
      <UiButton
        variant="secondary"
        data-testid="sheet-goto-special"
        @click="g.openSpecial()"
      >
        Special…
      </UiButton>
      <UiButton
        variant="secondary"
        data-testid="sheet-goto-cancel"
        @click="g.cancel()"
      >
        Cancel
      </UiButton>
      <UiButton
        variant="primary"
        data-testid="sheet-goto-ok"
        @click="g.confirm()"
      >
        OK
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
.goto-caption {
  display: block;
  margin: 0 0 var(--space-1);
  color: var(--text-secondary);
  font-size: var(--text-sm);
}

.goto-list {
  height: 10rem;
  margin: 0 0 var(--space-3);
  padding: var(--space-1);
  overflow-y: auto;
  list-style: none;
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
}

.goto-item {
  padding: var(--space-1) var(--space-2);
  border-radius: var(--radius-sm);
  color: var(--text-primary);
  cursor: pointer;
  overflow-wrap: anywhere;
}

.goto-item.is-recent {
  color: var(--text-secondary);
}

.goto-item:hover {
  background: var(--state-hover);
}

.goto-item.is-chosen {
  background: var(--surface-selected);
  color: var(--text-accent);
}

.goto-item:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: calc(-1 * var(--space-1));
}

.goto-field {
  width: 100%;
  box-sizing: border-box;
  font: inherit;
  padding: var(--space-2) var(--space-3);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
}

.goto-field:focus {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: calc(-1 * var(--border-width));
}

.goto-error {
  margin: var(--space-2) 0 0;
  color: var(--text-danger);
}
</style>
