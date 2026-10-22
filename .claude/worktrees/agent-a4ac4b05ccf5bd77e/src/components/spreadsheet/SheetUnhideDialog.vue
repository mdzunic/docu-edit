<script setup lang="ts">
// SPEC-121 "Unhide…": every "hidden" sheet (never a "veryHidden" one — the
// view leaves those out of `sheets`) in a multi-select listbox. Click
// toggles an option; ArrowUp/ArrowDown/Home/End move the active option (a
// roving tabindex), Space toggles it and Enter confirms. Unhide is disabled
// until something is chosen, and confirms the chosen model indices in list
// order. Runtime prop declarations per the espree constraint (see
// AppShell.vue); the prop is narrowed in sheetTabBar.ts.
import { computed, nextTick, ref, useTemplateRef, watch } from "vue";
import BaseDialog from "../BaseDialog.vue";
import UiButton from "../ui/UiButton.vue";
import { asHiddenSheetEntries } from "./sheetTabBar";

const props = defineProps({
  /** `[{ index, name }]`: the hidden sheets, in sheet order. */
  sheets: { type: Array, default: () => [] },
});

const emit = defineEmits(["confirm", "cancel"]);

const entries = computed(() => asHiddenSheetEntries(props.sheets));
/** Whether each entry (by list position) is chosen. */
const chosen = ref(entries.value.map(() => false));
const activePosition = ref(0);
const listbox = useTemplateRef("listbox");

watch(entries, (list) => {
  chosen.value = list.map(() => false);
  activePosition.value = 0;
});

const anyChosen = computed(() => chosen.value.some(Boolean));

function focusOption(position = 0) {
  const option = listbox.value?.querySelectorAll('[role="option"]')[position];
  if (option instanceof HTMLElement) {
    option.focus();
  }
}

function toggle(position = 0) {
  chosen.value[position] = !chosen.value[position];
  activePosition.value = position;
}

function confirm() {
  if (anyChosen.value) {
    emit(
      "confirm",
      entries.value.filter((_, position) => chosen.value[position]).map((sheet) => sheet.index),
    );
  }
}

function onKeydown(event = new KeyboardEvent("keydown")) {
  const count = entries.value.length;
  if (count === 0) {
    return;
  }
  if (event.key === " ") {
    event.preventDefault();
    toggle(activePosition.value);
    return;
  }
  if (event.key === "Enter") {
    event.preventDefault();
    confirm();
    return;
  }
  const targets = {
    ArrowDown: Math.min(count - 1, activePosition.value + 1),
    ArrowUp: Math.max(0, activePosition.value - 1),
    Home: 0,
    End: count - 1,
  };
  const target = Object.entries(targets).find(([key]) => key === event.key)?.[1];
  if (target === undefined) {
    return;
  }
  event.preventDefault();
  activePosition.value = target;
  void nextTick(() => focusOption(target));
}
</script>

<template>
  <BaseDialog
    title="Unhide sheets"
    testid="sheet-unhide-dialog"
    width="22rem"
    @dismiss="emit('cancel')"
  >
    <ul
      ref="listbox"
      class="unhide-list"
      role="listbox"
      aria-multiselectable="true"
      aria-label="Hidden sheets"
      data-testid="sheet-unhide-list"
      @keydown="onKeydown"
    >
      <li
        v-for="(sheet, position) in entries"
        :key="sheet.index"
        class="unhide-option"
        :class="{ 'is-chosen': chosen[position] }"
        role="option"
        data-testid="sheet-unhide-option"
        :aria-selected="chosen[position] ? 'true' : 'false'"
        :tabindex="position === activePosition ? 0 : -1"
        :autofocus="position === 0 ? true : undefined"
        @click="toggle(position)"
      >
        <span
          class="unhide-check"
          aria-hidden="true"
        />
        <span class="unhide-name">{{ sheet.name }}</span>
      </li>
    </ul>
    <template #actions>
      <UiButton
        variant="secondary"
        data-testid="sheet-unhide-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </UiButton>
      <UiButton
        variant="primary"
        data-testid="sheet-unhide-confirm"
        :disabled="!anyChosen"
        @click="confirm"
      >
        Unhide
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
.unhide-list {
  max-height: 16rem;
  margin: 0;
  padding: var(--space-1);
  overflow-y: auto;
  list-style: none;
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
}

.unhide-option {
  display: flex;
  align-items: center;
  gap: var(--space-2);
  padding: var(--space-1) var(--space-2);
  border-radius: var(--radius-sm);
  color: var(--text-primary);
  cursor: pointer;
}

.unhide-option:hover {
  background: var(--state-hover);
}

.unhide-option:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: calc(-1 * var(--space-1));
}

/* The check box is drawn, not printed, so an option's text stays the sheet
   name whether or not it is chosen. */
.unhide-check {
  flex: none;
  width: 0.875rem;
  height: 0.875rem;
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
}

.unhide-option.is-chosen .unhide-check {
  border-color: var(--accent);
  background: var(--accent);
}

.unhide-name {
  overflow-wrap: anywhere;
}
</style>
