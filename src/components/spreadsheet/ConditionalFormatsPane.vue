<script setup lang="ts">
// The conditional-formatting rules manager (SPEC-129 "The side panel"): the
// `conditionalFormats` pane of the spreadsheet side panel, `cf-manager`,
// bound to useConditionalFormatManager. `cf-scope` ("Current selection" /
// "This sheet"), `cf-rule-list` (a listbox) of `cf-rule` options in priority
// order — each with its `cf-rule-preview`, `cf-rule-description`, editable
// `cf-rule-applies-to` (committed on change or Enter; an invalid entry is an
// inline `cf-rule-applies-to-error` and nothing is saved) and, for a
// formatting rule, `cf-rule-stop-if-true`; the actions `cf-add` (the quick
// presets, `cf-preset-*`), `cf-edit`, `cf-delete`, `cf-move-up`,
// `cf-move-down`, `cf-clear-selection` and `cf-clear-sheet`. While a rule is
// being added or edited the pane shows the rule editor instead. Runtime prop
// declarations per the espree constraint (see AppShell.vue).
import { computed, onBeforeUnmount, onMounted, useTemplateRef } from "vue";
import { CF_ADD_PRESET_GROUPS } from "../../views/spreadsheet/cfPresets";
import { asConditionalFormatManager, CF_SCOPES } from "../../views/spreadsheet/useConditionalFormatManager";
import AppIcon from "../AppIcon.vue";
import CfRulePreview from "./CfRulePreview.vue";
import ConditionalFormatEditor from "./ConditionalFormatEditor.vue";

const props = defineProps({
  /** useConditionalFormatManager's return. */
  manager: { type: Object, required: true },
  /** SPEC-138: no edit applies — every control that changes a rule is
   * disabled (the scope and the list still show the rules). */
  readOnly: { type: Boolean, default: false },
});

const m = computed(() => asConditionalFormatManager(props.manager));
const selected = computed(() => m.value.selected.value);

/** An input or select's value, as text. */
function valueOf(event = new Event("change")) {
  const el = event.target;
  return el instanceof HTMLInputElement || el instanceof HTMLSelectElement ? el.value : "";
}

function checked(event = new Event("change")) {
  const el = event.target;
  return el instanceof HTMLInputElement && el.checked;
}

/** Enter commits the "Applies to" field the way leaving it does. */
function blurTarget(event = new Event("keydown")) {
  const el = event.target;
  if (el instanceof HTMLInputElement) {
    el.blur();
  }
}

function onListKey(event = new KeyboardEvent("keydown")) {
  if (event.target instanceof HTMLInputElement) {
    return;
  }
  if (event.key === "ArrowDown" || event.key === "ArrowUp") {
    event.preventDefault();
    m.value.selectBy(event.key === "ArrowDown" ? 1 : -1);
  }
}

function togglePresets() {
  m.value.presetsOpen.value = !m.value.presetsOpen.value;
}

// The presets menu closes on a press anywhere outside it (or its button).
const addMenu = useTemplateRef("addMenu");
function onDocumentPointerDown(event = new Event("pointerdown")) {
  const root = addMenu.value;
  if (m.value.presetsOpen.value && root && event.target instanceof Node && !root.contains(event.target)) {
    m.value.presetsOpen.value = false;
  }
}
onMounted(() => {
  document.addEventListener("pointerdown", onDocumentPointerDown);
});
onBeforeUnmount(() => {
  document.removeEventListener("pointerdown", onDocumentPointerDown);
});
</script>

<template>
  <section
    class="cf-manager"
    data-testid="cf-manager"
    aria-label="Conditional formatting rules"
  >
    <ConditionalFormatEditor
      v-if="m.editor.open.value"
      :editor="m.editor"
    />
    <template v-else>
      <label class="cf-scope">Show rules for
        <select
          data-testid="cf-scope"
          :value="m.scope.value"
          @change="m.setScope(valueOf($event))"
        >
          <option
            v-for="scope in CF_SCOPES"
            :key="scope.id"
            :value="scope.id"
          >{{ scope.label }}</option>
        </select>
      </label>

      <div
        class="cf-actions"
        role="toolbar"
        aria-label="Rule actions"
      >
        <div
          ref="addMenu"
          class="cf-add"
        >
          <button
            type="button"
            class="cf-button"
            data-testid="cf-add"
            aria-haspopup="menu"
            :aria-expanded="m.presetsOpen.value"
            :disabled="props.readOnly || !m.hasSheet.value"
            @click="togglePresets"
          >
            <AppIcon name="plus" /> Add rule
          </button>
          <div
            v-if="m.presetsOpen.value"
            class="cf-presets"
            role="menu"
            aria-label="Add a rule"
            data-testid="cf-presets"
            @keydown.escape="m.presetsOpen.value = false"
          >
            <template
              v-for="group in CF_ADD_PRESET_GROUPS"
              :key="group.label"
            >
              <div
                v-if="group.label"
                class="cf-presets-group"
                role="presentation"
              >
                {{ group.label }}
              </div>
              <button
                v-for="preset in group.presets"
                :key="preset.testid"
                type="button"
                class="cf-preset"
                role="menuitem"
                :data-testid="preset.testid"
                @click="m.applyPreset(preset.testid)"
              >
                {{ preset.label }}
              </button>
            </template>
          </div>
        </div>
        <button
          type="button"
          class="cf-button"
          data-testid="cf-edit"
          :disabled="props.readOnly || !selected || selected.rule.kind === 'preserved'"
          @click="m.editSelected()"
        >
          Edit
        </button>
        <button
          type="button"
          class="cf-button"
          data-testid="cf-delete"
          :disabled="props.readOnly || !selected"
          @click="m.deleteSelected()"
        >
          Delete
        </button>
        <button
          type="button"
          class="cf-button"
          data-testid="cf-move-up"
          aria-label="Move rule up"
          :disabled="props.readOnly || !m.canMoveUp.value"
          @click="m.moveUp()"
        >
          <AppIcon name="chevron-up" />
        </button>
        <button
          type="button"
          class="cf-button"
          data-testid="cf-move-down"
          aria-label="Move rule down"
          :disabled="props.readOnly || !m.canMoveDown.value"
          @click="m.moveDown()"
        >
          <AppIcon name="chevron-down" />
        </button>
      </div>

      <ul
        class="cf-rule-list"
        role="listbox"
        aria-label="Rules, highest priority first"
        data-testid="cf-rule-list"
        tabindex="0"
        @keydown="onListKey"
      >
        <li
          v-for="entry in m.entries.value"
          :key="entry.rule.id"
          class="cf-rule"
          :class="{ selected: entry.rule.id === m.selectedRuleId.value }"
          role="option"
          data-testid="cf-rule"
          :data-rule-id="entry.rule.id"
          :aria-selected="entry.rule.id === m.selectedRuleId.value"
          @click="m.select(entry.rule.id)"
        >
          <div class="cf-rule-head">
            <CfRulePreview :preview="entry.preview" />
            <span
              class="cf-rule-description"
              data-testid="cf-rule-description"
            >{{ entry.description }}</span>
          </div>
          <label class="cf-rule-applies">Applies to
            <input
              type="text"
              data-testid="cf-rule-applies-to"
              :disabled="props.readOnly"
              :value="entry.appliesTo"
              :aria-invalid="Boolean(m.appliesToErrors.value[entry.format.id])"
              @change="m.setAppliesTo(entry.format.id, valueOf($event))"
              @keydown.enter.prevent="blurTarget"
            >
          </label>
          <span
            v-if="m.appliesToErrors.value[entry.format.id]"
            class="cf-error"
            role="alert"
            data-testid="cf-rule-applies-to-error"
          >{{ m.appliesToErrors.value[entry.format.id] }}</span>
          <label
            v-if="entry.formatting"
            class="cf-rule-stop"
          ><input
            type="checkbox"
            data-testid="cf-rule-stop-if-true"
            :disabled="props.readOnly"
            :checked="entry.rule.stopIfTrue === true"
            @change="m.setStop(entry.rule.id, checked($event))"
          > Stop if true</label>
        </li>
      </ul>
      <p
        v-if="m.entries.value.length === 0"
        class="cf-empty"
        data-testid="cf-empty"
      >
        No conditional formatting rules.
      </p>

      <div class="cf-clear">
        <button
          type="button"
          class="cf-button"
          data-testid="cf-clear-selection"
          :disabled="props.readOnly || !m.hasSheet.value"
          @click="m.clearSelection()"
        >
          Clear rules from selected cells
        </button>
        <button
          type="button"
          class="cf-button"
          data-testid="cf-clear-sheet"
          :disabled="props.readOnly || !m.hasSheet.value"
          @click="m.clearSheet()"
        >
          Clear rules from entire sheet
        </button>
      </div>
    </template>
  </section>
</template>

<style scoped>
.cf-manager {
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
  padding: var(--space-3);
  font-size: var(--text-md);
}
.cf-scope {
  display: flex;
  align-items: center;
  gap: var(--space-3);
}
.cf-actions {
  display: flex;
  flex-wrap: wrap;
  gap: var(--space-2);
}
.cf-add {
  position: relative;
}
.cf-button {
  display: inline-flex;
  align-items: center;
  gap: var(--space-2);
  padding: var(--space-2) var(--space-3);
  cursor: pointer;
}
.cf-presets {
  position: absolute;
  z-index: 20;
  top: 100%;
  left: 0;
  display: flex;
  flex-direction: column;
  min-width: 15rem;
  max-height: 24rem;
  overflow-y: auto;
  padding: var(--space-2) 0;
  background: var(--surface-raised);
  border: var(--border-width) solid var(--border-default);
  box-shadow: var(--shadow-2);
}
.cf-presets-group {
  padding: var(--space-2) var(--space-4);
  font-weight: var(--weight-bold);
  color: var(--text-secondary);
}
.cf-preset {
  padding: var(--space-2) var(--space-5);
  border: none;
  background: transparent;
  color: var(--text-primary);
  text-align: left;
  cursor: pointer;
}
.cf-preset:hover {
  background: var(--state-hover);
}
.cf-rule-list {
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
  margin: 0;
  padding: 0;
  list-style: none;
}
.cf-rule {
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
  padding: var(--space-2) var(--space-3);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
  cursor: pointer;
}
.cf-rule.selected {
  border-color: var(--accent);
  background: var(--surface-selected);
}
.cf-rule-head {
  display: flex;
  align-items: center;
  gap: var(--space-3);
}
.cf-rule-description {
  overflow-wrap: anywhere;
}
.cf-rule-applies {
  display: flex;
  align-items: center;
  gap: var(--space-2);
}
.cf-rule-applies input {
  flex: 1;
  min-width: 0;
}
.cf-rule-stop {
  display: flex;
  align-items: center;
  gap: var(--space-2);
}
.cf-error {
  color: var(--text-danger);
}
.cf-empty {
  margin: 0;
  color: var(--text-secondary);
}
.cf-clear {
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
}
</style>
