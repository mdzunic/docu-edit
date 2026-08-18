<script setup lang="ts">
// The find (and replace) bar (SPEC-017). Owns only its own inputs and the
// keyboard affordances local to it (Enter/Shift+Enter step, Escape closes,
// Ctrl/Cmd+H reveals replace); the editor owns the matches, the highlight
// overlay, navigation state and every model edit, so this component is a
// controlled view — every value comes in as a prop and every intent leaves as
// an event. Runtime prop declarations, not defineProps<{...}>() generics —
// see AppShell.vue's note on the same espree constraint.
import { nextTick, useTemplateRef, watch } from "vue";

const props = defineProps({
  visible: { type: Boolean, default: false },
  /** Whether the replace row is revealed (AC: Ctrl/Cmd+H or the expand
   * toggle). */
  expanded: { type: Boolean, default: false },
  query: { type: String, default: "" },
  replaceText: { type: String, default: "" },
  matchCase: { type: Boolean, default: false },
  wholeWords: { type: Boolean, default: false },
  useRegex: { type: Boolean, default: false },
  /** Total matches in the document, and the 0-based index of the current one
   * (-1 when there is none), rendered as the "N of M" count. */
  matchCount: { type: Number, default: 0 },
  currentIndex: { type: Number, default: -1 },
  /** An inline error (an invalid regular expression). Non-empty disables the
   * search actions. */
  error: { type: String, default: "" },
  /** A transient hint — "Wrapped to top", "No results" — shown beside the
   * count. */
  hint: { type: String, default: "" },
});

const emit = defineEmits([
  "update:query",
  "update:replaceText",
  "update:matchCase",
  "update:wholeWords",
  "update:useRegex",
  "toggle-expand",
  "next",
  "prev",
  "replace",
  "replace-all",
  "close",
]);

// useTemplateRef infers the element type from the matching `ref` attribute,
// so no generic is needed — .vue scripts are parsed by espree, which rejects
// TS-only syntax (see AppShell.vue's note).
const queryInput = useTemplateRef("queryInput");
const replaceInput = useTemplateRef("replaceInput");

// Focus the query field whenever the bar opens (AC: the caret lands in the
// query field ready to type), and the replace field the moment it is revealed
// (Ctrl/Cmd+H / the expand toggle), so the user can type the replacement
// straight away.
watch(
  () => props.visible,
  (visible) => {
    if (visible) {
      void nextTick(() => {
        queryInput.value?.focus();
        queryInput.value?.select();
      });
    }
  },
  { immediate: true },
);

watch(
  () => props.expanded,
  (expanded) => {
    if (expanded && props.visible) {
      void nextTick(() => replaceInput.value?.focus());
    }
  },
);

function countLabel() {
  const position = props.currentIndex >= 0 ? props.currentIndex + 1 : 0;
  return `${position} of ${props.matchCount}`;
}

function hasMatches() {
  return props.matchCount > 0 && props.error === "";
}

function canReplaceCurrent() {
  return hasMatches() && props.currentIndex >= 0;
}

// "Replace all" stays enabled with zero matches so it can report "Replaced 0
// occurrences" (AC), but not for an empty query or an invalid regex.
function canReplaceAll() {
  return props.query !== "" && props.error === "";
}

/** The secondary line beside the count: the error wins, then any transient
 * hint, then a "No results" note for a non-empty query that matched nothing. */
function statusText() {
  if (props.error !== "") {
    return props.error;
  }
  if (props.hint !== "") {
    return props.hint;
  }
  if (props.query !== "" && props.matchCount === 0) {
    return "No results";
  }
  return "";
}

function onQueryInput(event = new Event("input")) {
  const target = event.target;
  if (target instanceof HTMLInputElement) {
    emit("update:query", target.value);
  }
}

function onReplaceInput(event = new Event("input")) {
  const target = event.target;
  if (target instanceof HTMLInputElement) {
    emit("update:replaceText", target.value);
  }
}

function onQueryKeydown(event = new KeyboardEvent("keydown")) {
  if (event.key === "Enter") {
    event.preventDefault();
    emit(event.shiftKey ? "prev" : "next");
  } else if (event.key === "Escape") {
    event.preventDefault();
    emit("close");
  }
}

function onReplaceKeydown(event = new KeyboardEvent("keydown")) {
  if (event.key === "Enter") {
    event.preventDefault();
    emit("replace");
  } else if (event.key === "Escape") {
    event.preventDefault();
    emit("close");
  }
}
</script>

<template>
  <div
    v-if="visible"
    class="find-bar"
    data-testid="find-bar"
    role="search"
    aria-label="Find and replace"
  >
    <div class="find-row">
      <button
        type="button"
        class="find-icon-button expand-toggle"
        data-testid="find-expand"
        :aria-expanded="expanded ? 'true' : 'false'"
        :aria-label="expanded ? 'Hide replace' : 'Show replace'"
        :title="expanded ? 'Hide replace' : 'Show replace'"
        @click="emit('toggle-expand')"
      >
        {{ expanded ? "▾" : "▸" }}
      </button>
      <input
        ref="queryInput"
        class="find-input"
        data-testid="find-query"
        type="text"
        aria-label="Find"
        placeholder="Find in document"
        :value="query"
        @input="onQueryInput"
        @keydown="onQueryKeydown"
      >
      <span
        class="find-count"
        data-testid="find-count"
        aria-live="polite"
      >{{ countLabel() }}</span>
      <span
        v-if="statusText() !== ''"
        class="find-status"
        data-testid="find-status"
        :role="error !== '' ? 'alert' : 'status'"
      >{{ statusText() }}</span>
      <button
        type="button"
        class="find-icon-button"
        data-testid="find-prev"
        aria-label="Previous match"
        title="Previous match (Shift+Enter)"
        :disabled="!hasMatches()"
        @click="emit('prev')"
      >
        ↑
      </button>
      <button
        type="button"
        class="find-icon-button"
        data-testid="find-next"
        aria-label="Next match"
        title="Next match (Enter)"
        :disabled="!hasMatches()"
        @click="emit('next')"
      >
        ↓
      </button>
      <button
        type="button"
        class="find-icon-button"
        data-testid="find-close"
        aria-label="Close find bar"
        title="Close (Escape)"
        @click="emit('close')"
      >
        ✕
      </button>
    </div>

    <div class="find-options">
      <button
        type="button"
        class="find-option"
        data-testid="find-match-case"
        :aria-pressed="matchCase ? 'true' : 'false'"
        title="Match case"
        @click="emit('update:matchCase', !matchCase)"
      >
        Match case
      </button>
      <button
        type="button"
        class="find-option"
        data-testid="find-whole-words"
        :aria-pressed="wholeWords ? 'true' : 'false'"
        title="Whole words only"
        @click="emit('update:wholeWords', !wholeWords)"
      >
        Whole words
      </button>
      <button
        type="button"
        class="find-option"
        data-testid="find-regex"
        :aria-pressed="useRegex ? 'true' : 'false'"
        title="Use regular expressions"
        @click="emit('update:useRegex', !useRegex)"
      >
        Regex
      </button>
      <span
        class="find-scope"
        data-testid="find-scope"
        title="Search covers body text only. Headers, footers and footnotes are excluded."
      >Body text only</span>
    </div>

    <div
      v-if="expanded"
      class="find-row replace-row"
      data-testid="find-replace-row"
    >
      <input
        ref="replaceInput"
        class="find-input"
        data-testid="find-replace-input"
        type="text"
        aria-label="Replace with"
        placeholder="Replace with"
        :value="replaceText"
        @input="onReplaceInput"
        @keydown="onReplaceKeydown"
      >
      <button
        type="button"
        class="find-text-button"
        data-testid="find-replace"
        :disabled="!canReplaceCurrent()"
        @click="emit('replace')"
      >
        Replace
      </button>
      <button
        type="button"
        class="find-text-button"
        data-testid="find-replace-all"
        :disabled="!canReplaceAll()"
        @click="emit('replace-all')"
      >
        Replace all
      </button>
    </div>
  </div>
</template>

<style scoped>
.find-bar {
  position: absolute;
  top: 0.5rem;
  right: 1.5rem;
  z-index: 20;
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
  padding: var(--space-3) var(--space-3);
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
  box-shadow: var(--shadow-2);
  font-size: var(--text-md);
}

.find-row {
  display: flex;
  align-items: center;
  gap: var(--space-3);
}

.find-input {
  font: inherit;
  font-size: var(--text-md);
  padding: var(--space-2) var(--space-3);
  width: 15rem;
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
}

.find-input:focus {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: calc(-1 * var(--border-width));
}

.find-count {
  min-width: 3.5rem;
  color: var(--text-secondary);
  white-space: nowrap;
  text-align: right;
}

.find-status {
  color: var(--text-danger);
  white-space: nowrap;
}

.find-icon-button {
  font: inherit;
  min-width: 1.75rem;
  height: 1.75rem;
  padding: 0 var(--space-3);
  border: var(--border-width) solid transparent;
  border-radius: var(--radius-sm);
  background: transparent;
  color: var(--text-primary);
  cursor: pointer;
}

.find-icon-button:hover:not(:disabled) {
  background: var(--surface-sunken);
}

.find-icon-button:disabled {
  color: var(--text-disabled);
  cursor: default;
}

.find-options {
  display: flex;
  align-items: center;
  gap: var(--space-3);
}

.find-option {
  font: inherit;
  font-size: var(--text-sm);
  padding: var(--space-1) var(--space-3);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  color: var(--text-primary);
  cursor: pointer;
}

.find-option[aria-pressed="true"] {
  background: var(--surface-selected);
  border-color: var(--accent-subtle);
  color: var(--text-accent);
}

.find-scope {
  margin-left: auto;
  color: var(--text-secondary);
  cursor: help;
}

.find-text-button {
  font: inherit;
  font-size: var(--text-md);
  padding: var(--space-2) var(--space-4);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  color: var(--text-primary);
  cursor: pointer;
}

.find-text-button:disabled {
  color: var(--text-disabled);
  cursor: default;
}
</style>
