<script setup lang="ts">
import { computed } from "vue";
import { formatTimestamp } from "../lib/format";
import { asChangeEntries } from "../views/document/useTrackChanges";

// The track-changes review pane (SPEC-039): a flat, always-in-document-
// order list of changes with per-change Accept/Reject, Accept all/Reject
// all, Next/Previous navigation and the display-mode selector — a simpler
// docked list than CommentsPane's margin-anchored cards, since AC-7 only
// asks for a document-order list, not per-change anchored positioning.
// Runtime prop/emit declarations, not the generic form — espree parses
// .vue scripts and rejects TS-only syntax (see CommentsPane.vue).
const props = defineProps({
  changes: { type: Array, default: () => [] },
  displayMode: { type: String, default: "all" },
  recording: { type: Boolean, default: false },
  authorName: { type: String, default: "You" },
  editable: { type: Boolean, default: true },
});

const emit = defineEmits([
  "close",
  "toggle-recording",
  "update:author-name",
  "update:display-mode",
  "accept",
  "reject",
  "accept-all",
  "reject-all",
  "next",
  "previous",
  "focus-change",
]);

const changesTyped = computed(() => asChangeEntries(props.changes));

function onDisplayModeChange(event = new Event("change")) {
  const target = event.target;
  if (target instanceof HTMLSelectElement) {
    emit("update:display-mode", target.value);
  }
}

function onAuthorNameChange(event = new Event("change")) {
  const target = event.target;
  if (target instanceof HTMLInputElement) {
    emit("update:author-name", target.value);
  }
}
</script>

<template>
  <aside
    class="track-changes-pane"
    aria-label="Track changes"
    data-testid="track-changes-pane"
  >
    <div class="track-changes-pane-header">
      <span class="track-changes-pane-title">Track changes</span>
      <button
        type="button"
        class="track-changes-close"
        data-testid="track-changes-pane-close"
        aria-label="Close track changes pane"
        @click="emit('close')"
      >
        ×
      </button>
    </div>

    <div class="track-changes-controls">
      <button
        type="button"
        class="track-changes-toggle"
        :aria-pressed="recording"
        data-testid="track-changes-recording-toggle"
        @click="emit('toggle-recording')"
      >
        {{ recording ? "Recording: On" : "Recording: Off" }}
      </button>

      <label class="track-changes-field">
        Display
        <select
          class="track-changes-select"
          data-testid="track-changes-display-mode-select"
          :value="displayMode"
          @change="onDisplayModeChange"
        >
          <option value="all">
            All markup
          </option>
          <option value="simple">
            Simple markup
          </option>
          <option value="noMarkup">
            No markup
          </option>
          <option value="original">
            Original
          </option>
        </select>
      </label>

      <label class="track-changes-field">
        Author
        <input
          type="text"
          class="track-changes-author-input"
          data-testid="track-changes-author-input"
          :value="authorName"
          @change="onAuthorNameChange"
        >
      </label>
    </div>

    <div class="track-changes-nav">
      <button
        type="button"
        data-testid="track-changes-previous"
        :disabled="changesTyped.length === 0"
        @click="emit('previous')"
      >
        ◀ Previous
      </button>
      <button
        type="button"
        data-testid="track-changes-next"
        :disabled="changesTyped.length === 0"
        @click="emit('next')"
      >
        Next ▶
      </button>
    </div>

    <div
      v-if="editable"
      class="track-changes-bulk-actions"
    >
      <button
        type="button"
        data-testid="track-changes-accept-all"
        :disabled="changesTyped.length === 0"
        @click="emit('accept-all')"
      >
        Accept all
      </button>
      <button
        type="button"
        data-testid="track-changes-reject-all"
        :disabled="changesTyped.length === 0"
        @click="emit('reject-all')"
      >
        Reject all
      </button>
    </div>

    <p
      v-if="changesTyped.length === 0"
      class="track-changes-empty"
      data-testid="track-changes-empty"
    >
      No tracked changes
    </p>
    <ul
      v-else
      class="track-changes-list"
    >
      <li
        v-for="entry in changesTyped"
        :key="entry.changeId"
        class="track-change-item"
        :data-testid="`track-change-${entry.changeId}`"
        @click="emit('focus-change', entry.changeId)"
      >
        <div class="track-change-header">
          <span class="track-change-author">{{ entry.author }}</span>
          <span class="track-change-timestamp">{{ formatTimestamp(new Date(entry.date)) }}</span>
        </div>
        <div class="track-change-description">
          {{ entry.description }}
        </div>
        <p
          v-if="entry.preview"
          class="track-change-preview"
        >
          “{{ entry.preview }}”
        </p>
        <div
          v-if="editable"
          class="track-change-actions"
        >
          <button
            type="button"
            data-testid="track-change-accept"
            @click.stop="emit('accept', entry.changeId)"
          >
            Accept
          </button>
          <button
            type="button"
            data-testid="track-change-reject"
            @click.stop="emit('reject', entry.changeId)"
          >
            Reject
          </button>
        </div>
      </li>
    </ul>
  </aside>
</template>

<style scoped>
.track-changes-pane {
  flex: none;
  align-self: flex-start;
  position: sticky;
  top: 0;
  width: 18rem;
  max-height: 100vh;
  display: flex;
  flex-direction: column;
  min-height: 0;
  overflow-y: auto;
  background: var(--surface-raised);
  border-left: var(--border-width) solid var(--border-default);
  box-sizing: border-box;
  font-size: var(--text-md);
}

.track-changes-pane-header {
  flex: none;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--space-3);
  padding: var(--space-3) var(--space-4);
}

.track-changes-pane-title {
  font-size: var(--text-sm);
  font-weight: var(--weight-bold);
  letter-spacing: 0.04em;
  text-transform: uppercase;
  color: var(--text-secondary);
}

.track-changes-close {
  border: none;
  background: transparent;
  font-size: var(--text-lg);
  line-height: 1;
  cursor: pointer;
  color: var(--text-secondary);
}

.track-changes-controls {
  flex: none;
  display: flex;
  flex-wrap: wrap;
  gap: var(--space-3);
  padding: 0 var(--space-4) var(--space-3);
  align-items: flex-end;
}

.track-changes-toggle {
  border: var(--border-width) solid var(--border-default);
  background: var(--surface-base);
  border-radius: var(--radius-sm);
  padding: var(--space-2) var(--space-3);
  font-size: var(--text-xs);
  cursor: pointer;
}

.track-changes-toggle[aria-pressed="true"] {
  background: var(--surface-sunken);
  color: var(--success);
  border-color: var(--success);
}

.track-changes-field {
  display: flex;
  flex-direction: column;
  gap: var(--space-1);
  font-size: var(--text-xs);
  color: var(--text-secondary);
}

.track-changes-select,
.track-changes-author-input {
  font-size: var(--text-sm);
  padding: var(--space-1) var(--space-2);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
}

.track-changes-nav,
.track-changes-bulk-actions {
  flex: none;
  display: flex;
  gap: var(--space-3);
  padding: 0 var(--space-4) var(--space-3);
}

.track-changes-nav button,
.track-changes-bulk-actions button {
  border: var(--border-width) solid var(--border-default);
  background: var(--surface-base);
  border-radius: var(--radius-sm);
  padding: var(--space-2) var(--space-3);
  font-size: var(--text-xs);
  cursor: pointer;
}

.track-changes-nav button:disabled,
.track-changes-bulk-actions button:disabled {
  opacity: 0.5;
  cursor: default;
}

.track-changes-empty {
  margin: 0;
  padding: var(--space-2) var(--space-4) var(--space-4);
  font-style: italic;
  color: var(--text-secondary);
}

.track-changes-list {
  list-style: none;
  margin: 0;
  padding: 0 var(--space-3) var(--space-3);
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
}

.track-change-item {
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
  padding: var(--space-3);
  cursor: pointer;
}

.track-change-header {
  display: flex;
  justify-content: space-between;
  gap: var(--space-3);
  font-size: var(--text-xs);
  color: var(--text-secondary);
  margin-bottom: var(--space-2);
}

.track-change-author {
  font-weight: var(--weight-bold);
  color: var(--text-primary);
}

.track-change-description {
  font-weight: var(--weight-bold);
  color: var(--text-primary);
}

.track-change-preview {
  margin: var(--space-2) 0 0;
  white-space: pre-wrap;
  overflow-wrap: anywhere;
  color: var(--text-secondary);
  font-style: italic;
}

.track-change-actions {
  display: flex;
  gap: var(--space-3);
  margin-top: var(--space-3);
}

.track-change-actions button {
  border: var(--border-width) solid var(--border-default);
  background: var(--surface-base);
  border-radius: var(--radius-sm);
  padding: var(--space-1) var(--space-3);
  font-size: var(--text-xs);
  cursor: pointer;
}
</style>
