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
  background: #f8f9fa;
  border-left: 1px solid #dadce0;
  box-sizing: border-box;
  font-size: 0.8125rem;
}

.track-changes-pane-header {
  flex: none;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 0.5rem;
  padding: 0.5rem 0.75rem;
}

.track-changes-pane-title {
  font-size: 0.75rem;
  font-weight: 600;
  letter-spacing: 0.04em;
  text-transform: uppercase;
  color: #5f6368;
}

.track-changes-close {
  border: none;
  background: transparent;
  font-size: 1rem;
  line-height: 1;
  cursor: pointer;
  color: #5f6368;
}

.track-changes-controls {
  flex: none;
  display: flex;
  flex-wrap: wrap;
  gap: 0.5rem;
  padding: 0 0.75rem 0.5rem;
  align-items: flex-end;
}

.track-changes-toggle {
  border: 1px solid #dadce0;
  background: #fff;
  border-radius: 4px;
  padding: 0.25rem 0.5rem;
  font-size: 0.7rem;
  cursor: pointer;
}

.track-changes-toggle[aria-pressed="true"] {
  background: #e6f4ea;
  color: #188038;
  border-color: #188038;
}

.track-changes-field {
  display: flex;
  flex-direction: column;
  gap: 0.15rem;
  font-size: 0.7rem;
  color: #5f6368;
}

.track-changes-select,
.track-changes-author-input {
  font-size: 0.75rem;
  padding: 0.15rem 0.3rem;
  border: 1px solid #dadce0;
  border-radius: 4px;
}

.track-changes-nav,
.track-changes-bulk-actions {
  flex: none;
  display: flex;
  gap: 0.4rem;
  padding: 0 0.75rem 0.5rem;
}

.track-changes-nav button,
.track-changes-bulk-actions button {
  border: 1px solid #dadce0;
  background: #fff;
  border-radius: 4px;
  padding: 0.2rem 0.5rem;
  font-size: 0.72rem;
  cursor: pointer;
}

.track-changes-nav button:disabled,
.track-changes-bulk-actions button:disabled {
  opacity: 0.5;
  cursor: default;
}

.track-changes-empty {
  margin: 0;
  padding: 0.25rem 0.75rem 0.75rem;
  font-style: italic;
  color: #5f6368;
}

.track-changes-list {
  list-style: none;
  margin: 0;
  padding: 0 0.5rem 0.5rem;
  display: flex;
  flex-direction: column;
  gap: 0.5rem;
}

.track-change-item {
  background: #fff;
  border: 1px solid #dadce0;
  border-radius: 6px;
  padding: 0.5rem;
  cursor: pointer;
}

.track-change-header {
  display: flex;
  justify-content: space-between;
  gap: 0.5rem;
  font-size: 0.7rem;
  color: #5f6368;
  margin-bottom: 0.2rem;
}

.track-change-author {
  font-weight: 600;
  color: #202124;
}

.track-change-description {
  font-weight: 600;
  color: #202124;
}

.track-change-preview {
  margin: 0.2rem 0 0;
  white-space: pre-wrap;
  overflow-wrap: anywhere;
  color: #5f6368;
  font-style: italic;
}

.track-change-actions {
  display: flex;
  gap: 0.4rem;
  margin-top: 0.4rem;
}

.track-change-actions button {
  border: 1px solid #dadce0;
  background: #fff;
  border-radius: 4px;
  padding: 0.1rem 0.45rem;
  font-size: 0.7rem;
  cursor: pointer;
}
</style>
