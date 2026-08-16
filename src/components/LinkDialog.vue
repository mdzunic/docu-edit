<script setup lang="ts">
import { computed, ref } from "vue";
import { asLinkTargetEntries, isDangerousUrl, sanitizeHref } from "../lib/links";
import BaseDialog from "./BaseDialog.vue";

// The link dialog (SPEC-025 AC-1..3, AC-14..16, AC-19..21): display text, a
// web address, and — as the alternative target — the document's headings and
// bookmarks, listed together. Owns URL-scheme validation (dangerous schemes
// rejected inline); the parent owns applying the link to the model.
// BaseDialog owns the shell, Escape and initial focus (the [autofocus] URL
// input). Runtime prop declarations, not defineProps<{...}>() generics — see
// AppShell.vue's comment on the same espree constraint.
const props = defineProps({
  /** Prefills the text field (the selected text; "" without a selection). */
  initialText: { type: String, default: "" },
  /** Prefills the URL field (editing an existing external link). */
  initialUrl: { type: String, default: "" },
  /** Preselects an internal target (editing an existing internal link). */
  initialAnchor: { type: String, default: "" },
  /** Headings offered as targets: [{ anchor, text, level }] (AC-19, AC-21). */
  headings: { type: Array, default: () => [] },
  /** Bookmarks offered as targets: [{ anchor, text }] (AC-20, AC-21). */
  bookmarks: { type: Array, default: () => [] },
});

const emit = defineEmits(["save", "cancel"]);

const draftText = ref(props.initialText);
const draftUrl = ref(props.initialUrl);
const selectedAnchor = ref(props.initialAnchor);

const trimmedUrl = computed(() => draftUrl.value.trim());

// AC-14..16: a dangerous scheme is rejected on entry — inline error, save
// blocked. An anchor selection bypasses the URL entirely.
const error = computed(() => {
  if (selectedAnchor.value !== "" || trimmedUrl.value === "") {
    return "";
  }
  return isDangerousUrl(trimmedUrl.value)
    ? "This address type isn't allowed in links"
    : "";
});

const canSave = computed(() => {
  if (selectedAnchor.value !== "") {
    return true;
  }
  return trimmedUrl.value !== "" && !error.value;
});

const anchorTargets = computed(() => {
  const headings = asLinkTargetEntries(props.headings).map((entry) => ({
    anchor: entry.anchor,
    label: entry.text || "(empty heading)",
  }));
  const bookmarks = asLinkTargetEntries(props.bookmarks).map((entry) => ({
    anchor: entry.anchor,
    label: entry.text || entry.anchor,
  }));
  return { headings, bookmarks };
});

function chooseAnchor(anchor = "") {
  selectedAnchor.value = selectedAnchor.value === anchor ? "" : anchor;
  if (selectedAnchor.value !== "") {
    draftUrl.value = "";
  }
}

function handleUrlInput() {
  if (draftUrl.value.trim() !== "") {
    selectedAnchor.value = "";
  }
}

function fallbackLabel() {
  if (selectedAnchor.value === "") {
    return trimmedUrl.value;
  }
  const all = [...anchorTargets.value.headings, ...anchorTargets.value.bookmarks];
  const chosen = all.find((target) => target.anchor === selectedAnchor.value);
  return chosen ? chosen.label : selectedAnchor.value;
}

function save() {
  if (!canSave.value) {
    return;
  }
  const text = draftText.value.trim() !== "" ? draftText.value : fallbackLabel();
  if (selectedAnchor.value !== "") {
    emit("save", { text, anchor: selectedAnchor.value });
    return;
  }
  const href = sanitizeHref(trimmedUrl.value);
  if (href === null) {
    return;
  }
  emit("save", { text, href });
}
</script>

<template>
  <BaseDialog
    title="Link"
    testid="link-dialog"
    width="24rem"
    @dismiss="emit('cancel')"
  >
    <label
      class="dialog-label"
      for="link-text-input"
    >Text</label>
    <input
      id="link-text-input"
      v-model="draftText"
      class="dialog-input"
      data-testid="link-text-input"
      @keydown.enter="save"
    >
    <label
      class="dialog-label"
      for="link-url-input"
    >Web address</label>
    <input
      id="link-url-input"
      v-model="draftUrl"
      class="dialog-input"
      data-testid="link-url-input"
      placeholder="https://"
      autofocus
      @input="handleUrlInput"
      @keydown.enter="save"
    >
    <p
      v-if="error"
      class="dialog-error"
      data-testid="link-error"
    >
      {{ error }}
    </p>
    <div
      class="target-section"
      data-testid="link-target-list"
    >
      <p class="dialog-label">
        Or link to this document
      </p>
      <p
        v-if="anchorTargets.headings.length === 0 && anchorTargets.bookmarks.length === 0"
        class="target-empty"
      >
        No headings or bookmarks yet
      </p>
      <template v-if="anchorTargets.headings.length > 0">
        <p class="target-group">
          Headings
        </p>
        <button
          v-for="target in anchorTargets.headings"
          :key="target.anchor"
          type="button"
          class="target-option"
          :class="{ selected: target.anchor === selectedAnchor }"
          :data-testid="`link-target-${target.anchor}`"
          :aria-pressed="target.anchor === selectedAnchor ? 'true' : 'false'"
          @click="chooseAnchor(target.anchor)"
        >
          {{ target.label }}
        </button>
      </template>
      <template v-if="anchorTargets.bookmarks.length > 0">
        <p class="target-group">
          Bookmarks
        </p>
        <button
          v-for="target in anchorTargets.bookmarks"
          :key="target.anchor"
          type="button"
          class="target-option"
          :class="{ selected: target.anchor === selectedAnchor }"
          :data-testid="`link-target-${target.anchor}`"
          :aria-pressed="target.anchor === selectedAnchor ? 'true' : 'false'"
          @click="chooseAnchor(target.anchor)"
        >
          {{ target.label }}
        </button>
      </template>
    </div>
    <template #actions>
      <button
        type="button"
        data-testid="link-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </button>
      <button
        type="button"
        class="primary"
        data-testid="link-save"
        :disabled="!canSave"
        @click="save"
      >
        Apply
      </button>
    </template>
  </BaseDialog>
</template>

<style scoped>
.dialog-label {
  display: block;
  margin: 0.5rem 0 0.25rem;
  font-size: 0.8125rem;
  color: #5f6368;
}

.dialog-input {
  font: inherit;
  width: 100%;
  box-sizing: border-box;
  padding: 0.375rem 0.5rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
}

.dialog-error {
  margin: 0.5rem 0 0;
  color: #b3261e;
  font-size: 0.8125rem;
}

.target-section {
  margin-top: 0.75rem;
  max-height: 12rem;
  overflow-y: auto;
  border: 1px solid #e0e0e0;
  border-radius: 4px;
  padding: 0.5rem;
}

.target-group {
  margin: 0.375rem 0 0.125rem;
  font-size: 0.75rem;
  font-weight: 600;
  text-transform: uppercase;
  letter-spacing: 0.03em;
  color: #5f6368;
}

.target-empty {
  margin: 0.25rem 0;
  font-size: 0.8125rem;
  color: #5f6368;
}

.target-option {
  display: block;
  width: 100%;
  text-align: start;
  font: inherit;
  font-size: 0.875rem;
  padding: 0.25rem 0.5rem;
  margin: 0.125rem 0;
  border: 1px solid transparent;
  border-radius: 4px;
  background: transparent;
  color: #202124;
  cursor: pointer;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.target-option:hover {
  background: #f1f3f4;
}

.target-option.selected {
  background: #e8f0fe;
  border-color: #aecbfa;
  color: #174ea6;
}

.dialog-actions button {
  font: inherit;
  padding: 0.375rem 1rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
}

.dialog-actions button.primary {
  background: #2c3e50;
  border-color: #2c3e50;
  color: #fff;
}

.dialog-actions button:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}
</style>
