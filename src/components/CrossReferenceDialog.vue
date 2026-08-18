<script setup lang="ts">
// "Insert → Cross-reference" (SPEC-053 AC-21..25): lists the document's
// headings (SPEC-019) and bookmarks (SPEC-025) as targets, and inserts a
// reference showing the target's text, page number or above/below position.
// The parent resolves a heading target to a (hidden) bookmark and computes the
// cached value. Runtime prop/emit declarations per the espree constraint.
import { computed, ref } from "vue";
import { asCrossRefBookmarks, asCrossRefCaptions, asCrossRefHeadings } from "../lib/fields";
import BaseDialog from "./BaseDialog.vue";

const props = defineProps({
  /** Heading targets: `[{ index, text, level }]` in document order. */
  headings: { type: Array, default: () => [] },
  /** Bookmark targets: `[{ name }]`, hidden bookmarks excluded. */
  bookmarks: { type: Array, default: () => [] },
  /** Caption targets: `[{ anchor, label, text }]` (SPEC-055 AC-23). */
  captions: { type: Array, default: () => [] },
});

const emit = defineEmits(["insert", "cancel"]);

const headings = computed(() => asCrossRefHeadings(props.headings));
const bookmarks = computed(() => asCrossRefBookmarks(props.bookmarks));
const captions = computed(() => asCrossRefCaptions(props.captions));

// The distinct caption labels, each its own category (SPEC-055 AC-23).
const captionLabels = computed(() => Array.from(new Set(captions.value.map((caption) => caption.label))));

// "text"/"position"/the caption formats insert a `ref`; "page" a `pageRef`.
const show = ref("text");
// The selected target: "heading:<index>", "bookmark:<name>" or "caption:<anchor>".
const target = ref(
  headings.value.length > 0
    ? `heading:${headings.value[0].index}`
    : bookmarks.value.length > 0
      ? `bookmark:${bookmarks.value[0].name}`
      : captions.value.length > 0
        ? `caption:${captions.value[0].anchor}`
        : "",
);

/** A caption target offers the six caption formats (SPEC-055 AC-24); a
 * heading/bookmark keeps SPEC-053's three. */
const isCaptionTarget = computed(() => target.value.startsWith("caption:"));

function captionsFor(label = "") {
  return captions.value.filter((caption) => caption.label === label);
}

function insert() {
  if (!target.value) {
    return;
  }
  const [type, id] = [target.value.slice(0, target.value.indexOf(":")), target.value.slice(target.value.indexOf(":") + 1)];
  emit("insert", { targetType: type, targetId: id, show: show.value });
}
</script>

<template>
  <BaseDialog
    title="Cross-reference"
    testid="cross-reference-dialog"
    width="26rem"
    @dismiss="emit('cancel')"
  >
    <label class="row">
      <span class="row-label">Insert reference to</span>
      <select
        v-model="show"
        class="row-input"
        data-testid="cross-reference-show"
      >
        <template v-if="isCaptionTarget">
          <option value="full">
            Entire caption
          </option>
          <option value="labelAndNumber">
            Label and number
          </option>
          <option value="numberOnly">
            Number only
          </option>
          <option value="textOnly">
            Caption text only
          </option>
          <option value="page">
            Page number
          </option>
          <option value="position">
            Above/below
          </option>
        </template>
        <template v-else>
          <option value="text">
            Text
          </option>
          <option value="page">
            Page number
          </option>
          <option value="position">
            Above/below
          </option>
        </template>
      </select>
    </label>

    <p class="hint">
      For which item:
    </p>
    <div
      class="target-list"
      role="listbox"
      aria-label="Cross-reference targets"
      data-testid="cross-reference-targets"
    >
      <template v-if="headings.length > 0">
        <p class="group-label">
          Headings
        </p>
        <label
          v-for="heading in headings"
          :key="`heading:${heading.index}`"
          class="target-row"
        >
          <input
            v-model="target"
            type="radio"
            name="cross-reference-target"
            :value="`heading:${heading.index}`"
            :data-testid="`cross-reference-heading-${heading.index}`"
          >
          <span :style="{ paddingLeft: `${(heading.level - 1) * 0.75}rem` }">{{ heading.text }}</span>
        </label>
      </template>
      <template v-if="bookmarks.length > 0">
        <p class="group-label">
          Bookmarks
        </p>
        <label
          v-for="bookmark in bookmarks"
          :key="`bookmark:${bookmark.name}`"
          class="target-row"
        >
          <input
            v-model="target"
            type="radio"
            name="cross-reference-target"
            :value="`bookmark:${bookmark.name}`"
            :data-testid="`cross-reference-bookmark-${bookmark.name}`"
          >
          <span>{{ bookmark.name }}</span>
        </label>
      </template>
      <template
        v-for="label in captionLabels"
        :key="`caption-group:${label}`"
      >
        <p class="group-label">
          {{ label }}
        </p>
        <label
          v-for="caption in captionsFor(label)"
          :key="`caption:${caption.anchor}`"
          class="target-row"
        >
          <input
            v-model="target"
            type="radio"
            name="cross-reference-target"
            :value="`caption:${caption.anchor}`"
            :data-testid="`cross-reference-caption-${caption.anchor}`"
          >
          <span>{{ caption.text }}</span>
        </label>
      </template>
      <p
        v-if="headings.length === 0 && bookmarks.length === 0 && captions.length === 0"
        class="empty"
        data-testid="cross-reference-empty"
      >
        This document has no headings, bookmarks or captions to reference yet.
      </p>
    </div>

    <template #actions>
      <button
        type="button"
        data-testid="cross-reference-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </button>
      <button
        type="button"
        class="primary"
        data-testid="cross-reference-insert"
        :disabled="!target"
        @click="insert"
      >
        Insert
      </button>
    </template>
  </BaseDialog>
</template>

<style scoped>
.row {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--space-4);
  font-size: var(--text-lg);
  margin-bottom: var(--space-4);
}

.row-label {
  color: var(--text-secondary);
}

.row-input {
  font: inherit;
  min-width: 12rem;
}

.hint {
  margin: 0 0 var(--space-2);
  font-size: var(--text-md);
  color: var(--text-secondary);
}

.target-list {
  max-height: 16rem;
  overflow-y: auto;
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  padding: var(--space-2) 0;
}

.group-label {
  margin: var(--space-2) var(--space-3) var(--space-1);
  font-size: var(--text-sm);
  font-weight: var(--weight-bold);
  text-transform: uppercase;
  color: var(--text-disabled);
}

.target-row {
  display: flex;
  align-items: center;
  gap: var(--space-3);
  padding: var(--space-2) var(--space-3);
  cursor: pointer;
  font-size: var(--text-lg);
}

.target-row:hover {
  background: var(--surface-canvas);
}

.empty {
  margin: var(--space-3) var(--space-3);
  font-size: var(--text-md);
  color: var(--text-secondary);
}

.dialog-actions button {
  font: inherit;
  padding: var(--space-3) var(--space-5);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  cursor: pointer;
}

.dialog-actions button.primary {
  background: var(--surface-brand);
  border-color: var(--surface-brand);
  color: var(--text-on-brand);
}

.dialog-actions button:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}
</style>
