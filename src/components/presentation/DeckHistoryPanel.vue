<script setup lang="ts">
// The deck's version history panel (SPEC-103) — SPEC-040's pane for slides,
// mounted in the stage column under the stage like the other below-stage
// panels. Newest first, each version with its time, author (the
// preferences' authorName), slide count and slide-count change, and the
// actions: Preview, Restore, Download (.pptx/.odp), Name/Remove name and
// Delete; a named-only filter; the paused notice. Preview draws the
// reconstructed deck read-only on its own SlideStage with its own slide strip
// — the live deck and its stage are never swapped, and there is no diff
// highlighting.
//
// SPEC-085/086's panel contract: all logic lives in the useDeckHistory
// controller passed in, this component emits nothing, and its typed
// transforms live in deckHistoryRender.ts. Runtime prop declarations, not
// defineProps<{...}>() generics (espree lints .vue scripts).
import { computed } from "vue";
import SlideStage from "./SlideStage.vue";
import UiButton from "../ui/UiButton.vue";
import {
  asDeckHistory,
  asVersionEntries,
  downloadLabel,
  previewFields,
  previewScale,
  previewSlide,
  previewTitle,
  RESTORE_UNDO_NOTE,
  slideCountLabel,
  slideDeltaLabel,
  triggerLabel,
  versionTime,
} from "./deckHistoryRender";

const props = defineProps({
  /** The useDeckHistory controller (useDeckSaving's `history`). */
  controller: { type: Object, required: true },
  slideNumber: { type: Number, default: 1 },
});

const c = asDeckHistory(props.controller);
const entries = computed(() => asVersionEntries(c.visibleVersions.value));
const preview = computed(() => c.previewModel.value);
const shownSlide = computed(() => previewSlide(preview.value, c.previewIndex.value));

function onNamedOnly(event = new Event("change")) {
  const el = event.target;
  if (el instanceof HTMLInputElement) c.namedOnly.value = el.checked;
}

function onName(versionId = "", event = new Event("submit")) {
  const form = event.target;
  const input = form instanceof HTMLFormElement ? form.elements.namedItem("name") : null;
  if (input instanceof HTMLInputElement) void c.nameVersion(versionId, input.value);
}
</script>

<template>
  <section
    class="history-panel"
    data-testid="slides-history-panel"
    :aria-label="`Version history (slide ${slideNumber})`"
  >
    <header class="history-head">
      <h2 class="history-title">
        Version history
      </h2>
      <label class="history-filter">
        <input
          type="checkbox"
          data-testid="history-named-only"
          :checked="c.namedOnly.value"
          @change="onNamedOnly"
        >
        Named versions only
      </label>
    </header>
    <p class="history-note">
      {{ RESTORE_UNDO_NOTE }}
    </p>
    <p
      v-if="c.pausedMessage.value"
      class="history-paused"
      role="status"
      data-testid="history-paused-notice"
    >
      {{ c.pausedMessage.value }}
    </p>

    <div
      v-if="preview"
      class="history-preview"
      data-testid="history-preview-banner"
    >
      <div class="history-preview-head">
        <span>Previewing {{ previewTitle(c.versions.value, c.previewVersionId.value) }} — read-only</span>
        <UiButton
          size="sm"
          data-testid="history-exit-preview"
          @click="c.exitPreview()"
        >
          Back to current deck
        </UiButton>
      </div>
      <SlideStage
        v-if="shownSlide"
        class="history-preview-stage"
        :slide="shownSlide"
        :width-cm="preview.slideWidthCm"
        :height-cm="preview.slideHeightCm"
        :media="preview.media ?? []"
        :scale="previewScale(preview)"
        :interactive="false"
        :fields="previewFields(preview, c.previewIndex.value)"
        :theme="preview.theme"
      />
      <div
        class="history-preview-strip"
        role="group"
        aria-label="Preview slides"
      >
        <button
          v-for="(slide, i) in preview.slides"
          :key="slide.id"
          type="button"
          class="history-preview-slide"
          :aria-pressed="i === c.previewIndex.value ? 'true' : 'false'"
          :aria-label="`Preview slide ${i + 1}`"
          @click="c.selectPreviewSlide(i)"
        >
          {{ i + 1 }}
        </button>
      </div>
    </div>

    <p
      v-if="entries.length === 0"
      class="history-empty"
      data-testid="history-empty"
    >
      No versions yet
    </p>
    <ol
      v-else
      class="history-list"
      data-testid="history-list"
    >
      <li
        v-for="entry in entries"
        :key="entry.id"
        class="history-version"
        :class="{ 'is-previewing': entry.id === c.previewVersionId.value }"
        data-testid="history-version"
      >
        <div class="history-meta">
          <span>{{ versionTime(entry) }}</span>
          <span class="history-author">{{ entry.author }}</span>
        </div>
        <div
          class="history-name"
          data-testid="history-name"
        >
          {{ entry.name || "Unnamed version" }}
        </div>
        <div class="history-facts">
          {{ slideCountLabel(entry.slides ?? 0) }} · {{ slideDeltaLabel(entry.slideDelta) }}<template v-if="triggerLabel(entry)">
            · <span data-testid="history-trigger">{{ triggerLabel(entry) }}</span>
          </template>
        </div>
        <div class="history-actions">
          <UiButton
            size="sm"
            data-testid="history-preview"
            @click="c.preview(entry.id)"
          >
            Preview
          </UiButton>
          <UiButton
            size="sm"
            data-testid="history-restore"
            @click="c.restore(entry.id)"
          >
            Restore
          </UiButton>
          <UiButton
            size="sm"
            data-testid="history-download"
            @click="c.downloadVersion(entry.id)"
          >
            {{ downloadLabel(c.format.value) }}
          </UiButton>
          <form
            class="history-name-form"
            @submit.prevent="(event) => onName(entry.id, event)"
          >
            <input
              name="name"
              type="text"
              placeholder="Name this version"
              aria-label="Version name"
              data-testid="history-name-input"
              :value="entry.name || ''"
            >
            <UiButton
              size="sm"
              type="submit"
              data-testid="history-name-submit"
            >
              Name
            </UiButton>
          </form>
          <UiButton
            v-if="entry.name"
            size="sm"
            data-testid="history-remove-name"
            @click="c.removeName(entry.id)"
          >
            Remove name
          </UiButton>
          <UiButton
            size="sm"
            variant="ghost"
            data-testid="history-delete"
            @click="c.deleteVersion(entry.id)"
          >
            Delete
          </UiButton>
        </div>
      </li>
    </ol>
  </section>
</template>

<style scoped>
.history-panel {
  flex: 0 0 auto;
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
  padding: var(--space-4);
  border-top: var(--border-width) solid var(--border-subtle);
  background: var(--surface-default);
  color: var(--text-primary);
  font-size: var(--text-sm);
  max-height: 45%;
  overflow-y: auto;
}

.history-head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--space-3);
}

.history-title {
  margin: 0;
  font-size: var(--text-sm);
  font-weight: var(--weight-bold);
  letter-spacing: 0.04em;
  text-transform: uppercase;
  color: var(--text-secondary);
}

.history-filter {
  display: inline-flex;
  align-items: center;
  gap: var(--space-2);
  color: var(--text-secondary);
}

.history-note {
  margin: 0;
  color: var(--text-secondary);
}

.history-paused {
  margin: 0;
  padding: var(--space-2) var(--space-3);
  border: var(--border-width) solid var(--warning);
  border-radius: var(--radius-sm);
  background: var(--surface-sunken);
  color: var(--warning);
}

.history-preview {
  display: flex;
  flex-direction: column;
  align-items: flex-start;
  gap: var(--space-2);
  padding: var(--space-3);
  border: var(--border-width) solid var(--accent);
  border-radius: var(--radius-sm);
  background: var(--surface-selected);
}

.history-preview-head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--space-3);
  align-self: stretch;
  color: var(--text-accent);
}

/* The preview is a picture of the version: nothing on it takes input. */
.history-preview-stage {
  pointer-events: none;
}

.history-preview-strip {
  display: flex;
  flex-wrap: wrap;
  gap: var(--space-1);
}

.history-preview-slide {
  min-width: 2rem;
  padding: var(--space-1) var(--space-2);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  color: var(--text-primary);
  cursor: pointer;
}

.history-preview-slide[aria-pressed="true"] {
  border-color: var(--accent);
  background: var(--surface-selected);
}

.history-empty {
  margin: 0;
  font-style: italic;
  color: var(--text-secondary);
}

.history-list {
  list-style: none;
  margin: 0;
  padding: 0;
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
}

.history-version {
  display: flex;
  flex-direction: column;
  gap: var(--space-1);
  padding: var(--space-3);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
  background: var(--surface-base);
}

.history-version.is-previewing {
  border-color: var(--accent);
  background: var(--surface-selected);
}

.history-meta {
  display: flex;
  justify-content: space-between;
  gap: var(--space-3);
  font-size: var(--text-xs);
  color: var(--text-secondary);
}

.history-author,
.history-name {
  font-weight: var(--weight-bold);
  color: var(--text-primary);
}

.history-facts {
  color: var(--text-secondary);
}

.history-actions {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: var(--space-2);
  margin-top: var(--space-1);
}

.history-name-form {
  display: inline-flex;
  gap: var(--space-1);
}

.history-name-form input {
  width: 9rem;
  padding: var(--space-1) var(--space-2);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  font-size: var(--text-xs);
}
</style>
