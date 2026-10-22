<script setup lang="ts">
// "Print and export" for a deck (SPEC-104): one BaseDialog-shelled dialog,
// opened on its PDF tab by the Download menu's "PDF (.pdf)…" and Cmd/Ctrl+P,
// or on its Images tab by "Images (.png)…". Every control is native and
// keyboard-operable: the slide range (all / current / custom numbers),
// "Include hidden slides", and per tab the layout, paper, orientation, note
// lines and "Frame slides" (PDF and print) or the PNG scale (Images). A dumb
// component in the TransitionsPanel mould — the useDeckExport controller owns
// the choices and the runs; this only renders them and forwards input. The
// off-screen DeckPrintSheet rides along. Runtime prop declarations per the
// espree constraint.
import { nextTick, useTemplateRef } from "vue";
import BaseDialog from "../BaseDialog.vue";
import UiButton from "../ui/UiButton.vue";
import DeckPrintSheet from "./DeckPrintSheet.vue";
import { eventValue } from "../../views/presentation/deckSetupFields";
import {
  DECK_PRINT_TABS,
  LAYOUT_OPTIONS,
  PAPER_OPTIONS,
  SCALE_OPTIONS,
  eventChecked,
} from "../../views/presentation/deckPrintFields";

const props = defineProps({
  /** The useDeckExport() controller object (refs + command functions). */
  controller: { type: Object, required: true },
});

const c = props.controller;
const tablist = useTemplateRef("tablist");

function setTab(id = "pdf") {
  c.tab.value = id;
  void nextTick(() => {
    const el = tablist.value?.querySelector(`#deck-print-tab-${id}`);
    if (el instanceof HTMLElement) el.focus();
  });
}

function onTabKeydown(event = new KeyboardEvent("keydown")) {
  if (event.key === "ArrowRight" || event.key === "ArrowLeft" || event.key === "Home" || event.key === "End") {
    event.preventDefault();
    const next = event.key === "Home" ? "pdf" : event.key === "End" ? "images" : c.tab.value === "pdf" ? "images" : "pdf";
    setTab(next);
  }
}

function setRange(value = "all") {
  c.range.value = value;
}
function onRangeText(event = new Event("input")) {
  c.rangeText.value = eventValue(event);
  c.range.value = "custom";
}
function onLayout(event = new Event("change")) {
  c.layout.value = eventValue(event);
}
function onPaper(event = new Event("change")) {
  c.paper.value = eventValue(event);
}
function setOrientation(value = "portrait") {
  c.orientation.value = value;
}
function onNoteLines(event = new Event("change")) {
  c.noteLines.value = eventChecked(event);
}
function onFrame(event = new Event("change")) {
  c.frameSlides.value = eventChecked(event);
}
function onHidden(event = new Event("change")) {
  c.includeHidden.value = eventChecked(event);
}
function onScale(event = new Event("change")) {
  c.scale.value = Number(eventValue(event));
}
</script>

<template>
  <BaseDialog
    title="Print and export"
    testid="deck-print-dialog"
    width="32rem"
    @dismiss="c.close()"
  >
    <div
      ref="tablist"
      class="tablist"
      role="tablist"
      aria-label="Output"
    >
      <button
        v-for="tab in DECK_PRINT_TABS"
        :id="`deck-print-tab-${tab.id}`"
        :key="tab.id"
        type="button"
        role="tab"
        class="tab"
        :class="{ active: c.tab.value === tab.id }"
        :aria-selected="c.tab.value === tab.id ? 'true' : 'false'"
        :aria-controls="`deck-print-panel-${tab.id}`"
        :tabindex="c.tab.value === tab.id ? 0 : -1"
        :data-testid="tab.testid"
        @click="setTab(tab.id)"
        @keydown="onTabKeydown"
      >
        {{ tab.label }}
      </button>
    </div>

    <fieldset
      class="group"
      :disabled="c.busy.value"
    >
      <legend>Slides</legend>
      <div class="choices">
        <label class="radio">
          <input
            type="radio"
            name="deck-print-range"
            value="all"
            data-testid="deck-print-range-all"
            :checked="c.range.value === 'all'"
            @change="setRange('all')"
          >
          All
        </label>
        <label class="radio">
          <input
            type="radio"
            name="deck-print-range"
            value="current"
            data-testid="deck-print-range-current"
            :checked="c.range.value === 'current'"
            @change="setRange('current')"
          >
          Current slide
        </label>
        <label class="radio">
          <input
            type="radio"
            name="deck-print-range"
            value="custom"
            data-testid="deck-print-range-custom"
            :checked="c.range.value === 'custom'"
            @change="setRange('custom')"
          >
          Slides
        </label>
        <input
          type="text"
          class="range-input"
          aria-label="Slide numbers"
          placeholder="e.g. 1-3, 5"
          data-testid="deck-print-range-input"
          :value="c.rangeText.value"
          @input="onRangeText"
        >
      </div>
      <p
        v-if="c.rangeError.value"
        class="error"
        role="alert"
        data-testid="deck-print-range-error"
      >
        {{ c.rangeError.value }}
      </p>
      <label class="check">
        <input
          type="checkbox"
          data-testid="deck-print-hidden"
          :checked="c.includeHidden.value"
          :disabled="!c.hasHidden.value"
          @change="onHidden"
        >
        Include hidden slides
      </label>
    </fieldset>

    <fieldset
      v-show="c.tab.value === 'pdf'"
      id="deck-print-panel-pdf"
      class="group"
      role="tabpanel"
      aria-labelledby="deck-print-tab-pdf"
      :disabled="c.busy.value"
    >
      <legend>Layout</legend>
      <label class="field">
        <span>Print</span>
        <select
          data-testid="deck-print-layout"
          :value="c.layout.value"
          @change="onLayout"
        >
          <option
            v-for="option in LAYOUT_OPTIONS"
            :key="option.value"
            :value="option.value"
          >
            {{ option.label }}
          </option>
        </select>
      </label>
      <label class="field">
        <span>Paper</span>
        <select
          data-testid="deck-print-paper"
          :value="c.paper.value"
          :disabled="c.layout.value === 'slides'"
          @change="onPaper"
        >
          <option
            v-for="option in PAPER_OPTIONS"
            :key="option.value"
            :value="option.value"
          >
            {{ option.label }}
          </option>
        </select>
      </label>
      <div
        class="choices"
        role="radiogroup"
        aria-label="Orientation"
      >
        <label class="radio">
          <input
            type="radio"
            name="deck-print-orientation"
            value="portrait"
            data-testid="deck-print-portrait"
            :checked="c.orientation.value === 'portrait'"
            :disabled="c.layout.value === 'slides'"
            @change="setOrientation('portrait')"
          >
          Portrait
        </label>
        <label class="radio">
          <input
            type="radio"
            name="deck-print-orientation"
            value="landscape"
            data-testid="deck-print-landscape"
            :checked="c.orientation.value === 'landscape'"
            :disabled="c.layout.value === 'slides'"
            @change="setOrientation('landscape')"
          >
          Landscape
        </label>
      </div>
      <label class="check">
        <input
          type="checkbox"
          data-testid="deck-print-note-lines"
          :checked="c.noteLinesChecked.value"
          :disabled="!c.noteLinesEnabled.value"
          @change="onNoteLines"
        >
        Note lines
      </label>
      <label class="check">
        <input
          type="checkbox"
          data-testid="deck-print-frame"
          :checked="c.frameSlides.value"
          @change="onFrame"
        >
        Frame slides
      </label>
      <p
        class="readout"
        data-testid="deck-print-page-count"
      >
        {{ c.pageCount.value === 1 ? "1 page" : `${c.pageCount.value} pages` }}
      </p>
    </fieldset>

    <fieldset
      v-show="c.tab.value === 'images'"
      id="deck-print-panel-images"
      class="group"
      role="tabpanel"
      aria-labelledby="deck-print-tab-images"
      :disabled="c.busy.value"
    >
      <legend>Images</legend>
      <label class="field">
        <span>Scale</span>
        <select
          data-testid="deck-print-scale"
          :value="String(c.scale.value)"
          @change="onScale"
        >
          <option
            v-for="option in SCALE_OPTIONS"
            :key="option.value"
            :value="option.value"
          >
            {{ option.label }}
          </option>
        </select>
      </label>
      <p
        class="readout"
        data-testid="deck-print-estimate"
      >
        {{ c.estimate.value }}
      </p>
    </fieldset>

    <p
      v-if="c.rotationWarning.value"
      class="note"
      data-testid="deck-print-rotation-note"
    >
      Text in rotated frames prints unrotated; shapes keep their rotation.
    </p>
    <p
      v-if="c.notice.value"
      class="error"
      role="alert"
      data-testid="deck-print-notice"
    >
      {{ c.notice.value }}
    </p>
    <div
      v-if="c.busy.value"
      class="progress"
    >
      <span
        role="status"
        data-testid="deck-print-progress"
      >Page {{ c.progress.value.done }} of {{ c.progress.value.total }}</span>
      <UiButton
        variant="secondary"
        data-testid="deck-print-abort"
        @click="c.abort()"
      >
        Stop
      </UiButton>
    </div>

    <template #actions>
      <UiButton
        variant="secondary"
        data-testid="deck-print-cancel"
        @click="c.close()"
      >
        Close
      </UiButton>
      <template v-if="c.tab.value === 'pdf'">
        <UiButton
          variant="secondary"
          data-testid="deck-print-print"
          :disabled="!c.canRun.value"
          @click="c.print()"
        >
          Print…
        </UiButton>
        <UiButton
          variant="primary"
          data-testid="deck-print-pdf"
          :disabled="!c.canRun.value"
          @click="c.exportPdf()"
        >
          Export PDF
        </UiButton>
      </template>
      <UiButton
        v-else
        variant="primary"
        data-testid="deck-print-images"
        :disabled="!c.canRun.value"
        @click="c.exportImages()"
      >
        Export images
      </UiButton>
    </template>
    <DeckPrintSheet :controller="c" />
  </BaseDialog>
</template>

<style scoped>
.tablist {
  display: flex;
  gap: var(--space-2);
  border-bottom: var(--border-width) solid var(--border-default);
  margin-bottom: var(--space-4);
}

.tab {
  font: inherit;
  padding: var(--space-3) var(--space-4);
  border: none;
  border-bottom: calc(2 * var(--border-width)) solid transparent;
  background: transparent;
  cursor: pointer;
  color: var(--text-secondary);
}

.tab.active {
  color: var(--surface-brand);
  border-bottom-color: var(--surface-brand);
  font-weight: var(--weight-bold);
}

.group {
  border: var(--border-width) solid var(--border-subtle);
  border-radius: var(--radius-sm);
  padding: var(--space-2) var(--space-3) var(--space-3);
  margin: 0 0 var(--space-4);
}

.group legend {
  padding: 0 var(--space-1);
  color: var(--text-secondary);
  font-size: var(--text-sm);
}

.choices {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: var(--space-2) var(--space-4);
  margin-bottom: var(--space-2);
}

.radio,
.check {
  display: inline-flex;
  align-items: center;
  gap: var(--space-1);
}

.check {
  display: flex;
  margin-top: var(--space-2);
}

.range-input {
  font: inherit;
  flex: 1 1 8rem;
  min-width: 6rem;
}

.field {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--space-3);
  margin-bottom: var(--space-3);
}

.field select {
  font: inherit;
  min-width: 12rem;
}

.readout,
.note {
  margin: var(--space-2) 0 0;
  color: var(--text-secondary);
  font-size: var(--text-sm);
}

.error {
  margin: var(--space-2) 0;
  color: var(--text-danger);
  font-size: var(--text-sm);
}

.progress {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--space-3);
  margin-top: var(--space-3);
}
</style>
