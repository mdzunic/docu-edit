<script setup lang="ts">
// The layout gallery (SPEC-089): a modal over BaseDialog showing the deck's own
// layouts plus the synthesised standard fallbacks as real thumbnails (a preview
// Slide of each layout's slots rendered read-only by SlideStage), with roving
// arrow-key focus, Apply / Reset slide / Close, and a dismissible "N items kept
// in place" report. Its props are exactly { controller, slideNumber }; all typed
// transforms live in the sibling layoutCards.ts (this SFC is espree-linted).
import { computed, nextTick, ref, useTemplateRef, watch } from "vue";
import BaseDialog from "../BaseDialog.vue";
import SlideStage from "./SlideStage.vue";
import { filmstripThumbScale } from "./slideRender";
import { keptInPlaceNames, reportHeadline, toLayoutCards } from "./layoutCards";

const props = defineProps({
  controller: { type: Object, required: true },
  slideNumber: { type: Number, default: 1 },
});
const c = props.controller;

const cards = computed(() => toLayoutCards(c.cards.value));
const thumbScale = computed(() => filmstripThumbScale(c.widthCm.value || 1, 176));
const title = computed(() => (c.mode.value === "new" ? "New slide with layout" : `Layout — slide ${props.slideNumber}`));

const selectedIndex = computed(() => {
  const i = cards.value.findIndex((card) => card.key === c.currentKey.value);
  return i >= 0 ? i : 0;
});
const focusIndex = ref(0);
const listRef = useTemplateRef("list");

function optionEls() {
  const root = listRef.value;
  return root ? Array.from(root.querySelectorAll('[role="option"]')) : [];
}
function focusOption(index = 0) {
  const el = optionEls()[index];
  if (el instanceof HTMLElement) {
    el.focus();
  }
}

watch(
  () => c.open.value,
  (open) => {
    if (open) {
      focusIndex.value = selectedIndex.value;
      nextTick(() => focusOption(focusIndex.value));
    }
  },
  { immediate: true },
);

function focusAt(index = 0) {
  const clamped = Math.min(Math.max(index, 0), cards.value.length - 1);
  focusIndex.value = clamped;
  nextTick(() => focusOption(clamped));
}
function onListKeydown(event = new KeyboardEvent("keydown")) {
  switch (event.key) {
    case "ArrowRight":
    case "ArrowDown":
      event.preventDefault();
      focusAt(focusIndex.value + 1);
      break;
    case "ArrowLeft":
    case "ArrowUp":
      event.preventDefault();
      focusAt(focusIndex.value - 1);
      break;
    case "Home":
      event.preventDefault();
      focusAt(0);
      break;
    case "End":
      event.preventDefault();
      focusAt(cards.value.length - 1);
      break;
    case "Enter":
    case " ":
      event.preventDefault();
      choose(focusIndex.value);
      c.apply();
      break;
    default:
      break;
  }
}
function choose(index = 0) {
  const card = cards.value[index];
  if (card) {
    c.choose(card.key);
  }
}

const reportNames = computed(() => (c.report.value ? keptInPlaceNames(c.report.value, c.currentSlide.value) : []));
const reportText = computed(() => (c.report.value ? reportHeadline(c.report.value) : ""));
</script>

<template>
  <BaseDialog
    title=""
    testid="slides-layout-gallery"
    width="46rem"
    @dismiss="c.close()"
  >
    <template #header-extra>
      <h2 class="gallery-title">
        {{ title }}
      </h2>
      <button
        type="button"
        class="gallery-close"
        data-testid="layout-gallery-close"
        aria-label="Close"
        @click="c.close()"
      >
        ✕
      </button>
    </template>

    <ul
      ref="list"
      class="layout-list"
      role="listbox"
      aria-label="Slide layouts"
      @keydown="onListKeydown"
    >
      <li
        v-for="(card, index) in cards"
        :key="card.key"
        class="layout-list-item"
        role="presentation"
      >
        <button
          type="button"
          class="layout-card"
          :class="{ 'layout-card--selected': index === selectedIndex }"
          role="option"
          :data-testid="`layout-card-${index}`"
          :aria-selected="index === selectedIndex"
          :aria-label="card.name"
          :tabindex="index === focusIndex ? 0 : -1"
          @click="choose(index)"
          @focus="focusIndex = index"
        >
          <span class="layout-card-thumb">
            <SlideStage
              :slide="card.previewSlide"
              :width-cm="c.widthCm.value"
              :height-cm="c.heightCm.value"
              :scale="thumbScale"
              :interactive="false"
            />
          </span>
          <span class="layout-card-name">{{ card.name }}</span>
          <span
            v-if="card.isNew"
            class="layout-card-new-badge"
            data-testid="layout-card-new-badge"
          >will be added to this deck</span>
        </button>
      </li>
    </ul>

    <div
      v-if="c.report.value"
      class="layout-report"
      data-testid="layout-report"
      role="status"
    >
      <span>{{ reportText }}<template v-if="reportNames.length"> — {{ reportNames.join(", ") }}</template></span>
      <button
        type="button"
        class="layout-report-dismiss"
        data-testid="layout-report-dismiss"
        @click="c.dismissReport()"
      >
        Dismiss
      </button>
    </div>

    <template #actions>
      <button
        type="button"
        class="layout-reset"
        data-testid="layout-reset-slide"
        :aria-disabled="!c.canReset.value"
        :disabled="!c.canReset.value"
        @click="c.resetSlide()"
      >
        Reset slide
      </button>
      <button
        type="button"
        class="layout-apply"
        data-testid="layout-apply"
        @click="c.apply()"
      >
        {{ c.mode.value === "new" ? "Add slide" : "Apply" }}
      </button>
    </template>
  </BaseDialog>
</template>

<style scoped>
.gallery-title {
  flex: 1;
  margin: 0;
  font-size: var(--text-lg, 1.125rem);
}
.gallery-close {
  border: none;
  background: transparent;
  cursor: pointer;
  font-size: 1rem;
  line-height: 1;
  padding: var(--space-1, 0.25rem);
}
.layout-list {
  list-style: none;
  margin: 0;
  padding: 0;
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(11rem, 1fr));
  gap: var(--space-3, 0.75rem);
}
.layout-list-item {
  margin: 0;
}
.layout-card {
  display: flex;
  flex-direction: column;
  gap: var(--space-1, 0.25rem);
  width: 100%;
  padding: var(--space-2, 0.5rem);
  border: 2px solid var(--border, #d0d0d0);
  border-radius: var(--radius-md, 6px);
  background: var(--surface, #fff);
  cursor: pointer;
  text-align: left;
}
.layout-card--selected {
  border-color: var(--accent, #2563eb);
}
.layout-card:focus-visible {
  outline: 2px solid var(--accent, #2563eb);
  outline-offset: 2px;
}
.layout-card-thumb {
  display: block;
  border: 1px solid var(--border, #e0e0e0);
  overflow: hidden;
}
.layout-card-name {
  font-weight: 600;
  font-size: var(--text-sm, 0.875rem);
}
.layout-card-new-badge {
  font-size: var(--text-xs, 0.75rem);
  color: var(--text-muted, #666);
}
.layout-report {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--space-2, 0.5rem);
  margin-top: var(--space-3, 0.75rem);
  padding: var(--space-2, 0.5rem);
  background: var(--surface-2, #f4f4f5);
  border-radius: var(--radius-md, 6px);
  font-size: var(--text-sm, 0.875rem);
}
.layout-reset[aria-disabled="true"] {
  opacity: 0.5;
  cursor: not-allowed;
}
</style>
