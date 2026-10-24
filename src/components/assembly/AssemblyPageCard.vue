<script setup lang="ts">
// One page of a PDF plan in the tray (SPEC-111): a paper-shaped placeholder
// at the page's REAL aspect (a quarter turn swaps its sides) carrying its
// position, and beside it the page's number in its source, the source badge,
// its size, the rotation it will have and whether it is selected. There is
// no rasterised preview — this product has no PDF renderer — and the card
// says so rather than faking one; the result preview shows the real pages.
// Its buttons turn the page left or right, duplicate it, delete it from the
// plan (never from the stored source) or select every page of its source —
// on the whole selection when this page is part of it. A blank page reads
// at the sheet's size. All logic is the useAssemblyTray controller passed
// in. Runtime prop declarations (espree lints .vue scripts).
import { computed } from "vue";
import UiButton from "../ui/UiButton.vue";
import { cardAspect, rotationLabel } from "../../lib/pdfMerge";
import { pageSizeLabel } from "../../lib/media/thumbnail";

const props = defineProps({
  /** The tray entry (a page or a blank page). */
  entry: { type: Object, required: true },
  /** Its 1-based place in the plan. */
  position: { type: Number, required: true },
  selected: { type: Boolean, default: false },
  /** The useAssemblyTray() controller. */
  tray: { type: Object, required: true },
  /** A blank page's size: the sheet's, in points (0 while unresolved). */
  sheetWidthPt: { type: Number, default: 0 },
  sheetHeightPt: { type: Number, default: 0 },
});

const page = computed(() => ({
  widthPt: props.entry.page ? props.entry.page.widthPt : props.sheetWidthPt,
  heightPt: props.entry.page ? props.entry.page.heightPt : props.sheetHeightPt,
  rotate: props.entry.page ? props.entry.page.rotate : 0,
  rotateBy: props.entry.rotateBy ?? 0,
}));

const aspect = computed(() => cardAspect(page.value));
// The paper fits a 3 × 3.6 rem box, at the page's own proportions.
const paperStyle = computed(() => {
  const width = Math.min(3, 3.6 * aspect.value);
  return { width: `${width.toFixed(3)}rem`, height: `${(width / aspect.value).toFixed(3)}rem` };
});

const sourceLabel = computed(() =>
  props.entry.page
    ? `${props.entry.title} · p. ${props.entry.page.index + 1} of ${props.entry.page.count}`
    : "Blank page",
);
const sizeLabel = computed(() =>
  page.value.widthPt > 0 && page.value.heightPt > 0 ? pageSizeLabel(page.value) : "Sheet size",
);
const rotation = computed(() => rotationLabel(page.value));
// How the buttons name this page to a screen reader.
const name = computed(() => `page ${props.position}`);
</script>

<template>
  <div class="page-card">
    <span
      class="page-paper-box"
      aria-hidden="true"
    >
      <span
        class="page-paper"
        :class="{ 'page-paper--blank': !props.entry.page }"
        data-testid="page-paper"
        :data-aspect="aspect.toFixed(4)"
        :style="paperStyle"
      >
        <span class="page-paper-number">{{ props.position }}</span>
      </span>
    </span>
    <div class="page-text">
      <span
        class="page-number"
        data-testid="page-number"
      >Page {{ props.position }}</span>
      <span
        class="page-source"
        data-testid="page-source"
      >{{ sourceLabel }}</span>
      <span class="page-meta">
        <span data-testid="page-size">{{ sizeLabel }}</span> · <span data-testid="page-rotation">{{ rotation }}</span>
      </span>
      <span class="page-meta page-no-preview">No preview</span>
      <span
        v-if="props.selected"
        class="page-selected-badge"
        data-testid="page-selected"
      >Selected</span>
    </div>
    <div class="page-actions">
      <UiButton
        size="sm"
        variant="ghost"
        data-testid="page-rotate-left"
        :aria-label="`Rotate ${name} left`"
        title="Rotate left"
        @click.stop="props.tray.rotate(props.entry.key, -90)"
      >
        ↺
      </UiButton>
      <UiButton
        size="sm"
        variant="ghost"
        data-testid="page-rotate-right"
        :aria-label="`Rotate ${name} right`"
        title="Rotate right"
        @click.stop="props.tray.rotate(props.entry.key, 90)"
      >
        ↻
      </UiButton>
      <UiButton
        size="sm"
        variant="ghost"
        data-testid="page-duplicate"
        :aria-label="`Duplicate ${name}`"
        :disabled="props.tray.isPagesFull.value"
        @click.stop="props.tray.duplicatePages(props.entry.key)"
      >
        Duplicate
      </UiButton>
      <UiButton
        size="sm"
        variant="ghost"
        data-testid="page-delete"
        :aria-label="`Delete ${name} from the plan`"
        @click.stop="props.tray.removePages(props.entry.key)"
      >
        Delete
      </UiButton>
      <UiButton
        v-if="props.entry.page"
        size="sm"
        variant="ghost"
        data-testid="page-select-source"
        :aria-label="`Select all pages from ${props.entry.title}`"
        :title="`Select all pages from ${props.entry.title}`"
        @click.stop="props.tray.selectSource(props.entry.key)"
      >
        Select file
      </UiButton>
    </div>
  </div>
</template>

<style scoped>
.page-card {
  display: grid;
  grid-template-columns: 3rem minmax(0, 1fr);
  grid-template-areas:
    "paper text"
    "actions actions";
  align-items: center;
  column-gap: var(--space-2);
  row-gap: var(--space-1);
  min-width: 0;
}

.page-paper-box {
  grid-area: paper;
  width: 3rem;
  height: 3.6rem;
  display: flex;
  align-items: center;
  justify-content: center;
}

/* Paper, not a picture: a white sheet at the page's proportions. */
.page-paper {
  display: flex;
  align-items: center;
  justify-content: center;
  box-sizing: border-box;
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  box-shadow: 0 1px 2px color-mix(in srgb, var(--text-primary) 15%, transparent);
}

.page-paper--blank {
  border-style: dashed;
}

.page-paper-number {
  font-weight: var(--weight-medium);
  color: var(--text-secondary);
}

.page-text {
  grid-area: text;
  display: flex;
  flex-direction: column;
  min-width: 0;
}

.page-number {
  font-weight: var(--weight-medium);
}

.page-source {
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.page-meta {
  color: var(--text-secondary);
}

.page-no-preview {
  font-style: italic;
}

.page-selected-badge {
  align-self: flex-start;
  margin-top: var(--space-1);
  padding: 0 var(--space-1);
  border-radius: var(--radius-sm);
  background: var(--surface-selected);
  color: var(--text-accent);
  font-size: var(--text-xs, var(--text-sm));
  font-weight: var(--weight-medium);
}

.page-actions {
  grid-area: actions;
  display: flex;
  flex-wrap: wrap;
  gap: var(--space-1);
}
</style>
