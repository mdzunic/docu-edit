<script setup lang="ts">
// The PDF's page strip (SPEC-110), beneath the preview: one tile per page,
// the current one highlighted (aria-current="page"), each a miniature of its
// sheet with the placed rectangles where they land. A click shows that page
// in the preview; PageUp/PageDown (handled by the preview around it) step
// through them. All logic is the useAssembly controller's `pdf`. Runtime
// prop declarations (espree lints .vue scripts).
import { computed } from "vue";
import { asPdfPages } from "../../views/assemble/usePdfAssembly";

const props = defineProps({
  /** The useAssembly() controller. */
  assembly: { type: Object, required: true },
});

const a = props.assembly;
const pdf = a.pdf;
// The pages, typed for the template.
const pages = computed(() => asPdfPages(pdf.pages.value));

// The tiles take the sheet's shape.
const aspect = computed(() => {
  const sheet = a.geometry.value;
  return sheet && sheet.heightCm > 0 ? `${sheet.widthCm} / ${sheet.heightCm}` : "1 / 1.414";
});

function rectStyle(item = { xCm: 0, yCm: 0, widthCm: 0, heightCm: 0 }) {
  const sheet = a.geometry.value;
  if (!sheet || sheet.widthCm <= 0 || sheet.heightCm <= 0) {
    return {};
  }
  return {
    left: `${(item.xCm / sheet.widthCm) * 100}%`,
    top: `${(item.yCm / sheet.heightCm) * 100}%`,
    width: `${(item.widthCm / sheet.widthCm) * 100}%`,
    height: `${(item.heightCm / sheet.heightCm) * 100}%`,
  };
}

function pageLabel(page = { itemIndices: [0] }, index = 0) {
  const titles = page.itemIndices.map((itemIndex = 0) => `“${a.tray.entries.value[itemIndex]?.title ?? ""}”`);
  return `Page ${index + 1} of ${pdf.pageCount.value}: ${titles.join(", ")}`;
}
</script>

<template>
  <nav
    class="page-strip"
    aria-label="Pages"
  >
    <ol
      class="page-strip-list"
      data-testid="assembly-pages"
    >
      <li
        v-for="(page, index) in pages"
        :key="index"
      >
        <button
          type="button"
          class="page-strip-page"
          :class="{ 'page-strip-page--current': index === pdf.current.value }"
          data-testid="assembly-page"
          :aria-current="index === pdf.current.value ? 'page' : undefined"
          :aria-label="pageLabel(page, index)"
          @click="pdf.goTo(index)"
        >
          <span
            class="page-strip-sheet"
            :style="{ aspectRatio: aspect }"
            aria-hidden="true"
          >
            <span
              v-for="item in page.placed"
              :key="item.index"
              class="page-strip-rect"
              :style="rectStyle(item)"
            />
          </span>
          <span
            class="page-strip-number"
            aria-hidden="true"
          >{{ index + 1 }}</span>
        </button>
      </li>
    </ol>
  </nav>
</template>

<style scoped>
.page-strip {
  min-width: 0;
}

.page-strip-list {
  list-style: none;
  margin: 0;
  padding: var(--space-1);
  display: flex;
  gap: var(--space-2);
  overflow-x: auto;
}

.page-strip-page {
  font: inherit;
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: var(--space-1);
  padding: var(--space-1);
  border: var(--border-width) solid transparent;
  border-radius: var(--radius-sm);
  background: transparent;
  color: var(--text-secondary);
  cursor: pointer;
}

.page-strip-page--current {
  border-color: var(--focus-ring);
  background: var(--surface-selected);
  color: var(--text-accent);
}

.page-strip-page:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: 0;
}

.page-strip-sheet {
  position: relative;
  display: block;
  width: 2.5rem;
  border: var(--border-width) solid var(--border-default);
  background: var(--surface-base);
}

.page-strip-rect {
  position: absolute;
  background: var(--border-default);
}

.page-strip-number {
  font-size: var(--text-xs);
}
</style>
