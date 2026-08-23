<script setup lang="ts">
import { onBeforeUnmount, onMounted, ref, useTemplateRef } from "vue";
import { getDocument } from "../../lib/documentStore";
import { runWhenIdle } from "../../lib/core/idleScheduler";
import {
  getCachedThumbnail,
  renderThumbnail,
  THUMBNAIL_ASPECT_RATIO,
} from "../../lib/libraryThumbnail";
import { formatBadgeLabel } from "../../lib/format";

// A document's first-page thumbnail (SPEC-067). It NEVER blocks the library:
// the card renders immediately with a paper-shaped skeleton, and the real
// preview fills in only once the card scrolls into view (IntersectionObserver
// gating, AC-67/68/71) and the browser is idle (SPEC-047 idle scheduler). A
// cached thumbnail (produced once per edit, keyed by id + lastModified) shows
// on the next mount without re-rendering. Runtime prop declarations — the .vue
// ESLint parser rejects TS-only syntax (see AppShell.vue).
const props = defineProps({
  id: { type: String, required: true },
  lastModified: { type: Number, required: true },
  format: { type: String, required: true },
});

const rootRef = useTemplateRef("root");
const thumbnail = ref(getCachedThumbnail(props.id, props.lastModified) ?? null);

// A disconnectable stub as the initial value gives `observer` an inferred
// { disconnect(): void } type without a TS annotation (barred in .vue scripts)
// — a real IntersectionObserver is assignable to it.
let observer = { disconnect() {} };

// runWhenIdle is not itself cancellable, so a scheduled render can fire up to
// 500ms after the card has scrolled out and unmounted — during whatever route
// the user navigated to next. This guard drops that orphaned work so the
// library never spends idle time rendering a thumbnail for a gone card.
let unmounted = false;

function produce() {
  if (thumbnail.value) {
    return;
  }
  runWhenIdle(() => {
    if (unmounted || thumbnail.value) {
      return;
    }
    const model = getDocument(props.id);
    thumbnail.value = model
      ? renderThumbnail(props.id, props.lastModified, model)
      : { lines: [], failed: true };
  }, 500);
}

onMounted(() => {
  if (thumbnail.value) {
    return;
  }
  // No IntersectionObserver (jsdom, older engines): render on the next idle
  // tick rather than never — correctness over the perf optimisation.
  if (typeof IntersectionObserver !== "function" || !rootRef.value) {
    produce();
    return;
  }
  const io = new IntersectionObserver((entries) => {
    if (entries.some((entry) => entry.isIntersecting)) {
      io.disconnect();
      produce();
    }
  });
  io.observe(rootRef.value);
  observer = io;
});

onBeforeUnmount(() => {
  unmounted = true;
  observer.disconnect();
});
</script>

<template>
  <div
    ref="root"
    class="thumbnail"
    :style="{ aspectRatio: String(THUMBNAIL_ASPECT_RATIO) }"
    data-testid="row-thumbnail"
    aria-hidden="true"
  >
    <div
      v-if="thumbnail && !thumbnail.failed && thumbnail.lines.length > 0"
      class="thumbnail-page"
    >
      <p
        v-for="(line, index) in thumbnail.lines"
        :key="index"
        class="thumbnail-line"
        :class="line.level > 0 ? `is-heading level-${Math.min(line.level, 3)}` : 'is-body'"
      >
        {{ line.text }}
      </p>
    </div>
    <div
      v-else-if="thumbnail"
      class="thumbnail-fallback"
    >
      <span class="thumbnail-badge">{{ formatBadgeLabel(props.format) }}</span>
    </div>
    <div
      v-else
      class="thumbnail-skeleton"
    >
      <span class="skeleton-line" />
      <span class="skeleton-line" />
      <span class="skeleton-line short" />
    </div>
  </div>
</template>

<style scoped>
.thumbnail {
  width: 100%;
  overflow: hidden;
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-subtle);
  border-radius: var(--radius-sm);
}

.thumbnail-page {
  height: 100%;
  padding: var(--space-3);
  display: flex;
  flex-direction: column;
  gap: calc(var(--space-1) / 2);
}

/* Font sizes are a fraction of the smallest type token (calc is treated as an
   opaque token by the design-system guard) — the miniature needs sizes far
   below any standalone token. */
.thumbnail-line {
  margin: 0;
  color: var(--text-primary);
  font-size: calc(var(--text-sm) * 0.5);
  overflow: hidden;
  white-space: nowrap;
  text-overflow: ellipsis;
}

.thumbnail-line.is-body {
  color: var(--text-secondary);
}

.thumbnail-line.is-heading {
  font-weight: var(--weight-bold);
}

.thumbnail-line.level-1 {
  font-size: calc(var(--text-sm) * 0.85);
}

.thumbnail-line.level-2 {
  font-size: calc(var(--text-sm) * 0.72);
}

.thumbnail-line.level-3 {
  font-size: calc(var(--text-sm) * 0.6);
}

.thumbnail-fallback {
  height: 100%;
  display: flex;
  align-items: center;
  justify-content: center;
  background: var(--surface-raised);
}

.thumbnail-badge {
  padding: var(--space-1) var(--space-3);
  border-radius: var(--radius-pill);
  background: var(--surface-selected);
  color: var(--text-accent);
  font-size: var(--text-sm);
  font-weight: var(--weight-bold);
}

.thumbnail-skeleton {
  height: 100%;
  padding: var(--space-4);
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
  justify-content: flex-start;
  background: var(--surface-raised);
}

.skeleton-line {
  height: var(--space-2);
  border-radius: var(--radius-pill);
  background: var(--surface-sunken);
}

.skeleton-line.short {
  width: 60%;
}
</style>
