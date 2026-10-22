<script setup lang="ts">
import { onMounted, ref } from "vue";
import { getDocument, getMedia, getPresentation, getSpreadsheet } from "../../lib/documentStore";
import { runWhenIdle } from "../../lib/core/idleScheduler";
import { getDocumentLibraryStats } from "../../lib/libraryStats";
import { formatBadgeLabel, formatFileSize, formatRelativeTime } from "../../lib/format";
import { editorPathFor } from "../../lib/documentKinds";
import { currentImageSize } from "../../lib/media/editModel";
import LibraryThumbnail from "./LibraryThumbnail.vue";
import HighlightedTitle from "./HighlightedTitle.vue";
import RowMenu from "./RowMenu.vue";

// A grid card (SPEC-067): a first-page thumbnail (AC-3), the title (AC-4),
// a format badge (AC-5), relative modified time (AC-6) and page/word counts
// (AC-7). Word/page counts are computed off the critical path (idle scheduler)
// so a 200-card grid still paints its first screen immediately (AC-69).
// Runtime prop/emit declarations — the .vue ESLint parser rejects TS-only
// syntax (see AppShell.vue). `entry` is a LibraryEntry; each menu item is
// { id, label, testid, danger? }.
const props = defineProps({
  entry: { type: Object, required: true },
  search: { type: String, default: "" },
  now: { type: Number, required: true },
  selected: { type: Boolean, default: false },
  focused: { type: Boolean, default: false },
  menuOpen: { type: Boolean, default: false },
  highlight: { type: Boolean, default: false },
  menuItems: { type: Array, default: () => [] },
});

const emit = defineEmits([
  "open",
  "preview",
  "select",
  "toggle-check",
  "toggle-star",
  "menu-toggle",
  "menu-close",
  "menu-select",
  "focus",
]);

// Seeded (not null) so its inferred type is the stats shape without a TS
// annotation; `statsReady` gates the display until the idle pass fills it in.
const statsReady = ref(false);
const stats = ref({ words: 0, pages: 0 });
// A presentation entry's count line (its kind is on the entry itself).
const slideTotal = ref(0);
// A media entry's count line (SPEC-107): dimensions or page count, plus the
// verbatim source's byte size.
const mediaLabel = ref("");
// A workbook entry's count line (SPEC-116): "N sheets", hidden ones counted.
const sheetTotal = ref(0);

onMounted(() => {
  runWhenIdle(() => {
    if (props.entry.kind === "presentation") {
      const presentation = getPresentation(props.entry.id);
      if (presentation) {
        slideTotal.value = presentation.slides.length;
        statsReady.value = true;
      }
      return;
    }
    if (props.entry.kind === "spreadsheet") {
      const workbook = getSpreadsheet(props.entry.id);
      if (workbook) {
        sheetTotal.value = workbook.sheets.length;
        statsReady.value = true;
      }
      return;
    }
    if (props.entry.kind === "media") {
      const media = getMedia(props.entry.id);
      if (media) {
        const size = formatFileSize(media.byteLength);
        if (media.pdf) {
          const pages = media.pdf.pageCount === 1 ? "1 page" : `${media.pdf.pageCount} pages`;
          mediaLabel.value = `${pages} · ${size}`;
        } else if (media.image) {
          // The current image's size — an edited image's render (SPEC-112).
          const image = currentImageSize(media) ?? media.image;
          mediaLabel.value = `${image.widthPx} × ${image.heightPx} px · ${size}`;
        } else {
          mediaLabel.value = size;
        }
        statsReady.value = true;
      }
      return;
    }
    const model = getDocument(props.entry.id);
    if (model) {
      stats.value = getDocumentLibraryStats(props.entry.id, props.entry.lastModified, model);
      statsReady.value = true;
    }
  }, 800);
});

// A plain click previews the card in the quick-look pane rather than opening
// it (the handoff's select-then-open grid); opening is the double-click, the
// keyboard Enter, or the quick look's own Open button. Modifier clicks keep
// their multi-selection semantics.
function onClick(event = new MouseEvent("click")) {
  if (event.shiftKey || event.ctrlKey || event.metaKey) {
    event.preventDefault();
    emit("select", {
      id: props.entry.id,
      shift: event.shiftKey,
      ctrlOrMeta: event.ctrlKey || event.metaKey,
    });
    return;
  }
  emit("preview", props.entry.id);
}
</script>

<template>
  <div
    class="library-card"
    role="row"
    data-testid="library-row"
    :data-doc-id="props.entry.id"
    :class="{ 'is-selected': props.selected, 'row-highlight': props.highlight }"
    :tabindex="props.focused ? 0 : -1"
    :aria-selected="props.selected"
    @click="onClick"
    @dblclick="emit('open', props.entry.id)"
    @focus="emit('focus', props.entry.id)"
  >
    <div class="card-thumb">
      <LibraryThumbnail
        :id="props.entry.id"
        :last-modified="props.entry.lastModified"
        :format="props.entry.format"
      />
      <input
        type="checkbox"
        class="card-check"
        data-testid="row-select"
        :checked="props.selected"
        :aria-label="`Select ${props.entry.title}`"
        @click.stop
        @change="emit('toggle-check', props.entry.id)"
      >
      <button
        type="button"
        class="card-star"
        data-testid="row-star"
        :class="{ 'is-starred': props.entry.starred }"
        :aria-pressed="props.entry.starred"
        :aria-label="props.entry.starred ? `Unstar ${props.entry.title}` : `Star ${props.entry.title}`"
        @click.stop="emit('toggle-star', props.entry.id)"
      >
        {{ props.entry.starred ? "★" : "☆" }}
      </button>
    </div>
    <div class="card-body">
      <div
        class="card-title"
        data-testid="row-title"
      >
        <!-- A real link for the address (hover URL, "Open in new tab", "Copy
             link", middle-click), while the card keeps the click: a plain one
             previews, a double-click opens and a modified one multi-selects,
             so this anchor's own default is prevented. -->
        <a
          class="title-link"
          :href="editorPathFor(props.entry.kind, props.entry.id)"
          tabindex="-1"
          @click.prevent
        >
          <HighlightedTitle
            :title="props.entry.title"
            :search="props.search"
          />
        </a>
      </div>
      <div class="card-meta">
        <span
          class="badge"
          data-testid="row-format"
        >{{ formatBadgeLabel(props.entry.format, props.entry.packageVariant) }}</span>
        <span
          class="card-modified"
          data-testid="row-modified"
        >{{ formatRelativeTime(props.entry.lastModified, props.now) }}</span>
      </div>
      <div
        class="card-counts"
        data-testid="row-counts"
      >
        <template v-if="statsReady && props.entry.kind === 'presentation'">
          {{ slideTotal }} {{ slideTotal === 1 ? "slide" : "slides" }}
        </template>
        <template v-else-if="statsReady && props.entry.kind === 'media'">
          {{ mediaLabel }}
        </template>
        <template v-else-if="statsReady && props.entry.kind === 'spreadsheet'">
          {{ sheetTotal }} {{ sheetTotal === 1 ? "sheet" : "sheets" }}
        </template>
        <template v-else-if="statsReady">
          {{ stats.pages }} {{ stats.pages === 1 ? "page" : "pages" }} ·
          {{ stats.words }} {{ stats.words === 1 ? "word" : "words" }}
        </template>
        <template v-else>
          …
        </template>
      </div>
    </div>
    <div
      class="card-actions"
      @click.stop
    >
      <RowMenu
        :open="props.menuOpen"
        :label="`Actions for ${props.entry.title}`"
        :items="props.menuItems"
        @toggle="emit('menu-toggle', props.entry.id)"
        @close="emit('menu-close')"
        @select="(action) => emit('menu-select', { id: props.entry.id, action })"
      />
    </div>
  </div>
</template>

<style scoped>
.library-card {
  position: relative;
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
  padding: var(--space-3);
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
  box-shadow: var(--shadow-1);
  cursor: pointer;
  transition: box-shadow var(--motion-fast) var(--easing);
}

.library-card:hover {
  box-shadow: var(--shadow-2);
  background: var(--surface-raised);
}

.library-card:focus-visible {
  outline: var(--focus-ring-width, 2px) solid var(--border-focus, var(--text-accent));
  outline-offset: 2px;
}

.library-card.is-selected {
  border-color: var(--text-accent);
  background: var(--surface-selected);
}

/* The handoff's thumbnail well: a fixed 4:3 tinted ground bleeding to the
   card's edges, with the miniature centred inside it — the card no longer
   grows to the page preview's full portrait height. */
.card-thumb {
  position: relative;
  aspect-ratio: 4 / 3;
  display: grid;
  place-items: center;
  overflow: hidden;
  margin: calc(-1 * var(--space-3)) calc(-1 * var(--space-3)) 0;
  padding: var(--space-3);
  background: var(--state-hover);
  border-bottom: var(--border-width) solid var(--border-subtle);
}

.card-thumb > :deep(.thumbnail) {
  width: 58%;
}

.card-check {
  position: absolute;
  top: var(--space-2);
  left: var(--space-2);
}

.card-star {
  position: absolute;
  top: var(--space-2);
  right: var(--space-2);
  font-size: var(--text-lg);
  line-height: 1;
  border: none;
  background: transparent;
  cursor: pointer;
  color: var(--text-secondary);
}

.card-star.is-starred {
  color: var(--text-accent);
}

.card-body {
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
}

.card-title {
  font-weight: var(--weight-medium);
  color: var(--text-primary);
  overflow-wrap: anywhere;
}

.card-meta {
  display: flex;
  align-items: center;
  gap: var(--space-3);
}

.card-modified {
  color: var(--text-secondary);
  font-size: var(--text-sm);
}

.card-counts {
  color: var(--text-secondary);
  font-size: var(--text-sm);
}

.badge {
  display: inline-block;
  padding: var(--space-1) var(--space-3);
  border-radius: var(--radius-pill);
  background: var(--surface-selected);
  color: var(--text-accent);
  font-size: var(--text-sm);
  font-weight: var(--weight-bold);
}

.card-actions {
  position: absolute;
  bottom: var(--space-3);
  right: var(--space-3);
}

.row-highlight {
  animation: row-flash var(--motion-base) var(--easing);
}

@keyframes row-flash {
  from {
    background: var(--surface-selected);
  }
  to {
    background: transparent;
  }
}

@media (prefers-reduced-motion: reduce) {
  .row-highlight {
    animation: none;
  }
}

/* The card owns the click and the focus; the anchor carries the address. */
.title-link {
  color: inherit;
  text-decoration: none;
}
</style>
