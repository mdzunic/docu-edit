<script setup lang="ts">
// The assembly tray (SPEC-109): the chosen images, in layout order, each row
// showing its thumbnail, title and pixel size. A row reorders by drag, or by
// Alt+↑ / Alt+↓ while it has focus (focus follows the moved row, and the
// move is spoken through the role="status" live region); its buttons remove
// or duplicate it; a click (or Enter / Space) selects it for the inspector. "Add from library…"
// asks the view to open the picker. For a PDF (SPEC-110) each row says how
// its image will be embedded — "original JPEG", or why it is re-encoded.
// SPEC-111: a PDF in the tray is its PAGES, each an AssemblyPageCard, and the
// tray selects as a set — Shift+click a range, Cmd/Ctrl+click one more,
// Cmd/Ctrl+A everything — for the cards' rotate, duplicate and delete to act
// on; for a PDF it also offers a blank page. All logic is the
// useAssemblyTray controller passed in. Runtime prop declarations (espree
// lints .vue scripts).
import { computed, nextTick, ref, useTemplateRef } from "vue";
import UiButton from "../ui/UiButton.vue";
import LibraryThumbnail from "../library/LibraryThumbnail.vue";
import AssemblyPageCard from "./AssemblyPageCard.vue";
import { MAX_ASSEMBLY_ITEMS } from "../../lib/core/assemblyLayout";
import { pageCountLabel } from "../../lib/assembly";
import { asTrayEntries, isPageEntry, pageCardLabel } from "../../views/assembly/useAssemblyTray";

const props = defineProps({
  /** The useAssemblyTray() controller. */
  tray: { type: Object, required: true },
  /** Per row, how a PDF embeds its image ("" or absent: say nothing). */
  encodings: { type: Array, default: () => [] },
  /** The output is a PDF: a blank page can be added. */
  pdf: { type: Boolean, default: false },
  /** The tray is a PDF plan (it holds pages). */
  merging: { type: Boolean, default: false },
  /** The sheet's size in points, for a blank page's card. */
  sheetWidthPt: { type: Number, default: 0 },
  sheetHeightPt: { type: Number, default: 0 },
});

const emit = defineEmits(["add"]);

const t = props.tray;
const list = useTemplateRef("list");
// The rows, typed for the template.
const rows = computed(() => asTrayEntries(t.entries.value));

function encodingAt(index = 0) {
  const label = props.encodings[index];
  return typeof label === "string" ? label : "";
}

function sizeLabel(entry = { widthPx: 0, heightPx: 0 }) {
  return entry.widthPx > 0 && entry.heightPx > 0 ? `${entry.widthPx} × ${entry.heightPx} px` : "Size unknown";
}

async function focusRow(key = "") {
  await nextTick();
  const row = list.value?.querySelector(`[data-tray-key="${key}"]`);
  if (row instanceof HTMLElement) {
    row.focus();
  }
}

function onRowClick(event = new MouseEvent("click"), key = "") {
  t.pick(key, { range: event.shiftKey, toggle: event.metaKey || event.ctrlKey });
}

// Cmd/Ctrl+A anywhere in the list selects every entry (never a text field's
// own select-all: none lives in the list).
function onListKeydown(event = new KeyboardEvent("keydown")) {
  if ((event.metaKey || event.ctrlKey) && !event.altKey && event.key.toLowerCase() === "a") {
    event.preventDefault();
    t.selectAll();
  }
}

function onRowKeydown(event = new KeyboardEvent("keydown"), key = "") {
  if (event.target !== event.currentTarget) {
    return;
  }
  if (event.key === "Enter" || event.key === " ") {
    // The keyboard's click: select the row for the inspector.
    event.preventDefault();
    t.pick(key, { range: event.shiftKey, toggle: event.metaKey || event.ctrlKey });
    return;
  }
  if (!event.altKey || (event.key !== "ArrowUp" && event.key !== "ArrowDown")) {
    return;
  }
  event.preventDefault();
  t.move(key, event.key === "ArrowUp" ? -1 : 1);
  void focusRow(key);
}

// Drag and drop: the dragged row's key is held here (the dataTransfer copy
// is for the browser's sake), and a drop lands the row at the target's slot.
const dragKey = ref("");
const overIndex = ref(-1);

function onDragStart(event = new Event("dragstart"), key = "") {
  dragKey.value = key;
  const transfer = Reflect.get(event, "dataTransfer");
  if (transfer && typeof transfer.setData === "function") {
    transfer.effectAllowed = "move";
    transfer.setData("text/plain", key);
  }
}

function onDragOver(event = new Event("dragover"), index = 0) {
  if (dragKey.value) {
    event.preventDefault();
    overIndex.value = index;
  }
}

function onDrop(event = new Event("drop"), index = 0) {
  const key = dragKey.value;
  if (!key) {
    return;
  }
  event.preventDefault();
  t.moveTo(key, index);
  dragKey.value = "";
  overIndex.value = -1;
  void focusRow(key);
}

function onDragEnd() {
  dragKey.value = "";
  overIndex.value = -1;
}
</script>

<template>
  <section
    class="assembly-tray"
    aria-labelledby="assembly-tray-heading"
  >
    <div class="tray-head">
      <h2
        id="assembly-tray-heading"
        class="tray-heading"
      >
        {{ props.merging ? "Pages" : "Images" }}
      </h2>
      <span
        class="tray-count"
        data-testid="assembly-count"
      >{{ props.merging ? pageCountLabel(t.entries.value.length) : `${t.entries.value.length} of ${MAX_ASSEMBLY_ITEMS}` }}</span>
    </div>
    <div class="tray-buttons">
      <UiButton
        data-testid="assembly-add"
        :disabled="t.isFull.value"
        @click="emit('add')"
      >
        Add from library…
      </UiButton>
      <UiButton
        v-if="props.pdf"
        data-testid="assembly-blank-page"
        :disabled="t.isPagesFull.value"
        @click="t.addBlank()"
      >
        Add blank page
      </UiButton>
    </div>
    <p
      v-if="t.refusal.value"
      class="tray-refusal"
      data-testid="assembly-tray-refusal"
      role="alert"
    >
      {{ t.refusal.value }}
    </p>
    <ol
      ref="list"
      class="tray-list"
      data-testid="assembly-tray"
      :aria-label="props.merging ? 'Pages, in output order' : 'Images, in layout order'"
      aria-describedby="assembly-tray-hint"
      @keydown="onListKeydown"
    >
      <li
        v-for="(entry, index) in rows"
        :key="entry.key"
        class="tray-row"
        :class="{
          'tray-row--selected': entry.key === t.selectedKey.value || t.isSelected(entry.key),
          'tray-row--over': index === overIndex,
          'tray-row--page': isPageEntry(entry),
          'page-selected': isPageEntry(entry) && t.isSelected(entry.key),
        }"
        :data-testid="isPageEntry(entry) ? 'assembly-page-card' : 'assembly-item'"
        :data-tray-key="entry.key"
        :data-selected="t.isSelected(entry.key) ? 'true' : undefined"
        :aria-current="entry.key === t.selectedKey.value ? 'true' : undefined"
        :aria-label="isPageEntry(entry) ? pageCardLabel(entry, index + 1, t.isSelected(entry.key)) : undefined"
        tabindex="0"
        draggable="true"
        @click="onRowClick($event, entry.key)"
        @keydown="onRowKeydown($event, entry.key)"
        @dragstart="onDragStart($event, entry.key)"
        @dragover="onDragOver($event, index)"
        @drop="onDrop($event, index)"
        @dragend="onDragEnd"
      >
        <AssemblyPageCard
          v-if="isPageEntry(entry)"
          :entry="entry"
          :position="index + 1"
          :selected="t.isSelected(entry.key)"
          :tray="t"
          :sheet-width-pt="props.sheetWidthPt"
          :sheet-height-pt="props.sheetHeightPt"
        />
        <template v-else>
          <span
            class="tray-position"
            aria-hidden="true"
          >{{ index + 1 }}</span>
          <div class="tray-thumb">
            <LibraryThumbnail
              :id="entry.id"
              :last-modified="entry.lastModified"
              :format="entry.format"
            />
          </div>
          <div class="tray-text">
            <span
              class="tray-title"
              data-testid="item-title"
            >{{ entry.title }}</span>
            <span
              class="tray-size"
              data-testid="item-size"
            >{{ sizeLabel(entry) }}</span>
            <span
              v-if="encodingAt(index)"
              class="tray-encoding"
              data-testid="item-encoding"
            >{{ encodingAt(index) }}</span>
          </div>
          <div class="tray-actions">
            <UiButton
              size="sm"
              variant="ghost"
              data-testid="item-duplicate"
              :aria-label="`Duplicate ${entry.title}`"
              :disabled="t.isFull.value"
              @click.stop="t.duplicate(entry.key)"
            >
              Duplicate
            </UiButton>
            <UiButton
              size="sm"
              variant="ghost"
              data-testid="item-remove"
              :aria-label="`Remove ${entry.title}`"
              @click.stop="t.remove(entry.key)"
            >
              Remove
            </UiButton>
          </div>
        </template>
      </li>
    </ol>
    <p
      v-if="t.entries.value.length === 0"
      class="tray-empty"
      data-testid="assembly-tray-empty"
    >
      No images yet. Add some from the library.
    </p>
    <p
      id="assembly-tray-hint"
      class="tray-hint"
    >
      Drag a row, or press Alt+↑ or Alt+↓ on it, to change the order.
      <template v-if="props.merging">
        Shift-click selects a range, Ctrl/Cmd-click one more, Ctrl/Cmd+A everything.
      </template>
    </p>
    <p
      class="visually-hidden"
      data-testid="assembly-tray-status"
      role="status"
    >
      {{ t.announcement.value }}
    </p>
  </section>
</template>

<style scoped>
.assembly-tray {
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
  min-height: 0;
  font-size: var(--text-sm);
  color: var(--text-primary);
}

.tray-head {
  display: flex;
  align-items: baseline;
  justify-content: space-between;
}

.tray-heading {
  margin: 0;
  font-size: var(--text-md);
}

.tray-buttons {
  display: flex;
  flex-wrap: wrap;
  gap: var(--space-2);
}

.tray-count,
.tray-size,
.tray-encoding,
.tray-hint,
.tray-empty {
  color: var(--text-secondary);
}

.tray-hint,
.tray-empty,
.tray-refusal {
  margin: 0;
}

.tray-refusal {
  color: var(--text-danger);
}

.tray-list {
  list-style: none;
  margin: 0;
  padding: 0;
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
  overflow-y: auto;
  min-height: 0;
}

.tray-row {
  display: grid;
  grid-template-columns: 1.5rem 3rem minmax(0, 1fr);
  grid-template-areas:
    "pos thumb text"
    "pos thumb actions";
  align-items: center;
  column-gap: var(--space-2);
  padding: var(--space-2);
  border: var(--border-width) solid var(--border-subtle);
  border-radius: var(--radius-md);
  background: var(--surface-base);
  cursor: grab;
}

.tray-row--selected {
  border-color: var(--border-default);
  background: var(--surface-selected);
}

/* A page card lays itself out. */
.tray-row--page {
  display: block;
}

.tray-row.page-selected {
  border-color: var(--focus-ring);
}

.tray-row--over {
  border-color: var(--focus-ring);
}

.tray-row:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: 0;
}

.tray-position {
  grid-area: pos;
  color: var(--text-secondary);
  text-align: center;
}

.tray-thumb {
  grid-area: thumb;
  width: 3rem;
  max-height: 3rem;
  overflow: hidden;
  border-radius: var(--radius-sm);
}

.tray-text {
  grid-area: text;
  display: flex;
  flex-direction: column;
  min-width: 0;
}

.tray-title {
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  font-weight: var(--weight-medium);
}

.tray-actions {
  grid-area: actions;
  display: flex;
  gap: var(--space-1);
}
</style>
