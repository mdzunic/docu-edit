<script setup lang="ts">
import { computed } from "vue";
import BaseDialog from "./BaseDialog.vue";
import UiButton from "./ui/UiButton.vue";
import DocTableBlock from "./DocTableBlock.vue";
import {
  asTableLookKey,
  builtInTableStyle,
  tableStyleGalleryEntries,
  TABLE_LOOK_FLAGS,
} from "../lib/tableStyles";
import { asTable } from "../lib/tables";
import { asNamedStyles } from "../lib/documentModel";

// The Table styles gallery (SPEC-076 AC-31..AC-35): five built-in looks,
// each tile a LIVE preview of the selected table with that look applied
// (AC-32) — the same DocTableBlock the canvas paints with, given a
// throwaway style catalogue, so a tile can never drift from what clicking
// it produces. Six checkboxes alongside toggle the table's own look flags
// (AC-35). Runtime prop declarations per the .vue espree constraint.
const props = defineProps({
  /** The table the gallery previews and applies to (documentModel.Table). */
  table: { type: Object, required: true },
  /** The document's style catalogue — the applied entry is marked selected
   * (AC-34) and a document style keeps its own definition in the preview. */
  styles: { type: Array, default: () => [] },
});

const emit = defineEmits(["apply", "toggle-look", "cancel"]);

const entries = computed(() => tableStyleGalleryEntries());

/** The catalogue a tile previews through: the document's own styles plus the
 * built-in this tile stands for, so a preview never needs the style to have
 * been materialised into the model first (AC-33 does that on click). */
function previewStyles(id = "") {
  const built = id ? builtInTableStyle(id) : undefined;
  const catalogue = asNamedStyles(props.styles).filter((style) => style.id !== id);
  return built ? [...catalogue, built] : catalogue;
}

/** The table as this tile would leave it — a shallow copy, never the model's
 * own object, so hovering the gallery cannot touch the document. */
function previewTable(entry = entries.value[0]) {
  const copy = { ...asTable(props.table) };
  if (entry.id === null) {
    delete copy.styleId;
    delete copy.look;
  } else {
    copy.styleId = entry.id;
    copy.look = { ...entry.look };
  }
  // Previews stay small: at most the first four rows of the real table.
  copy.rows = copy.rows.slice(0, 4);
  return copy;
}

const appliedId = computed(() => asTable(props.table).styleId ?? null);

function lookOn(key = "firstRow") {
  return !!asTable(props.table).look?.[asTableLookKey(key)];
}
</script>

<template>
  <BaseDialog
    title="Table styles"
    testid="table-style-gallery"
    width="34rem"
    @dismiss="emit('cancel')"
  >
    <div
      class="gallery"
      role="listbox"
      aria-label="Table styles"
    >
      <button
        v-for="(entry, index) in entries"
        :key="entry.id ?? 'none'"
        type="button"
        class="gallery-tile"
        role="option"
        :class="{ selected: appliedId === entry.id }"
        :aria-selected="appliedId === entry.id"
        :data-testid="`table-style-${entry.id ?? 'none'}`"
        :autofocus="index === 0"
        @click="emit('apply', entry.id)"
      >
        <span class="gallery-preview">
          <DocTableBlock
            :block="previewTable(entry)"
            :section="0"
            :block-index="0"
            :styles="previewStyles(entry.id ?? '')"
          />
        </span>
        <span class="gallery-name">{{ entry.name }}</span>
      </button>
    </div>
    <fieldset class="look-flags">
      <legend>Table style options</legend>
      <label
        v-for="flag in TABLE_LOOK_FLAGS"
        :key="flag.key"
        class="look-flag"
      >
        <input
          type="checkbox"
          :data-testid="`table-look-${flag.key}`"
          :checked="lookOn(flag.key)"
          @change="emit('toggle-look', flag.key)"
        >
        <span>{{ flag.label }}</span>
      </label>
    </fieldset>
    <template #actions>
      <UiButton
        variant="secondary"
        data-testid="table-style-close"
        @click="emit('cancel')"
      >
        Close
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
.gallery {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: var(--space-3);
}

.gallery-tile {
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
  align-items: stretch;
  padding: var(--space-2);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  font: inherit;
  cursor: pointer;
}

.gallery-tile.selected {
  outline: var(--space-1) solid var(--accent);
  outline-offset: -1px;
}

/* The preview is a real table shrunk to tile size: scaled, not restyled, so
   what a tile shows is what applying it paints (AC-32). */
.gallery-preview {
  display: block;
  overflow: hidden;
  height: 4.5rem;
  pointer-events: none;
}

.gallery-preview :deep(.doc-table) {
  width: 100% !important;
  font-size: var(--text-xs);
  transform-origin: top left;
}

.gallery-name {
  font-size: var(--text-md);
  color: var(--text-primary);
  text-align: center;
}

.look-flags {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: var(--space-2);
  margin: var(--space-4) 0 0;
  padding: var(--space-3);
  border: var(--border-width) solid var(--border-subtle);
  border-radius: var(--radius-sm);
}

.look-flag {
  display: flex;
  align-items: center;
  gap: var(--space-2);
  font-size: var(--text-md);
  color: var(--text-primary);
}
</style>
