<script setup lang="ts">
// The slide sorter (SPEC-099): a grid of every slide under its section
// headers, swapped in for the filmstrip + stage by the ribbon's Sorter toggle.
// One `ul[role=listbox][aria-multiselectable]` holds the section header rows
// and the slide cards (each the shared SlideThumb at 208 px); clicks and the
// selection keys go through the shared useSlideSelection, a drag moves the
// whole selection as a block behind one drop indicator, and each card's menu
// is the filmstrip's own slide menu (SLIDE_MENU_ITEMS, emitted as the same
// `command` the filmstrip emits). Everything that changes the deck runs
// through the useSlideSorter controller. Runtime prop/emit declarations, not
// the generic forms (espree lints .vue scripts).
import { computed, nextTick, ref, useTemplateRef } from "vue";
import SlideThumb from "./SlideThumb.vue";
import SlideContextMenu from "./SlideContextMenu.vue";
import {
  asStageFieldsBySlide,
  filmstripLabel,
  filmstripThumbScale,
  gestureModifiers,
  SLIDE_MENU_ITEMS,
} from "./slideRender";
import { MAX_SECTION_NAME_LENGTH } from "../../lib/presentation/sections";
import { eventValue } from "../../views/presentation/deckSetupFields";
import { asIndexes, asSorterGroups, focusAndSelect } from "../../views/presentation/sorterFields";

const props = defineProps({
  /** The useSlideSorter() controller. */
  controller: { type: Object, required: true },
  widthCm: { type: Number, required: true },
  heightCm: { type: Number, required: true },
  currentIndex: { type: Number, default: 0 },
  media: { type: Array, default: () => [] },
  theme: { type: Object, default: null },
  /** SPEC-093: each slide's pre-resolved field band, indexed to the slides. */
  fieldsBySlide: { type: Array, default: () => [] },
});

const emit = defineEmits(["command"]);

const c = props.controller;

/** Target card width in px (AC-3). */
const CARD_WIDTH_PX = 208;

const SECTION_MENU_ITEMS = [
  { action: "add", label: "Add section", testid: "section-add" },
  { action: "rename", label: "Rename section", testid: "section-rename" },
  { action: "move-up", label: "Move section up", testid: "section-move-up" },
  { action: "move-down", label: "Move section down", testid: "section-move-down" },
  { action: "remove", label: "Remove section", testid: "section-remove" },
  { action: "remove-with-slides", label: "Remove section and slides", testid: "section-remove-with-slides" },
];

const scale = computed(() => filmstripThumbScale(props.widthCm, CARD_WIDTH_PX));
const fields = computed(() => asStageFieldsBySlide(props.fieldsBySlide));
const groups = computed(() => asSorterGroups(c.rows.value));
const total = computed(() => groups.value.reduce((sum, group) => sum + group.cards.length, 0));
const selected = computed(() => asIndexes(c.selection.indexes.value));

const listRef = useTemplateRef("list");

// The card / section whose menu is open, and where.
const cardMenu = ref(-1);
const sectionMenu = ref("");
const menuPos = ref({ x: 0, y: 0 });

// Drag state: the block is the selection; the drop is an insertion point plus
// the section it lands in, shown by the one indicator.
const dragging = ref(false);
const drop = ref({ index: -1, key: "", edge: "before" });

function optionId(index = 0) {
  return `sorter-option-${index}`;
}

function cardLabel(card = groups.value[0].cards[0]) {
  return filmstripLabel(card.slide, card.index, total.value);
}

function onCardClick(index = 0, event = new MouseEvent("click")) {
  c.selectAt(index, event);
  listRef.value?.focus();
}

// --- menus -------------------------------------------------------------------

function openCardMenu(index = 0, event = new MouseEvent("contextmenu")) {
  event.preventDefault();
  if (!selected.value.includes(index)) {
    c.selection.click(index, gestureModifiers(event));
  }
  sectionMenu.value = "";
  cardMenu.value = index;
  menuPos.value = { x: event.clientX, y: event.clientY };
}

function runCardMenu(action = "") {
  const index = cardMenu.value;
  cardMenu.value = -1;
  if (index >= 0) {
    emit("command", { action, index });
  }
}

function openSectionMenu(key = "", event = new MouseEvent("click")) {
  cardMenu.value = -1;
  sectionMenu.value = key;
  menuPos.value = { x: event.clientX, y: event.clientY };
}

function closeMenus(refocus = false) {
  cardMenu.value = -1;
  sectionMenu.value = "";
  if (refocus) {
    void nextTick(() => listRef.value?.focus());
  }
}

function runSectionMenu(action = "") {
  const key = sectionMenu.value;
  sectionMenu.value = "";
  const group = groups.value.find((entry) => entry.key === key);
  if (!group) {
    return;
  }
  if (action === "add") {
    c.addSection(selected.value[0] ?? group.startIndex);
  } else if (action === "rename") {
    startRename(key);
  } else if (action === "move-up" || action === "move-down") {
    c.moveSection(key, action === "move-up" ? -1 : 1);
  } else if (action === "remove" || action === "remove-with-slides") {
    c.removeSection(key, action === "remove" ? "sectionOnly" : "withSlides");
  }
}

// --- section rename ------------------------------------------------------------

const draft = ref("");
const renameError = ref(false);

function startRename(key = "") {
  const group = groups.value.find((entry) => entry.key === key);
  draft.value = group ? group.name : "";
  renameError.value = false;
  c.renaming.value = key;
}

function onRenameKeydown(key = "", event = new KeyboardEvent("keydown")) {
  if (event.key === "Enter") {
    event.preventDefault();
    if (c.renameSection(key, draft.value)) {
      void nextTick(() => listRef.value?.focus());
    } else {
      // Empty after trim: rejected, the input stays open.
      renameError.value = true;
      c.renaming.value = key;
    }
  } else if (event.key === "Escape") {
    event.preventDefault();
    event.stopPropagation();
    c.renaming.value = null;
  }
}

function isRenaming(key = "") {
  return c.renaming.value === key;
}

function onRenameInput(event = new Event("input")) {
  draft.value = eventValue(event);
  renameError.value = false;
}

// --- drag reorder ------------------------------------------------------------

function onDragStart(index = 0, event = new DragEvent("dragstart")) {
  if (!selected.value.includes(index)) {
    c.selection.click(index);
  }
  dragging.value = true;
  if (event.dataTransfer) {
    event.dataTransfer.effectAllowed = "move";
    event.dataTransfer.setData("text/plain", selected.value.map((i) => i + 1).join(","));
  }
}

function onCardDragOver(card = { index: 0 }, key = "", event = new DragEvent("dragover")) {
  if (!dragging.value) {
    return;
  }
  event.preventDefault();
  const target = event.currentTarget;
  const after =
    target instanceof HTMLElement && event.clientX > target.getBoundingClientRect().left + target.getBoundingClientRect().width / 2;
  drop.value = { index: after ? card.index + 1 : card.index, key, edge: after ? "after" : "before" };
}

function onHeaderDragOver(group = groups.value[0], event = new DragEvent("dragover")) {
  if (!dragging.value) {
    return;
  }
  event.preventDefault();
  drop.value = { index: group.startIndex, key: group.key ?? "", edge: "header" };
}

function onDrop(event = new DragEvent("drop")) {
  event.preventDefault();
  if (dragging.value && drop.value.index >= 0) {
    c.reorder(drop.value.index, drop.value.key || undefined);
  }
  endDrag();
}

function endDrag() {
  // Fires whether or not a drop landed — dropping outside cancels.
  dragging.value = false;
  drop.value = { index: -1, key: "", edge: "before" };
}

function indicatorOn(index = 0, edge = "before") {
  return dragging.value && drop.value.edge === edge && drop.value.index === (edge === "after" ? index + 1 : index);
}
</script>

<template>
  <div
    class="slide-sorter"
    data-testid="slides-sorter"
  >
    <div
      class="sorter-bar"
      role="toolbar"
      aria-label="Slide sorter"
    >
      <button
        type="button"
        data-testid="sorter-copy"
        @click="c.copy()"
      >
        Copy
      </button>
      <button
        type="button"
        data-testid="sorter-cut"
        @click="c.cut()"
      >
        Cut
      </button>
      <button
        type="button"
        data-testid="sorter-paste"
        :disabled="!c.canPaste.value"
        @click="c.paste()"
      >
        Paste
      </button>
      <button
        type="button"
        data-testid="sorter-duplicate"
        @click="c.duplicate()"
      >
        Duplicate
      </button>
      <button
        type="button"
        data-testid="sorter-delete"
        @click="c.remove()"
      >
        Delete
      </button>
      <button
        type="button"
        data-testid="sorter-add-section"
        @click="c.addSection()"
      >
        Add section
      </button>
      <button
        type="button"
        data-testid="slides-reuse-toggle"
        aria-haspopup="dialog"
        @click="c.reuse.show()"
      >
        Reuse slides…
      </button>
      <span class="sorter-count">{{ selected.length }} of {{ total }} selected</span>
    </div>

    <p
      v-if="c.odpNote.value"
      class="sorter-note"
      data-testid="sorter-odp-note"
    >
      Sections are saved with this deck, but LibreOffice Impress won't display them.
    </p>

    <div
      v-if="c.report.value"
      class="sorter-report"
      role="status"
      data-testid="sorter-import-report"
    >
      <strong>{{ c.report.value.title }}</strong>
      <ul>
        <li
          v-for="line in c.report.value.lines"
          :key="line"
        >
          {{ line }}
        </li>
      </ul>
      <button
        type="button"
        data-testid="sorter-import-report-dismiss"
        @click="c.dismissReport()"
      >
        Dismiss
      </button>
    </div>

    <ul
      ref="list"
      class="sorter-list"
      role="listbox"
      tabindex="0"
      aria-label="Slides"
      aria-multiselectable="true"
      :aria-activedescendant="total > 0 ? optionId(currentIndex) : undefined"
      data-testid="sorter-listbox"
      @drop="onDrop"
      @dragover.prevent
    >
      <template
        v-for="group in groups"
        :key="group.key ?? 'deck'"
      >
        <li
          v-if="group.key !== null"
          class="sorter-section"
          :class="{ 'drop-header': indicatorOn(group.startIndex, 'header') }"
          role="none"
          data-testid="sorter-section"
          :data-section-key="group.key"
          @dragover="onHeaderDragOver(group, $event)"
        >
          <button
            type="button"
            class="sorter-section-toggle"
            data-testid="sorter-section-toggle"
            :aria-expanded="group.collapsed ? 'false' : 'true'"
            :aria-label="`${group.collapsed ? 'Expand' : 'Collapse'} section ${group.name}`"
            @click="c.toggleCollapsed(group.key)"
          >
            {{ group.collapsed ? "▸" : "▾" }}
          </button>
          <input
            v-if="isRenaming(group.key)"
            :ref="focusAndSelect"
            class="section-rename-input"
            data-testid="section-rename-input"
            :aria-label="`Rename section ${group.name}`"
            :aria-invalid="renameError ? 'true' : undefined"
            :maxlength="MAX_SECTION_NAME_LENGTH"
            :value="draft"
            @input="onRenameInput"
            @keydown="onRenameKeydown(group.key, $event)"
            @blur="c.renaming.value = null"
          >
          <span
            v-else
            class="sorter-section-name"
            data-testid="sorter-section-name"
            @dblclick="startRename(group.key)"
          >{{ group.name }}</span>
          <span
            class="sorter-section-count"
            data-testid="sorter-section-count"
          >{{ group.cards.length === 1 ? "1 slide" : `${group.cards.length} slides` }}</span>
          <button
            type="button"
            class="sorter-section-menu-button"
            data-testid="sorter-section-menu"
            aria-haspopup="menu"
            :aria-expanded="sectionMenu === group.key"
            :aria-label="`Section options for ${group.name}`"
            @click.stop="openSectionMenu(group.key, $event)"
          >
            ⋯
          </button>
          <span
            v-if="indicatorOn(group.startIndex, 'header')"
            class="sorter-drop-indicator header"
            data-testid="sorter-drop-indicator"
          />
        </li>
        <template v-if="!group.collapsed">
          <li
            v-for="card in group.cards"
            :id="optionId(card.index)"
            :key="card.slide.id"
            class="sorter-slide"
            :class="{ current: card.index === currentIndex, selected: selected.includes(card.index) }"
            role="option"
            :aria-selected="selected.includes(card.index)"
            :aria-label="cardLabel(card)"
            data-testid="sorter-slide"
            draggable="true"
            @click="onCardClick(card.index, $event)"
            @contextmenu="openCardMenu(card.index, $event)"
            @dragstart="onDragStart(card.index, $event)"
            @dragover="onCardDragOver(card, group.key ?? '', $event)"
            @dragend="endDrag"
          >
            <SlideThumb
              :slide="card.slide"
              :index="card.index"
              :width-cm="widthCm"
              :height-cm="heightCm"
              :media="media"
              :scale="scale"
              :fields="fields[card.index] || []"
              :theme="theme"
              stacked
            />
            <span
              v-if="indicatorOn(card.index, 'before') || indicatorOn(card.index, 'after')"
              class="sorter-drop-indicator"
              :class="indicatorOn(card.index, 'after') ? 'after' : 'before'"
              data-testid="sorter-drop-indicator"
            />
          </li>
        </template>
      </template>
    </ul>

    <SlideContextMenu
      v-if="cardMenu >= 0"
      :items="SLIDE_MENU_ITEMS"
      :x="menuPos.x"
      :y="menuPos.y"
      @run="runCardMenu"
      @close="closeMenus"
    />
    <SlideContextMenu
      v-if="sectionMenu !== ''"
      :items="SECTION_MENU_ITEMS"
      :x="menuPos.x"
      :y="menuPos.y"
      testid="sorter-section-context-menu"
      backdrop-testid="sorter-section-menu-backdrop"
      label="Section options"
      @run="runSectionMenu"
      @close="closeMenus"
    />

    <div
      class="sr-only"
      role="status"
      aria-live="polite"
      data-testid="sorter-live"
    >
      {{ c.announcement.value }}
    </div>
  </div>
</template>

<style scoped>
.slide-sorter {
  flex: 1 1 auto;
  min-height: 0;
  display: flex;
  flex-direction: column;
  overflow: hidden;
}

.sorter-bar {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: var(--space-2);
  padding: var(--space-2) var(--space-3);
  border-bottom: var(--border-width) solid var(--border-subtle);
}

.sorter-bar button,
.sorter-report button {
  padding: var(--space-1) var(--space-2);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  background: var(--surface-raised);
  color: var(--text-primary);
  cursor: pointer;
}

.sorter-bar button:disabled {
  color: var(--text-secondary);
  cursor: default;
}

.sorter-count {
  margin-inline-start: auto;
  color: var(--text-secondary);
  font-size: var(--text-sm);
}

.sorter-note,
.sorter-report {
  margin: var(--space-2) var(--space-3) 0;
  padding: var(--space-2) var(--space-3);
  border: var(--border-width) solid var(--border-subtle);
  border-radius: var(--radius-sm);
  background: var(--surface-raised);
  color: var(--text-primary);
  font-size: var(--text-sm);
}

.sorter-report ul {
  margin: var(--space-1) 0;
  padding-inline-start: var(--space-4);
}

.sorter-list {
  flex: 1 1 auto;
  min-height: 0;
  overflow-y: auto;
  display: flex;
  flex-wrap: wrap;
  align-content: flex-start;
  gap: var(--space-3);
  margin: 0;
  padding: var(--space-3);
  list-style: none;
  outline: none;
}

.sorter-list:focus-visible {
  outline: var(--border-width) solid var(--focus-ring);
  outline-offset: calc(-1 * var(--border-width));
}

.sorter-section {
  position: relative;
  flex: 1 0 100%;
  display: flex;
  align-items: center;
  gap: var(--space-2);
  padding: var(--space-1) 0;
  border-bottom: var(--border-width) solid var(--border-subtle);
  color: var(--text-primary);
  font-weight: var(--weight-bold);
}

.sorter-section.drop-header {
  border-bottom-color: var(--accent);
}

.sorter-section-toggle,
.sorter-section-menu-button {
  border: 0;
  background: transparent;
  color: var(--text-secondary);
  cursor: pointer;
  line-height: 1;
  padding: 0 var(--space-1);
}

.sorter-section-count {
  color: var(--text-secondary);
  font-size: var(--text-sm);
  font-weight: var(--weight-medium);
}

.section-rename-input {
  font: inherit;
  padding: 0 var(--space-1);
}

.section-rename-input[aria-invalid="true"] {
  outline: var(--border-width) solid var(--danger);
}

.sorter-slide {
  position: relative;
  padding: var(--space-1);
  border: var(--border-width) solid transparent;
  border-radius: var(--radius-sm);
  cursor: pointer;
}

.sorter-slide.selected {
  border-color: var(--accent);
  background: var(--surface-raised);
}

.sorter-slide.current {
  outline: var(--border-width) solid var(--accent);
}

.sorter-drop-indicator {
  position: absolute;
  top: 0;
  bottom: 0;
  width: 3px;
  border-radius: var(--radius-sm);
  background: var(--accent);
  pointer-events: none;
}

.sorter-drop-indicator.before {
  left: calc(-1 * var(--space-2));
}

.sorter-drop-indicator.after {
  right: calc(-1 * var(--space-2));
}

.sorter-drop-indicator.header {
  top: auto;
  bottom: -2px;
  left: 0;
  right: 0;
  width: auto;
  height: 3px;
}

.sr-only {
  position: absolute;
  width: 1px;
  height: 1px;
  padding: 0;
  margin: calc(-1 * var(--border-width));
  overflow: hidden;
  clip: rect(0, 0, 0, 0);
  white-space: nowrap;
  border: 0;
}
</style>
