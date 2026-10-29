<script setup lang="ts">
// The sheet tabs (SPEC-116; managed since SPEC-121): one tab per visible
// sheet, under the grid as in every spreadsheet. The ARIA tablist pattern:
// role="tablist"/"tab", aria-selected, a roving tabindex (only the selected
// tab is in the Tab order), and Left/Right/Home/End moving between tabs —
// selection follows focus, since showing a sheet is instant. Enter/Space
// activate, F2 or a double-click asks to rename, a right-click, Shift+F10 or
// the ContextMenu key asks for the tab's menu, and a pointer drag (past 4 px)
// reorders behind a drop indicator — Escape, a pointer cancel or a drop
// outside the bar cancels. While `renamingIndex` names a tab, an input
// stands in for it. The bar owns no sheet state: every gesture is an emit
// the view's useSheetCommands acts on. `tabs` is [{ index, id, name, color }]
// (index = the sheet's position in the model). SPEC-138: `readOnly` disables
// Add sheet. Runtime prop declarations per the espree constraint; typed
// helpers live in sheetTabBar.ts.
import { computed, nextTick, onBeforeUnmount, onMounted, ref, useTemplateRef, watch } from "vue";
import {
  asSheetTabItems,
  beyondDragThreshold,
  dropGapAt,
  gapX,
  insideRect,
  moveForDrop,
  slotSpans,
  tabElementAt,
  tabTint,
} from "./sheetTabBar";

const props = defineProps({
  tabs: { type: Array, default: () => [] },
  /** The model index of the sheet on show. */
  activeIndex: { type: Number, default: 0 },
  /** The id of the panel the tabs control. */
  panelId: { type: String, default: "" },
  /** The model index of the tab being renamed in place; -1 for none. */
  renamingIndex: { type: Number, default: -1 },
  renameDraft: { type: String, default: "" },
  /** The "Sheet names" message for an invalid draft; "" when valid. */
  renameError: { type: String, default: "" },
  /** Read out by the polite live region ("Sheet "Q1" moved to position 2"). */
  announcement: { type: String, default: "" },
  /** SPEC-138: no edit applies (another tab holds the workbook, a version is
   * previewed) — Add sheet is disabled. */
  readOnly: { type: Boolean, default: false },
});

const emit = defineEmits([
  "select",
  "add",
  "menu",
  "rename-start",
  "rename-input",
  "rename-commit",
  "rename-cancel",
  "rename-blur",
  "move",
]);

const bar = useTemplateRef("bar");
const list = useTemplateRef("list");
const sheetTabs = computed(() => asSheetTabItems(props.tabs));
const errorId = "sheet-tabs-name-error";

function focusTabAt(position = 0) {
  tabElementAt(list.value, position)?.focus();
}

/** Focuses the tab of the sheet at model index `index` (exposed, so the view
 * can hand focus back after a menu or dialog the tab opened). */
function focusTab(index = 0) {
  const position = sheetTabs.value.findIndex((tab) => tab.index === index);
  if (position >= 0) {
    focusTabAt(position);
  }
}

defineExpose({ focusTab });

async function selectAt(position = 0) {
  const tab = sheetTabs.value[position];
  if (!tab) {
    return;
  }
  emit("select", tab.index);
  await nextTick();
  focusTabAt(position);
}

// The ContextMenu key and Shift+F10 also make the browser fire `contextmenu`
// at the focused tab; the keyboard path has already asked for the menu, so
// that echo is ignored — until a pointer press, which only a real right-click
// makes.
const KEYBOARD_MENU_ECHO_MS = 500;
let keyboardMenuAt = 0;

function onTabKeydown(position = 0, event = new KeyboardEvent("keydown")) {
  const tab = sheetTabs.value[position];
  if (!tab || event.target !== event.currentTarget) {
    return;
  }
  const count = sheetTabs.value.length;
  const targets = {
    ArrowRight: (position + 1) % count,
    ArrowLeft: (position - 1 + count) % count,
    Home: 0,
    End: count - 1,
  };
  const target = Object.entries(targets).find(([key]) => key === event.key)?.[1];
  if (target !== undefined) {
    event.preventDefault();
    void selectAt(target);
  } else if (event.key === "Enter" || event.key === " ") {
    event.preventDefault();
    emit("select", tab.index);
  } else if (event.key === "F2") {
    event.preventDefault();
    emit("rename-start", tab.index);
  } else if (event.key === "ContextMenu" || (event.key === "F10" && event.shiftKey)) {
    event.preventDefault();
    const element = event.currentTarget;
    const rect = element instanceof HTMLElement ? element.getBoundingClientRect() : null;
    keyboardMenuAt = Date.now();
    emit("menu", { index: tab.index, x: rect ? rect.left : 0, y: rect ? rect.bottom : 0 });
  }
}

function onTabContextMenu(position = 0, event = new MouseEvent("contextmenu")) {
  event.preventDefault();
  const tab = sheetTabs.value[position];
  if (tab && Date.now() - keyboardMenuAt > KEYBOARD_MENU_ECHO_MS) {
    emit("menu", { index: tab.index, x: event.clientX, y: event.clientY });
  }
}

// ---- drag-reorder (pointer events; the listeners live on window while a
// press is down, so the drag follows the pointer anywhere) ------------------

const press = { active: false, position: -1, x: 0, y: 0, dragging: false, cancelled: false };
/** The gap the indicator marks while dragging; -1 hides it. */
const dropGap = ref(-1);
const indicatorLeft = ref(0);
/** A press that became a drag must not also select on its click. */
let suppressClick = false;

function onTabClick(position = 0) {
  if (suppressClick) {
    suppressClick = false;
    return;
  }
  const tab = sheetTabs.value[position];
  if (tab) {
    emit("select", tab.index);
  }
}

function onTabPointerDown(position = 0, event = new MouseEvent("pointerdown")) {
  // A pointer press starts a real (mouse) context menu: no keyboard echo.
  keyboardMenuAt = 0;
  if (event.button !== 0 || press.active) {
    return;
  }
  Object.assign(press, { active: true, position, x: event.clientX, y: event.clientY, dragging: false, cancelled: false });
  window.addEventListener("pointermove", onPointerMove);
  window.addEventListener("pointerup", onPointerUp);
  window.addEventListener("pointercancel", endPress);
  window.addEventListener("keydown", onDragKeydown, true);
}

function showIndicator(clientX = 0) {
  const element = list.value;
  if (!element) {
    return;
  }
  const spans = slotSpans(element);
  const gap = dropGapAt(clientX, spans);
  dropGap.value = gap;
  indicatorLeft.value = gapX(gap, spans) - element.getBoundingClientRect().left + element.scrollLeft;
}

function onPointerMove(event = new MouseEvent("pointermove")) {
  if (!press.active || press.cancelled) {
    return;
  }
  if (!press.dragging) {
    if (!beyondDragThreshold(event.clientX - press.x, event.clientY - press.y)) {
      return;
    }
    press.dragging = true;
  }
  showIndicator(event.clientX);
}

function onPointerUp(event = new MouseEvent("pointerup")) {
  if (!press.active) {
    return;
  }
  const { position, dragging, cancelled } = press;
  endPress();
  if (!dragging) {
    return;
  }
  suppressClick = true;
  setTimeout(() => {
    suppressClick = false;
  }, 0);
  const box = bar.value?.getBoundingClientRect();
  if (cancelled || !box || !insideRect(event.clientX, event.clientY, box)) {
    return;
  }
  const move = moveForDrop(sheetTabs.value, position, dropGapAt(event.clientX, slotSpans(list.value)));
  if (move) {
    emit("move", move);
  }
}

function onDragKeydown(event = new KeyboardEvent("keydown")) {
  if (event.key !== "Escape" || !press.dragging || press.cancelled) {
    return;
  }
  event.preventDefault();
  event.stopPropagation();
  press.cancelled = true;
  dropGap.value = -1;
}

function endPress() {
  press.active = false;
  press.dragging = false;
  press.cancelled = false;
  dropGap.value = -1;
  window.removeEventListener("pointermove", onPointerMove);
  window.removeEventListener("pointerup", onPointerUp);
  window.removeEventListener("pointercancel", endPress);
  window.removeEventListener("keydown", onDragKeydown, true);
}

onBeforeUnmount(endPress);

// ---- rename in place ---------------------------------------------------------

const errorLeft = ref(0);

function focusRenameInput() {
  const input = list.value?.querySelector('[data-testid="sheet-tab-rename-input"]');
  if (input instanceof HTMLInputElement) {
    input.focus();
    input.select();
  }
}

onMounted(() => {
  if (props.renamingIndex >= 0) {
    focusRenameInput();
  }
});

watch(
  () => props.renamingIndex,
  (index, previous) => {
    if (index >= 0) {
      focusRenameInput();
      return;
    }
    // The input took focus with it when it went: hand it back to the tab,
    // unless the rename ended because the user moved focus elsewhere.
    const active = document.activeElement;
    if (previous >= 0 && (active === null || active === document.body)) {
      focusTab(previous);
    }
  },
  { flush: "post" },
);

// The message sits above the bar (the strip scrolls sideways, which would
// clip it inside), aligned with the input.
watch(
  () => [props.renameError, props.renamingIndex],
  () => {
    const slot = list.value?.querySelector(".sheet-tab-rename");
    const box = bar.value?.getBoundingClientRect();
    errorLeft.value = slot && box ? Math.max(0, slot.getBoundingClientRect().left - box.left) : 0;
  },
  { flush: "post" },
);

function onRenameInput(event = new Event("input")) {
  const target = event.target;
  if (target instanceof HTMLInputElement) {
    emit("rename-input", target.value);
  }
}

function onRenameKeydown(event = new KeyboardEvent("keydown")) {
  if (event.key === "Enter") {
    event.preventDefault();
    emit("rename-commit");
  } else if (event.key === "Escape") {
    event.preventDefault();
    event.stopPropagation();
    emit("rename-cancel");
  }
}

// A focused input removed by a commit or cancel can still report a blur in
// some browsers; once the rename is over there is nothing to blur.
function onRenameBlur() {
  if (props.renamingIndex >= 0) {
    emit("rename-blur");
  }
}
</script>

<template>
  <div
    ref="bar"
    class="sheet-tab-bar"
    :class="{ 'is-dragging': dropGap >= 0 }"
  >
    <div
      ref="list"
      class="sheet-tabs"
      role="tablist"
      aria-label="Sheets"
      aria-orientation="horizontal"
      data-testid="sheet-tabs"
    >
      <template
        v-for="(tab, position) in sheetTabs"
        :key="tab.id"
      >
        <div
          v-if="tab.index === props.renamingIndex"
          class="sheet-tab-rename"
          :data-tab-position="position"
        >
          <input
            class="sheet-tab-rename-input"
            data-testid="sheet-tab-rename-input"
            type="text"
            aria-label="Sheet name"
            autocomplete="off"
            spellcheck="false"
            :value="props.renameDraft"
            :aria-invalid="props.renameError ? 'true' : undefined"
            :aria-describedby="props.renameError ? errorId : undefined"
            @input="onRenameInput"
            @keydown="onRenameKeydown"
            @blur="onRenameBlur"
          >
        </div>
        <button
          v-else
          :id="`sheet-tab-${tab.id}`"
          type="button"
          class="sheet-tab"
          :class="{ 'is-selected': tab.index === props.activeIndex }"
          role="tab"
          data-testid="sheet-tab"
          :data-tab-position="position"
          :data-tab-color="tab.color || undefined"
          :aria-selected="tab.index === props.activeIndex"
          :aria-controls="props.panelId || undefined"
          :tabindex="tab.index === props.activeIndex ? 0 : -1"
          :style="tabTint(tab.color, tab.index === props.activeIndex)"
          @click="onTabClick(position)"
          @dblclick="emit('rename-start', tab.index)"
          @keydown="onTabKeydown(position, $event)"
          @contextmenu="onTabContextMenu(position, $event)"
          @pointerdown="onTabPointerDown(position, $event)"
        >
          <span class="sheet-tab-name">{{ tab.name }}</span>
          <span
            v-if="tab.color && tab.index === props.activeIndex"
            class="sheet-tab-color-bar"
            aria-hidden="true"
            :style="{ background: tab.color }"
          />
        </button>
      </template>
      <div
        v-if="dropGap >= 0"
        class="sheet-tab-drop-indicator"
        data-testid="sheet-tab-drop-indicator"
        aria-hidden="true"
        :style="{ left: `${indicatorLeft}px` }"
      />
    </div>
    <button
      type="button"
      class="sheet-add"
      data-testid="sheet-add"
      aria-label="Add sheet"
      title="Add sheet"
      :disabled="props.readOnly"
      @click="emit('add')"
    >
      +
    </button>
    <div
      v-if="props.renamingIndex >= 0 && props.renameError"
      :id="errorId"
      class="sheet-name-error"
      data-testid="sheet-name-error"
      role="alert"
      :style="{ left: `${errorLeft}px` }"
    >
      {{ props.renameError }}
    </div>
    <div
      class="visually-hidden"
      aria-live="polite"
      data-testid="sheet-tabs-announcement"
    >
      {{ props.announcement }}
    </div>
  </div>
</template>

<style scoped>
.sheet-tab-bar {
  position: relative;
  display: flex;
  flex: none;
  align-items: center;
  gap: var(--space-1);
  padding: var(--space-1) var(--space-3);
  background: var(--surface-sunken);
  border-top: var(--border-width) solid var(--border-default);
}

.sheet-tab-bar.is-dragging {
  cursor: grabbing;
}

/* The strip scrolls sideways when the tabs outgrow the bar; the Add button
   stays pinned after it. */
.sheet-tabs {
  position: relative;
  display: flex;
  flex: 0 1 auto;
  min-width: 0;
  gap: var(--space-1);
  overflow-x: auto;
}

.sheet-tab {
  position: relative;
  flex: none;
  max-width: 16rem;
  padding: var(--space-1) var(--space-4);
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  border: var(--border-width) solid transparent;
  border-radius: var(--radius-sm);
  background: transparent;
  color: var(--text-secondary);
  font: inherit;
  font-size: var(--text-sm);
  cursor: pointer;
  user-select: none;
}

.sheet-tab:hover {
  background: var(--state-hover);
}

.sheet-tab.is-selected {
  background: var(--surface-raised);
  border-color: var(--border-default);
  color: var(--text-primary);
  font-weight: var(--weight-medium);
}

.sheet-tab:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: 1px;
}

/* The tab colour on the tab on show: a bar along its foot (the others are
   tinted inline instead). */
.sheet-tab-color-bar {
  position: absolute;
  left: 0;
  right: 0;
  bottom: 0;
  height: 3px;
  pointer-events: none;
}

.sheet-tab-rename {
  flex: none;
  display: flex;
  align-items: center;
}

.sheet-tab-rename-input {
  width: 10rem;
  padding: var(--space-1) var(--space-2);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  color: var(--text-primary);
  font: inherit;
  font-size: var(--text-sm);
}

.sheet-tab-rename-input:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: 1px;
}

.sheet-tab-rename-input[aria-invalid="true"] {
  border-color: var(--danger);
}

.sheet-name-error {
  position: absolute;
  bottom: 100%;
  z-index: 20;
  max-width: 24rem;
  margin-bottom: var(--space-1);
  padding: var(--space-1) var(--space-2);
  border: var(--border-width) solid var(--danger);
  border-radius: var(--radius-sm);
  background: var(--surface-raised);
  box-shadow: var(--shadow-2);
  color: var(--text-danger);
  font-size: var(--text-sm);
}

/* The 2 px insertion bar a drag drops at, centred on the gap. */
.sheet-tab-drop-indicator {
  position: absolute;
  top: 0;
  bottom: 0;
  width: 2px;
  transform: translateX(-50%);
  background: var(--accent);
  pointer-events: none;
}

.sheet-add {
  flex: none;
  padding: var(--space-1) var(--space-3);
  border: var(--border-width) solid transparent;
  border-radius: var(--radius-sm);
  background: transparent;
  color: var(--text-secondary);
  font: inherit;
  font-size: var(--text-md);
  line-height: 1;
  cursor: pointer;
}

.sheet-add:hover:not(:disabled) {
  background: var(--state-hover);
  color: var(--text-primary);
}

.sheet-add:disabled {
  color: var(--text-disabled);
  cursor: default;
}

.sheet-add:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: 1px;
}

.visually-hidden {
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
