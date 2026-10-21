<script setup lang="ts">
// The in-cell dropdown (SPEC-130): the `validation-dropdown` button against
// the active cell's right edge (aria-label "Choose from list",
// aria-haspopup="listbox", aria-expanded) and, while expanded, the
// `validation-list` listbox under the cell — at least the cell's width, eight
// rows visible then scrolling — of `validation-option`s (aria-selected on the
// selected one, which aria-activedescendant names and which is scrolled into
// view; a blank item is an empty row), or one disabled
// `validation-list-unavailable` option when the source cannot be listed.
// Placement is validationChrome's, from the active cell's rect in unscaled
// paper px (SheetGrid's validation layer), so both follow scroll and zoom.
//
// The list takes focus when it opens; its keys are validationChrome's
// listKeyAction (Up/Down, Home/End, PageUp/PageDown, a typed character,
// Enter/Tab, Escape) and stop at the list — none reaches the grid. It emits:
// `toggle` (the button), `select(index)` (a key moved the selection),
// `pick(index)` (Enter, Tab or a click), `close` (Escape: the caller returns
// focus to the grid) and `dismiss` (focus left the list). A press on either
// never reaches the grid's pointer handling, and the button's press keeps
// focus where it is. Runtime prop declarations per the espree constraint
// (see AppShell.vue).
import { computed, nextTick, onMounted, useTemplateRef, watch } from "vue";
import { LIST_UNAVAILABLE_TEXT, listKeyAction, placeDropdownButton, placeValidationList, VALIDATION_LIST_ID } from "../../views/spreadsheet/validationChrome";
import AppIcon from "../AppIcon.vue";

const props = defineProps({
  /** The active cell's rect (left, top, width, height), paper px. */
  cellRect: { type: Object, required: true },
  /** The cell area's size (width, height), paper px. */
  bounds: { type: Object, required: true },
  /** Whether the list is open. */
  expanded: { type: Boolean, default: false },
  /** The items' texts, in list order. */
  items: { type: Array, default: () => [] },
  /** The selected item's index (-1: none). */
  selected: { type: Number, default: -1 },
  /** The source cannot be listed. */
  unavailable: { type: Boolean, default: false },
});

const emit = defineEmits(["toggle", "select", "pick", "close", "dismiss"]);

const root = useTemplateRef("root");
const list = useTemplateRef("list");

const texts = computed(() => props.items.map((item) => String(item ?? "")));
const cell = computed(() => {
  const rect = props.cellRect;
  return { left: Number(rect.left) || 0, top: Number(rect.top) || 0, width: Number(rect.width) || 0, height: Number(rect.height) || 0 };
});
const area = computed(() => ({ width: Number(props.bounds.width) || 0, height: Number(props.bounds.height) || 0 }));
const buttonStyle = computed(() => {
  const box = placeDropdownButton(cell.value, area.value);
  return { left: `${box.left}px`, top: `${box.top}px`, width: `${box.width}px`, height: `${box.height}px` };
});
const listStyle = computed(() => {
  const box = placeValidationList(cell.value, props.unavailable ? 1 : texts.value.length, area.value);
  return { left: `${box.left}px`, top: `${box.top}px`, minWidth: `${box.minWidth}px`, maxHeight: `${box.maxHeight}px` };
});
const activeDescendant = computed(() => (props.selected >= 0 && props.selected < texts.value.length ? `validation-option-${props.selected}` : undefined));

/** Focus to the list as it opens; the selected option into view. */
async function reveal(focus = false) {
  await nextTick();
  const el = list.value;
  if (!el) {
    return;
  }
  if (focus) {
    el.focus({ preventScroll: true });
  }
  const option = activeDescendant.value ? el.querySelector(`#${activeDescendant.value}`) : null;
  if (option && typeof option.scrollIntoView === "function") {
    option.scrollIntoView({ block: "nearest" });
  }
}

onMounted(() => {
  if (props.expanded) {
    void reveal(true);
  }
});
watch(
  () => props.expanded,
  (open) => {
    if (open) {
      void reveal(true);
    }
  },
);
watch(
  () => props.selected,
  () => void reveal(false),
);

function onKeydown(event = new KeyboardEvent("keydown")) {
  event.stopPropagation();
  const action = listKeyAction(event, texts.value, props.selected);
  if (!action) {
    return;
  }
  event.preventDefault();
  if (action.kind === "close") {
    emit("close");
  } else if (action.kind === "pick") {
    emit("pick", action.index);
  } else {
    emit("select", action.index);
  }
}

/** Focus leaving the button and the list closes it, focus left there. */
function onFocusOut(event = new FocusEvent("focusout")) {
  const next = event.relatedTarget;
  if (!(next instanceof Node && root.value && root.value.contains(next))) {
    emit("dismiss");
  }
}
</script>

<template>
  <div
    ref="root"
    class="validation-dropdown-root"
    @pointerdown.stop
    @mousedown.stop
    @dblclick.stop
    @contextmenu.stop
  >
    <button
      type="button"
      class="validation-dropdown"
      data-testid="validation-dropdown"
      aria-label="Choose from list"
      aria-haspopup="listbox"
      :aria-expanded="props.expanded ? 'true' : 'false'"
      :aria-controls="props.expanded ? VALIDATION_LIST_ID : undefined"
      tabindex="-1"
      :style="buttonStyle"
      @mousedown.prevent
      @click.stop="emit('toggle')"
    >
      <AppIcon name="chevron-down" />
    </button>
    <div
      v-if="props.expanded"
      :id="VALIDATION_LIST_ID"
      ref="list"
      class="validation-list"
      role="listbox"
      aria-label="Choose from list"
      data-testid="validation-list"
      tabindex="0"
      :aria-activedescendant="activeDescendant"
      :style="listStyle"
      @keydown="onKeydown"
      @focusout="onFocusOut"
    >
      <div
        v-if="props.unavailable"
        class="validation-option is-unavailable"
        role="option"
        aria-disabled="true"
        aria-selected="false"
        data-testid="validation-list-unavailable"
      >
        {{ LIST_UNAVAILABLE_TEXT }}
      </div>
      <template v-else>
        <div
          v-for="(text, index) in texts"
          :id="`validation-option-${index}`"
          :key="index"
          class="validation-option"
          :class="{ 'is-selected': index === props.selected }"
          role="option"
          data-testid="validation-option"
          :aria-selected="index === props.selected ? 'true' : 'false'"
          @mousedown.prevent
          @click="emit('pick', index)"
        >
          {{ text }}
        </div>
      </template>
    </div>
  </div>
</template>

<style scoped>
/* The button and the list are placed in the validation layer's paper px. */
.validation-dropdown {
  position: absolute;
  box-sizing: border-box;
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 0;
  pointer-events: auto;
  color: var(--text-primary);
  background: var(--surface-raised);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  cursor: default;
}

.validation-dropdown :deep(svg) {
  width: 100%;
  height: auto;
  max-height: 100%;
}

.validation-dropdown:hover,
.validation-dropdown[aria-expanded="true"] {
  background: var(--surface-sunken);
}

.validation-list {
  position: absolute;
  box-sizing: border-box;
  overflow-y: auto;
  pointer-events: auto;
  color: var(--text-primary);
  background: var(--surface-raised);
  border: var(--border-width) solid var(--border-default);
  box-shadow: var(--shadow-2);
  font-family: var(--font-ui);
  font-size: var(--text-sm);
}

.validation-list:focus-visible {
  outline: var(--border-width) solid var(--focus-ring);
}

/* One row per item (LIST_ROW_PX tall), a blank item an empty row. */
.validation-option {
  box-sizing: border-box;
  height: 20px;
  line-height: 20px;
  padding: 0 var(--space-2);
  white-space: pre;
  overflow: hidden;
  text-overflow: ellipsis;
  cursor: default;
}

.validation-option:hover {
  background: var(--surface-sunken);
}

.validation-option.is-selected {
  background: var(--accent-subtle, var(--surface-sunken));
  font-weight: var(--weight-button);
}

.validation-option.is-unavailable {
  color: var(--text-secondary);
  font-style: italic;
}
</style>
