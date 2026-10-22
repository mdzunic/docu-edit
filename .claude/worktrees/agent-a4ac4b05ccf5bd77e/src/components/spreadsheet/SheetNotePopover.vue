<script setup lang="ts">
// The hover popover of an annotated cell (SPEC-135): shown once the pointer
// rests on the cell for NOTE_HOVER_DELAY_MS (useCellNotes). A note's part,
// `sheet-note-popover` (role="tooltip"), shows its author, then its text; a
// thread's, `sheet-thread-popover`, its opening message's author and text,
// "N replies" and `sheet-thread-open`, which opens the comments panel on that
// thread. Every text is plain text — interpolated, never markup. SheetNoteLayer
// places it beside its cell (paper px, so it follows scroll and zoom); it
// takes the pointer — the grid's own handling never sees a press on it — and
// reports the pointer entering and leaving it (`data-note-popover` is what
// useSheetGrid's notePointer recognises). Runtime prop declarations per the
// espree constraint (see AppShell.vue).
import { computed } from "vue";
import { repliesLabel } from "../../views/spreadsheet/useCellNotes";

const props = defineProps({
  /** The note part: { author, text }; absent = none. */
  note: { type: Object, default: undefined },
  /** The thread part: { author, text, replies, resolved }; absent = none. */
  thread: { type: Object, default: undefined },
  /** The cell's address, for the accessible names. */
  address: { type: String, default: "" },
  left: { type: Number, default: 0 },
  top: { type: Number, default: 0 },
  width: { type: Number, default: 240 },
});

const emit = defineEmits(["open-thread", "enter", "leave"]);

/** A part's field as text (a missing one is ""). */
function textOf(part = Object(), field = "") {
  const value = part ? Reflect.get(part, field) : "";
  return typeof value === "string" ? value : "";
}

const replies = computed(() => {
  const count = props.thread ? Number(Reflect.get(props.thread, "replies")) : 0;
  return repliesLabel(Number.isFinite(count) ? count : 0);
});
const resolved = computed(() => props.thread !== undefined && Reflect.get(props.thread, "resolved") === true);
</script>

<template>
  <div
    class="sheet-note-popover-stack"
    data-note-popover="true"
    :style="{ left: `${props.left}px`, top: `${props.top}px`, width: `${props.width}px` }"
    @pointerenter="emit('enter')"
    @pointerleave="emit('leave')"
    @pointerdown.stop
    @mousedown.stop
    @click.stop
    @dblclick.stop
    @contextmenu.stop
  >
    <div
      v-if="props.note"
      class="sheet-popover-card"
      role="tooltip"
      data-testid="sheet-note-popover"
    >
      <div
        v-if="textOf(props.note, 'author')"
        class="sheet-popover-author"
        data-testid="sheet-note-popover-author"
      >
        {{ textOf(props.note, "author") }}
      </div>
      <div
        class="sheet-popover-text"
        data-testid="sheet-note-popover-text"
      >
        {{ textOf(props.note, "text") }}
      </div>
    </div>
    <div
      v-if="props.thread"
      class="sheet-popover-card is-thread"
      :class="{ 'is-resolved': resolved }"
      role="group"
      :aria-label="`Comment on ${props.address}`"
      data-testid="sheet-thread-popover"
      :data-resolved="resolved ? 'true' : undefined"
    >
      <div class="sheet-popover-author">
        {{ textOf(props.thread, "author") }}
      </div>
      <div
        class="sheet-popover-text"
        data-testid="sheet-thread-popover-text"
      >
        {{ textOf(props.thread, "text") }}
      </div>
      <div class="sheet-popover-footer">
        <span data-testid="sheet-thread-popover-replies">{{ replies }}</span>
        <span v-if="resolved">· Resolved</span>
        <button
          type="button"
          class="sheet-popover-open"
          data-testid="sheet-thread-open"
          @click="emit('open-thread')"
        >
          Open
        </button>
      </div>
    </div>
  </div>
</template>

<style scoped>
.sheet-note-popover-stack {
  position: absolute;
  z-index: 2;
  display: flex;
  flex-direction: column;
  gap: var(--space-1);
  pointer-events: auto;
}

.sheet-popover-card {
  box-sizing: border-box;
  max-height: 12rem;
  overflow: auto;
  padding: var(--space-2) var(--space-3);
  font-family: var(--font-ui);
  font-size: var(--text-sm);
  color: var(--text-primary);
  background: var(--surface-raised);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  box-shadow: var(--shadow-1);
}

.sheet-popover-card.is-resolved {
  opacity: 0.8;
}

.sheet-popover-author {
  font-weight: var(--weight-bold);
  margin-bottom: var(--space-1);
}

.sheet-popover-text {
  white-space: pre-wrap;
  overflow-wrap: anywhere;
}

.sheet-popover-footer {
  display: flex;
  align-items: center;
  gap: var(--space-2);
  margin-top: var(--space-2);
  font-size: var(--text-xs);
  color: var(--text-secondary);
}

.sheet-popover-open {
  margin-left: auto;
  padding: var(--space-1) var(--space-2);
  font: inherit;
  color: var(--text-primary);
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  cursor: pointer;
}

.sheet-popover-open:hover,
.sheet-popover-open:focus-visible {
  background: var(--state-hover);
}
</style>
