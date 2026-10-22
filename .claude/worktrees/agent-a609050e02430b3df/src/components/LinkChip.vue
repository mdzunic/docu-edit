<script setup lang="ts">
// The link chip (SPEC-025 AC-8..13): floats under a caret sitting inside a
// hyperlink, shows the target, and offers Open / Edit / Copy link / Remove
// link. Emits intents only — the editor owns navigation, the dialog and the
// remove command. @mousedown.prevent keeps the editor caret where it is.
// Runtime prop declarations — see AppShell.vue's espree comment.
const props = defineProps({
  /** Viewport coordinates the chip is pinned at (below the caret). */
  left: { type: Number, required: true },
  top: { type: Number, required: true },
  /** The displayed target: the URL, a heading's text or a bookmark name. */
  label: { type: String, required: true },
  /** External targets open in a new tab; internal ones jump. */
  external: { type: Boolean, default: false },
});

const emit = defineEmits(["open", "edit", "copy", "remove", "dismiss"]);

function chipStyle() {
  return { left: `${props.left}px`, top: `${props.top}px` };
}
</script>

<template>
  <div
    class="link-chip"
    data-testid="link-chip"
    role="toolbar"
    aria-label="Link options"
    :style="chipStyle()"
    @keydown.escape.stop="emit('dismiss')"
  >
    <button
      type="button"
      class="link-chip-target"
      data-testid="link-chip-open"
      :title="external ? 'Open link in new tab' : 'Go to target'"
      @mousedown.prevent
      @click="emit('open')"
    >
      <span
        class="link-chip-label"
        data-testid="link-chip-target"
      >{{ label }}</span>
    </button>
    <button
      type="button"
      class="link-chip-action"
      data-testid="link-chip-edit"
      @mousedown.prevent
      @click="emit('edit')"
    >
      Edit
    </button>
    <button
      type="button"
      class="link-chip-action"
      data-testid="link-chip-copy"
      @mousedown.prevent
      @click="emit('copy')"
    >
      Copy link
    </button>
    <button
      type="button"
      class="link-chip-action"
      data-testid="link-chip-remove"
      @mousedown.prevent
      @click="emit('remove')"
    >
      Remove link
    </button>
  </div>
</template>

<style scoped>
.link-chip {
  position: fixed;
  z-index: 30;
  display: flex;
  align-items: center;
  gap: 0.25rem;
  padding: 0.25rem;
  background: #fff;
  border: 1px solid #dadce0;
  border-radius: 6px;
  box-shadow: 0 2px 8px rgb(0 0 0 / 0.25);
  max-width: 32rem;
}

.link-chip-target {
  font: inherit;
  font-size: 0.8125rem;
  padding: 0.25rem 0.625rem;
  border: 1px solid transparent;
  border-radius: 4px;
  background: transparent;
  color: #1a73e8;
  cursor: pointer;
  overflow: hidden;
  max-width: 16rem;
}

.link-chip-label {
  display: block;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  text-decoration: underline;
}

.link-chip-target:hover {
  background: #f1f3f4;
}

.link-chip-action {
  font: inherit;
  font-size: 0.8125rem;
  padding: 0.25rem 0.625rem;
  border: 1px solid transparent;
  border-radius: 4px;
  background: transparent;
  color: #202124;
  cursor: pointer;
  white-space: nowrap;
}

.link-chip-action:hover {
  background: #f1f3f4;
}
</style>
