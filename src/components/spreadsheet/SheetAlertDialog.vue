<script setup lang="ts">
// The workbook's refusal alert (SPEC-128 "UI"): `sheet-alert`, a modal
// role="alertdialog" named and described by its message, with one OK
// (`sheet-alert-ok`, focused on open). A command that cannot run — a sort
// over merged cells, a filter over an empty region — says why here and
// changes nothing; later specs (SPEC-139) reuse it for their refusals.
// BaseDialog's look, but its own markup: BaseDialog's panel is a
// role="dialog". Escape, OK and a press on the backdrop all close it (emit
// `close`); the caller returns focus to the grid. Runtime prop declarations
// per the espree constraint (see AppShell.vue).
import { nextTick, onMounted, useTemplateRef } from "vue";
import UiButton from "../ui/UiButton.vue";

const props = defineProps({
  /** What could not be done, and why. */
  message: { type: String, required: true },
});

const emit = defineEmits(["close"]);

const panel = useTemplateRef("panel");

onMounted(async () => {
  await nextTick();
  const button = panel.value ? panel.value.querySelector('[data-testid="sheet-alert-ok"]') : null;
  if (button instanceof HTMLElement) {
    button.focus();
  }
});

function onKeydown(event = new KeyboardEvent("keydown")) {
  if (event.key === "Escape") {
    event.stopPropagation();
    emit("close");
  } else if (event.key === "Tab") {
    // One control: focus stays on OK.
    event.preventDefault();
  }
}
</script>

<template>
  <div
    class="sheet-alert-backdrop"
    @click.self="emit('close')"
  >
    <div
      ref="panel"
      class="sheet-alert"
      role="alertdialog"
      aria-modal="true"
      aria-labelledby="sheet-alert-message"
      data-testid="sheet-alert"
      tabindex="-1"
      @keydown="onKeydown"
    >
      <p
        id="sheet-alert-message"
        class="sheet-alert-message"
        data-testid="sheet-alert-message"
      >
        {{ props.message }}
      </p>
      <div class="sheet-alert-actions">
        <UiButton
          variant="primary"
          data-testid="sheet-alert-ok"
          @click="emit('close')"
        >
          OK
        </UiButton>
      </div>
    </div>
  </div>
</template>

<style scoped>
.sheet-alert-backdrop {
  position: fixed;
  inset: 0;
  z-index: 50;
  display: flex;
  align-items: center;
  justify-content: center;
  background: var(--surface-overlay);
}

.sheet-alert {
  box-sizing: border-box;
  min-width: 20rem;
  max-width: min(32rem, 90vw);
  padding: var(--space-6);
  background: var(--surface-base);
  border-radius: var(--radius-lg);
  box-shadow: var(--shadow-3);
}

.sheet-alert-message {
  margin: 0;
  color: var(--text-primary);
  font-size: var(--text-md);
}

.sheet-alert-actions {
  display: flex;
  justify-content: flex-end;
  margin-top: var(--space-6);
}
</style>
