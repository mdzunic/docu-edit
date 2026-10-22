<script setup lang="ts">
// The data-validation alert (SPEC-130 "Entry gate"): `validation-alert`, a
// modal role="alertdialog" named by its title and described by its text
// (line breaks kept), with the style's AppIcon — `error` (stop), `warning`,
// `info` (information) — and buttons:
// - stop: `validation-alert-retry` "Retry" (focused) and
//   `validation-alert-cancel` "Cancel";
// - warning: the text, then "Continue?", and `validation-alert-yes` "Yes",
//   `validation-alert-no` "No" (focused) and `validation-alert-cancel`;
// - information: `validation-alert-ok` "OK" (focused) and
//   `validation-alert-cancel`.
// Each emits `answer` with its choice ("retry", "yes", "no", "ok",
// "cancel"); Escape is "cancel". SheetAlertDialog's model: BaseDialog's look
// in its own markup (BaseDialog's panel is a role="dialog"), and every
// keydown stops at the panel so no window or document shortcut runs behind
// it; Tab cycles its buttons. WorkbookBody renders it from
// useDataValidation's `alert`; the cell editor takes focus back on the
// answer. Runtime prop declarations per the espree constraint (see
// AppShell.vue).
import { computed, nextTick, onMounted, useTemplateRef } from "vue";
import AppIcon from "../AppIcon.vue";
import UiButton from "../ui/UiButton.vue";

const props = defineProps({
  /** "stop" | "warning" | "information". */
  alertStyle: { type: String, default: "stop" },
  /** The rule's error title, or "Invalid entry". */
  title: { type: String, required: true },
  /** The rule's error message, or the default one. */
  text: { type: String, required: true },
});

const emit = defineEmits(["answer"]);

const panel = useTemplateRef("panel");

const style = computed(() => (props.alertStyle === "warning" || props.alertStyle === "information" ? props.alertStyle : "stop"));
const icon = computed(() => (style.value === "warning" ? "warning" : style.value === "information" ? "info" : "error"));
/** The button focused on open. */
const initial = computed(() => (style.value === "warning" ? "no" : style.value === "information" ? "ok" : "retry"));
const describedBy = computed(() => (style.value === "warning" ? "validation-alert-text validation-alert-question" : "validation-alert-text"));

function buttons() {
  return panel.value ? Array.from(panel.value.querySelectorAll("button")) : [];
}

onMounted(async () => {
  await nextTick();
  const button = panel.value ? panel.value.querySelector(`[data-testid="validation-alert-${initial.value}"]`) : null;
  if (button instanceof HTMLElement) {
    button.focus();
  }
});

function onKeydown(event = new KeyboardEvent("keydown")) {
  event.stopPropagation();
  if (event.key === "Escape") {
    event.preventDefault();
    emit("answer", "cancel");
  } else if (event.key === "Tab") {
    // Focus stays among the alert's buttons.
    event.preventDefault();
    const all = buttons();
    if (all.length > 0) {
      const at = all.findIndex((button) => button === document.activeElement);
      const next = (at + (event.shiftKey ? -1 : 1) + all.length) % all.length;
      all[at < 0 ? 0 : next]?.focus();
    }
  }
}
</script>

<template>
  <div class="validation-alert-backdrop">
    <div
      ref="panel"
      class="validation-alert"
      role="alertdialog"
      aria-modal="true"
      aria-labelledby="validation-alert-title"
      :aria-describedby="describedBy"
      data-testid="validation-alert"
      :data-style="style"
      tabindex="-1"
      @keydown="onKeydown"
    >
      <div class="validation-alert-body">
        <span
          class="validation-alert-icon"
          :class="`is-${style}`"
          data-testid="validation-alert-icon"
          :data-icon="icon"
        >
          <AppIcon :name="icon" />
        </span>
        <div class="validation-alert-content">
          <h2
            id="validation-alert-title"
            class="validation-alert-title"
            data-testid="validation-alert-title"
          >
            {{ props.title }}
          </h2>
          <p
            id="validation-alert-text"
            class="validation-alert-text"
            data-testid="validation-alert-text"
          >
            {{ props.text }}
          </p>
          <p
            v-if="style === 'warning'"
            id="validation-alert-question"
            class="validation-alert-text"
          >
            Continue?
          </p>
        </div>
      </div>
      <div class="validation-alert-actions">
        <template v-if="style === 'stop'">
          <UiButton
            variant="primary"
            data-testid="validation-alert-retry"
            @click="emit('answer', 'retry')"
          >
            Retry
          </UiButton>
        </template>
        <template v-else-if="style === 'warning'">
          <UiButton
            data-testid="validation-alert-yes"
            @click="emit('answer', 'yes')"
          >
            Yes
          </UiButton>
          <UiButton
            variant="primary"
            data-testid="validation-alert-no"
            @click="emit('answer', 'no')"
          >
            No
          </UiButton>
        </template>
        <template v-else>
          <UiButton
            variant="primary"
            data-testid="validation-alert-ok"
            @click="emit('answer', 'ok')"
          >
            OK
          </UiButton>
        </template>
        <UiButton
          data-testid="validation-alert-cancel"
          @click="emit('answer', 'cancel')"
        >
          Cancel
        </UiButton>
      </div>
    </div>
  </div>
</template>

<style scoped>
.validation-alert-backdrop {
  position: fixed;
  inset: 0;
  z-index: 50;
  display: flex;
  align-items: center;
  justify-content: center;
  background: var(--surface-overlay);
}

.validation-alert {
  box-sizing: border-box;
  min-width: 20rem;
  max-width: min(32rem, 90vw);
  padding: var(--space-6);
  background: var(--surface-base);
  border-radius: var(--radius-lg);
  box-shadow: var(--shadow-3);
}

.validation-alert-body {
  display: flex;
  gap: var(--space-4);
  align-items: flex-start;
}

.validation-alert-icon {
  flex: none;
  color: var(--text-secondary);
}

.validation-alert-icon.is-stop {
  color: var(--text-danger);
}

.validation-alert-icon.is-information {
  color: var(--text-accent);
}

.validation-alert-title {
  margin: 0 0 var(--space-2);
  color: var(--text-primary);
  font-size: var(--text-md);
  font-weight: var(--weight-bold);
}

.validation-alert-text {
  margin: 0;
  color: var(--text-primary);
  font-size: var(--text-md);
  white-space: pre-line;
  overflow-wrap: anywhere;
}

.validation-alert-text + .validation-alert-text {
  margin-top: var(--space-2);
}

.validation-alert-actions {
  display: flex;
  justify-content: flex-end;
  gap: var(--space-2);
  margin-top: var(--space-6);
}
</style>
