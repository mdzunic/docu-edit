<script setup lang="ts">
// The show timer's surface (SPEC-146), in both presenting windows: the
// readout, the Start/Pause/Resume toggle and Stop, and a polite live region.
// With `editable` (presenter view only) a second row holds the Timer mode
// buttons and, in Countdown, the talk-length field. It holds no timer state:
// useShowTimer owns it, and this only renders what it is given and emits
// toggle / stop (and, editable, mode / length).
//
// `testid-prefix` names everything: "presenter" renders with UiButton
// size="sm" like the presenter's other controls; "play" renders the HUD's
// pills, whose theme-independent ink comes from PresentationPlayView
// (--hud-ink, --hud-pill, --hud-shade). Overtime is never colour alone: the
// minus sign, the "Time over" label and the colour change together. A HUD
// button keeps focus off itself on a pointer press, so the show's Space and
// Enter go on advancing. Runtime declarations, not defineProps<{...}>()
// generics (espree lints .vue scripts).
import { computed } from "vue";
import UiButton from "../ui/UiButton.vue";
import UiField from "../ui/UiField.vue";
import { parseCountdownMinutes } from "../../lib/presentation/presenterTimer";

const props = defineProps({
  /** "presenter" or "play": the prefix of every testid, and the styling. */
  testidPrefix: { type: String, default: "presenter" },
  /** Shows the mode and length controls (presenter view only). */
  editable: { type: Boolean, default: false },
  phase: { type: String, default: "stopped" },
  mode: { type: String, default: "stopwatch" },
  text: { type: String, default: "0:00" },
  overtime: { type: Boolean, default: false },
  status: { type: String, default: "" },
  lengthMinutes: { type: Number, default: 20 },
});

const emit = defineEmits(["toggle", "stop", "mode", "length"]);

const TOGGLE = {
  stopped: { label: "Start (T)", name: "Start timer" },
  running: { label: "Pause (T)", name: "Pause timer" },
  paused: { label: "Resume (T)", name: "Resume timer" },
};

const play = computed(() => props.testidPrefix === "play");
const control = computed(() => (play.value ? "button" : UiButton));
const controlAttrs = computed(() => (play.value ? { type: "button", class: "show-timer-pill" } : { size: "sm" }));
const toggle = computed(() => (props.phase === "running" ? TOGGLE.running : props.phase === "paused" ? TOGGLE.paused : TOGGLE.stopped));
const readoutLabel = computed(() => {
  if (props.overtime) return "Time over";
  return props.mode === "countdown" ? "Time remaining" : "Elapsed time";
});
const id = (suffix = "") => `${props.testidPrefix}-timer${suffix}`;

/** A HUD pill takes no focus from a pointer press (keyboard Tab still does). */
function keepFocus(event = new MouseEvent("mousedown")) {
  if (play.value) event.preventDefault();
}

/** Commits the typed length on change: rounded and clamped into 1–240; what
 * is not a number restores the current value. */
function commitLength(event = new Event("change")) {
  const input = event.target;
  if (!(input instanceof HTMLInputElement)) return;
  const minutes = parseCountdownMinutes(input.value);
  input.value = String(minutes ?? props.lengthMinutes);
  if (minutes !== null) emit("length", minutes);
}
</script>

<template>
  <div
    class="show-timer"
    :class="[`is-${testidPrefix}`, { 'is-overtime': overtime }]"
  >
    <div class="show-timer-row">
      <span
        class="show-timer-readout"
        role="timer"
        :aria-label="readoutLabel"
        :data-testid="id()"
        :data-state="phase"
        :data-mode="mode"
        :data-overtime="overtime ? 'true' : 'false'"
      >{{ text }}</span>
      <component
        :is="control"
        v-bind="controlAttrs"
        data-show-timer-control
        :data-testid="id('-toggle')"
        :aria-label="toggle.name"
        @mousedown="keepFocus"
        @click.stop="emit('toggle')"
      >
        {{ toggle.label }}
      </component>
      <component
        :is="control"
        v-bind="controlAttrs"
        data-show-timer-control
        :data-testid="id('-stop')"
        aria-label="Stop timer"
        :disabled="phase === 'stopped'"
        @mousedown="keepFocus"
        @click.stop="emit('stop')"
      >
        Stop (S)
      </component>
    </div>
    <div
      v-if="editable"
      class="show-timer-row"
    >
      <div
        class="show-timer-modes"
        role="group"
        aria-label="Timer mode"
      >
        <UiButton
          size="sm"
          :data-testid="id('-mode-stopwatch')"
          :aria-pressed="mode === 'stopwatch'"
          @click="emit('mode', 'stopwatch')"
        >
          Stopwatch
        </UiButton>
        <UiButton
          size="sm"
          :data-testid="id('-mode-countdown')"
          :aria-pressed="mode === 'countdown'"
          @click="emit('mode', 'countdown')"
        >
          Countdown
        </UiButton>
      </div>
      <UiField
        v-if="mode === 'countdown'"
        class="show-timer-length"
        label="Talk length (minutes)"
        :for="id('-length')"
      >
        <input
          :id="id('-length')"
          type="text"
          inputmode="numeric"
          class="show-timer-input"
          :data-testid="id('-length')"
          :value="String(lengthMinutes)"
          @change="commitLength"
        >
      </UiField>
    </div>
    <span
      class="sr-only"
      aria-live="polite"
      :data-testid="id('-status')"
    >{{ status }}</span>
  </div>
</template>

<style scoped>
.show-timer {
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
}

.show-timer-row,
.show-timer-modes {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: var(--space-3);
}

.show-timer-modes {
  gap: var(--space-1);
}

.show-timer-readout {
  font-variant-numeric: tabular-nums;
}

/* The presenter surface: the row SPEC-100's timer had, danger text in overtime. */
.is-presenter .show-timer-readout {
  flex: 1;
  font-size: var(--text-xl);
  font-weight: var(--weight-bold);
}

.is-presenter.is-overtime .show-timer-readout {
  color: var(--text-danger);
}

.show-timer .show-timer-length {
  flex-direction: row;
  align-items: center;
}

/* The Preferences dialog's .pref-text, sized for a few digits. */
.show-timer-input {
  width: 6ch;
  height: var(--height-control);
  padding: 0 var(--space-3);
  color: var(--text-primary);
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  font-family: var(--font-ui);
  font-size: var(--text-control);
}

.show-timer-input:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: var(--space-1);
}

/* The HUD: .hud-pill's look, its ink from PresentationPlayView. Overtime turns
   the readout's pill to the --danger family, darkened toward the HUD's black
   so white text keeps 4.5:1 in the dark schemes too (a light-pink --danger). */
.is-play .show-timer-row {
  gap: var(--space-2);
}

.is-play .show-timer-readout,
.show-timer-pill {
  border: none;
  border-radius: var(--radius-sm);
  padding: var(--space-1) var(--space-3);
  background: var(--hud-pill);
  color: var(--hud-ink);
  font: inherit;
  font-size: var(--text-md);
  font-variant-numeric: tabular-nums;
}

.is-play.is-overtime .show-timer-readout {
  background: color-mix(in srgb, var(--danger) 60%, var(--hud-shade));
}

.show-timer-pill {
  cursor: pointer;
}

.show-timer-pill:disabled {
  cursor: default;
  opacity: 0.5;
}

.show-timer-pill:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: var(--space-1);
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
