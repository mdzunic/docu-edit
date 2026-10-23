<script setup lang="ts">
// The animations side panel (SPEC-086): an ordered effect list (each row: order
// number, target frame name, effect, trigger glyph), the selected row's field
// controls (effect, direction, trigger, duration, delay), Add/Remove/Move
// up-down/Play buttons, and the inline confirm for replacing a slide's
// preserved out-of-subset timing. All logic lives in the `useAnimations`
// controller passed in; this component is presentation + event forwarding only.
// Every control is keyboard-operable (native buttons/select/number inputs) and
// carries the spec's testids. Runtime prop declarations, not
// defineProps<{...}>() generics (espree lints .vue scripts).
//
// SPEC-105: the effect list groups into Entrance / Emphasis / Exit / Motion
// path; the build, build-trigger, repeat, until-click, auto-reverse, rewind,
// trigger-frame, spin, scale, colour and motion-path controls each show only
// where the selected row uses them; and the timeline toggle reveals the
// AnimationTimeline view of the same list.
import UiButton from "../ui/UiButton.vue";
import AnimationTimeline from "./AnimationTimeline.vue";

const props = defineProps({
  /** The useAnimations() controller object (refs + command functions). */
  controller: { type: Object, required: true },
  slideNumber: { type: Number, default: 1 },
});

const EFFECT_GROUPS = [
  {
    label: "Entrance",
    effects: [
      { value: "appear", label: "Appear" },
      { value: "fadeIn", label: "Fade in" },
      { value: "flyIn", label: "Fly in" },
      { value: "wipeIn", label: "Wipe in" },
      { value: "zoomIn", label: "Zoom in" },
    ],
  },
  {
    label: "Emphasis",
    effects: [
      { value: "pulse", label: "Pulse" },
      { value: "spin", label: "Spin" },
      { value: "growShrink", label: "Grow/Shrink" },
      { value: "colorChange", label: "Colour change" },
    ],
  },
  {
    label: "Exit",
    effects: [
      { value: "fadeOut", label: "Fade out" },
      { value: "flyOut", label: "Fly out" },
    ],
  },
  { label: "Motion path", effects: [{ value: "motionPath", label: "Motion path" }] },
];
const EFFECTS = EFFECT_GROUPS.flatMap((g) => g.effects);
const DIRECTIONS = [
  { value: "left", label: "Left" },
  { value: "right", label: "Right" },
  { value: "up", label: "Up" },
  { value: "down", label: "Down" },
];
const TRIGGERS = [
  { value: "click", label: "On click" },
  { value: "withPrevious", label: "With previous" },
  { value: "afterPrevious", label: "After previous" },
  { value: "onClickOf", label: "On click of" },
];
const BUILD_MODES = [
  { value: "whole", label: "All at once" },
  { value: "byParagraph", label: "By paragraph" },
  { value: "byFirstLevelParagraph", label: "By first-level paragraph" },
];
const BUILD_TRIGGERS = TRIGGERS.slice(0, 3);
const PATH_PRESETS = [
  { value: "line", label: "Line" },
  { value: "arc", label: "Arc" },
  { value: "turn", label: "Turn" },
  { value: "loop", label: "Loop" },
  { value: "custom", label: "Custom" },
];
const c = props.controller;

function effectLabel(value = "") {
  return EFFECTS.find((e) => e.value === value)?.label ?? value;
}
/** The row's build, when it has one. */
function buildLabel(value = "") {
  return value ? BUILD_MODES.find((m) => m.value === value)?.label ?? "" : "";
}
/** A compact glyph per trigger for the row summary (AC-6). */
function triggerGlyph(value = "") {
  if (value === "withPrevious") return "⇊";
  if (value === "afterPrevious") return "↳";
  if (value === "onClickOf") return "◎";
  return "▷";
}
function triggerTitle(row = { trigger: "", triggerFrameName: "" }) {
  return row.trigger === "onClickOf" ? `On click of ${row.triggerFrameName}` : row.trigger;
}
function triggerOptionDisabled(value = "") {
  if (value === "onClickOf") return c.triggerFrameOptions.value.length === 0;
  return c.triggerLocked.value && value !== "click";
}
function selectValue(event = new Event("change")) {
  const el = event.target;
  return el instanceof HTMLSelectElement || el instanceof HTMLInputElement ? el.value : "";
}
function checked(event = new Event("change")) {
  const el = event.target;
  return el instanceof HTMLInputElement && el.checked;
}
function onEffect(event = new Event("change")) {
  const el = event.target;
  if (el instanceof HTMLSelectElement) c.setEffect(el.value);
}
function onDirection(event = new Event("change")) {
  const el = event.target;
  if (el instanceof HTMLSelectElement) c.setDirection(el.value);
}
function onTrigger(event = new Event("change")) {
  const el = event.target;
  if (el instanceof HTMLSelectElement) c.setTrigger(el.value);
}
function onDuration(event = new Event("change")) {
  const el = event.target;
  if (el instanceof HTMLInputElement) c.setDurationSeconds(Number(el.value));
}
function onDelay(event = new Event("change")) {
  const el = event.target;
  if (el instanceof HTMLInputElement) c.setDelaySeconds(Number(el.value));
}
</script>

<template>
  <section
    class="animations-panel"
    data-testid="slides-animation-panel"
    :aria-label="`Animations for slide ${slideNumber}`"
  >
    <ul
      class="effect-list"
      data-testid="animation-list"
      role="listbox"
      :aria-label="`Animations for slide ${slideNumber}`"
    >
      <li
        v-for="row in c.rows.value"
        :key="row.index"
      >
        <button
          type="button"
          class="row"
          role="option"
          :class="{ 'is-selected': row.index === c.selectedIndex.value }"
          :aria-selected="row.index === c.selectedIndex.value"
          data-testid="animation-row"
          @click="c.selectRow(row.index)"
        >
          <span class="row-num">{{ row.number }}</span>
          <span class="row-name">{{ row.frameName }}</span>
          <span class="row-effect">{{ effectLabel(row.effect) }}</span>
          <span
            v-if="row.buildBy"
            class="row-build"
          >{{ buildLabel(row.buildBy) }}</span>
          <span
            class="row-trigger"
            :title="triggerTitle(row)"
            aria-hidden="true"
          >{{ triggerGlyph(row.trigger) }}</span>
        </button>
      </li>
      <li
        v-if="c.rows.value.length === 0"
        class="empty"
        data-testid="animation-list-empty"
      >
        No animations yet.
      </li>
    </ul>

    <div class="toolbar-row">
      <UiButton
        variant="secondary"
        data-testid="animation-add"
        :disabled="!c.canAdd.value"
        @click="c.add()"
      >
        Add
      </UiButton>
      <UiButton
        variant="ghost"
        data-testid="animation-remove"
        :disabled="!c.hasSelection.value"
        @click="c.remove()"
      >
        Remove
      </UiButton>
      <UiButton
        variant="ghost"
        data-testid="animation-move-up"
        :disabled="!c.canMoveUp.value"
        @click="c.moveUp()"
      >
        Move up
      </UiButton>
      <UiButton
        variant="ghost"
        data-testid="animation-move-down"
        :disabled="!c.canMoveDown.value"
        @click="c.moveDown()"
      >
        Move down
      </UiButton>
      <UiButton
        variant="ghost"
        data-testid="animation-play"
        :disabled="c.rows.value.length === 0"
        @click="c.play()"
      >
        Play
      </UiButton>
      <UiButton
        variant="ghost"
        data-testid="animation-timeline-toggle"
        :aria-pressed="c.timelineOpen.value ? 'true' : 'false'"
        @click="c.toggleTimeline()"
      >
        Timeline
      </UiButton>
    </div>

    <AnimationTimeline
      v-if="c.timelineOpen.value"
      :controller="c"
    />

    <p
      v-if="!c.canAdd.value && c.addHelp.value"
      class="helper"
      data-testid="animation-add-help"
    >
      {{ c.addHelp.value }}
    </p>

    <div
      v-if="c.hasSelection.value"
      class="controls"
    >
      <label class="field">
        <span>Effect</span>
        <select
          data-testid="animation-effect"
          :value="c.selectedEffect.value"
          @change="onEffect"
        >
          <optgroup
            v-for="g in EFFECT_GROUPS"
            :key="g.label"
            :label="g.label"
          >
            <option
              v-for="e in g.effects"
              :key="e.value"
              :value="e.value"
            >
              {{ e.label }}
            </option>
          </optgroup>
        </select>
      </label>

      <label class="field">
        <span>Direction</span>
        <select
          data-testid="animation-direction"
          :disabled="!c.directionEnabled.value"
          :value="c.selectedDirection.value"
          @change="onDirection"
        >
          <option
            v-for="d in DIRECTIONS"
            :key="d.value"
            :value="d.value"
          >
            {{ d.label }}
          </option>
        </select>
      </label>

      <label class="field">
        <span>Trigger</span>
        <select
          data-testid="animation-trigger"
          :value="c.selectedTrigger.value"
          @change="onTrigger"
        >
          <option
            v-for="t in TRIGGERS"
            :key="t.value"
            :value="t.value"
            :disabled="triggerOptionDisabled(t.value)"
          >
            {{ t.label }}
          </option>
        </select>
      </label>

      <label class="field">
        <span>Trigger shape</span>
        <select
          data-testid="animation-trigger-frame"
          :disabled="!c.triggerFrameEnabled.value"
          :value="c.selectedTriggerFrameId.value"
          @change="c.setTriggerFrame(selectValue($event))"
        >
          <option
            v-if="!c.triggerFrameEnabled.value"
            value=""
          >
            —
          </option>
          <option
            v-for="f in c.triggerFrameOptions.value"
            :key="f.value"
            :value="f.value"
          >
            {{ f.label }}
          </option>
        </select>
      </label>

      <label class="field">
        <span>Build</span>
        <select
          data-testid="animation-build-by"
          :disabled="!c.buildByEnabled.value"
          :aria-describedby="c.buildByHelp.value ? 'animation-build-by-help' : undefined"
          :value="c.selectedBuildBy.value"
          @change="c.setBuildBy(selectValue($event))"
        >
          <option
            v-for="m in BUILD_MODES"
            :key="m.value"
            :value="m.value"
          >
            {{ m.label }}
          </option>
        </select>
      </label>
      <p
        v-if="c.buildByHelp.value"
        id="animation-build-by-help"
        class="helper"
        data-testid="animation-build-by-help"
      >
        {{ c.buildByHelp.value }}
      </p>

      <label
        v-if="c.buildTriggerShown.value"
        class="field"
      >
        <span>Each paragraph</span>
        <select
          data-testid="animation-build-trigger"
          :value="c.selectedBuildTrigger.value"
          @change="c.setBuildTrigger(selectValue($event))"
        >
          <option
            v-for="t in BUILD_TRIGGERS"
            :key="t.value"
            :value="t.value"
          >
            {{ t.label }}
          </option>
        </select>
      </label>

      <label class="field">
        <span>Duration (seconds)</span>
        <input
          type="number"
          data-testid="animation-duration"
          min="0.1"
          max="5"
          step="0.1"
          :value="c.durationSeconds.value"
          @change="onDuration"
        >
      </label>

      <label class="field">
        <span>Delay (seconds)</span>
        <input
          type="number"
          data-testid="animation-delay"
          min="0"
          max="10"
          step="0.1"
          :value="c.delaySeconds.value"
          @change="onDelay"
        >
      </label>

      <template v-if="c.repeatShown.value">
        <label class="field">
          <span>Repeat (times)</span>
          <input
            type="number"
            data-testid="animation-repeat"
            min="1"
            max="10"
            step="1"
            :disabled="c.selectedRepeatUntilClick.value"
            :value="c.selectedRepeatCount.value"
            @change="c.setRepeatCount(Number(selectValue($event)))"
          >
        </label>
        <label class="check">
          <input
            type="checkbox"
            data-testid="animation-repeat-until-click"
            :checked="c.selectedRepeatUntilClick.value"
            @change="c.setRepeatUntilClick(checked($event))"
          >
          <span>Repeat until next click</span>
        </label>
        <label class="check">
          <input
            type="checkbox"
            data-testid="animation-auto-reverse"
            :checked="c.selectedAutoReverse.value"
            @change="c.setAutoReverse(checked($event))"
          >
          <span>Auto-reverse</span>
        </label>
      </template>

      <label
        v-if="c.rewindShown.value"
        class="check"
      >
        <input
          type="checkbox"
          data-testid="animation-rewind"
          :checked="c.selectedRewind.value"
          @change="c.setRewind(checked($event))"
        >
        <span>Rewind when done</span>
      </label>

      <label
        v-if="c.selectedEffect.value === 'spin'"
        class="field"
      >
        <span>Spin (degrees)</span>
        <input
          type="number"
          data-testid="animation-spin-degrees"
          min="-3600"
          max="3600"
          step="15"
          :value="c.selectedSpinDegrees.value"
          @change="c.setSpinDegrees(Number(selectValue($event)))"
        >
      </label>

      <label
        v-if="c.selectedEffect.value === 'growShrink'"
        class="field"
      >
        <span>Scale (%)</span>
        <input
          type="number"
          data-testid="animation-scale-percent"
          min="10"
          max="400"
          step="10"
          :value="c.selectedScalePercent.value"
          @change="c.setScalePercent(Number(selectValue($event)))"
        >
      </label>

      <label
        v-if="c.selectedEffect.value === 'colorChange'"
        class="field"
      >
        <span>Colour</span>
        <input
          type="color"
          data-testid="animation-emphasis-color"
          :value="c.selectedEmphasisColor.value"
          @change="c.setEmphasisColor(selectValue($event))"
        >
      </label>

      <template v-if="c.selectedEffect.value === 'motionPath'">
        <label class="field">
          <span>Path</span>
          <select
            data-testid="animation-path-preset"
            :value="c.selectedPathPreset.value"
            @change="c.setPathPreset(selectValue($event))"
          >
            <option
              v-for="p in PATH_PRESETS"
              :key="p.value"
              :value="p.value"
              :disabled="p.value === 'custom' && c.selectedPathPreset.value !== 'custom'"
            >
              {{ p.label }}
            </option>
          </select>
        </label>
        <label
          v-if="c.pathDirectionShown.value"
          class="field"
        >
          <span>Path direction</span>
          <select
            data-testid="animation-path-direction"
            :value="c.selectedPathDirection.value"
            @change="c.setPathDirection(selectValue($event))"
          >
            <option
              v-for="d in DIRECTIONS"
              :key="d.value"
              :value="d.value"
            >
              {{ d.label }}
            </option>
          </select>
        </label>
        <label
          v-if="c.selectedPathPreset.value !== 'custom'"
          class="field"
        >
          <span>Distance (% of slide)</span>
          <input
            type="number"
            data-testid="animation-path-distance"
            min="5"
            max="100"
            step="1"
            :value="c.selectedPathDistancePercent.value"
            @change="c.setPathDistancePercent(Number(selectValue($event)))"
          >
        </label>
        <p
          v-if="c.selectedPathPreset.value === 'custom'"
          class="helper"
          data-testid="animation-path-custom-note"
        >
          An imported custom path — shown as it is and kept on save. Choose another path to replace it.
        </p>
        <UiButton
          variant="secondary"
          data-testid="animation-path-edit"
          :aria-pressed="c.pathEdit.armed.value ? 'true' : 'false'"
          @click="c.pathEdit.toggle()"
        >
          {{ c.pathEdit.armed.value ? "Done editing path" : "Edit path on slide" }}
        </UiButton>
      </template>
    </div>

    <div
      v-if="c.pendingReplaceFrameId.value"
      class="confirm"
      role="alertdialog"
      aria-label="Replace preserved animations"
      data-testid="animation-replace-confirm"
    >
      <p>Replace the original animations this editor can't show?</p>
      <div class="confirm-actions">
        <UiButton
          variant="secondary"
          data-testid="animation-replace-cancel"
          @click="c.cancelReplace()"
        >
          Cancel
        </UiButton>
        <UiButton
          variant="primary"
          data-testid="animation-replace-confirm-button"
          @click="c.confirmReplace()"
        >
          Replace
        </UiButton>
      </div>
    </div>
  </section>
</template>

<style scoped>
.animations-panel {
  flex: 0 0 auto;
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
  padding: var(--space-4);
  border-top: var(--border-width) solid var(--border-subtle);
  background: var(--surface-default);
  color: var(--text-primary);
  font-size: var(--text-sm);
  max-height: 45%;
  overflow-y: auto;
}

.effect-list {
  list-style: none;
  margin: 0;
  padding: 0;
  border: var(--border-width) solid var(--border-subtle);
  border-radius: var(--radius-sm);
  max-height: 10rem;
  overflow-y: auto;
}

.row {
  display: flex;
  align-items: center;
  gap: var(--space-2);
  width: 100%;
  padding: var(--space-1) var(--space-2);
  border: none;
  background: transparent;
  color: inherit;
  text-align: left;
  cursor: pointer;
  font-size: inherit;
}

.row:hover {
  background: var(--surface-sunken);
}

.row.is-selected {
  background: var(--accent-soft, var(--surface-sunken));
}

.row-num {
  min-width: 1.25rem;
  font-variant-numeric: tabular-nums;
  color: var(--text-secondary);
}

.row-name {
  flex: 1 1 auto;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.row-effect {
  color: var(--text-secondary);
}

.row-trigger {
  min-width: 1rem;
  text-align: center;
}

.row-build {
  color: var(--text-secondary);
  font-size: var(--text-xs);
}

.empty {
  padding: var(--space-2);
  color: var(--text-secondary);
  font-style: italic;
}

.toolbar-row {
  display: flex;
  flex-wrap: wrap;
  gap: var(--space-2);
}

.controls {
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
}

.field {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--space-3);
}

.field select,
.field input {
  min-width: 7rem;
}

.check {
  display: flex;
  align-items: center;
  gap: var(--space-2);
}

.helper {
  margin: 0;
  color: var(--text-secondary);
}

.confirm {
  border: var(--border-width) solid var(--border-strong, var(--border-subtle));
  border-radius: var(--radius-sm);
  padding: var(--space-3);
  background: var(--surface-sunken);
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
}

.confirm p {
  margin: 0;
}

.confirm-actions {
  display: flex;
  gap: var(--space-3);
  justify-content: flex-end;
}
</style>
