<script setup lang="ts">
// The stage column (SPEC-096 extraction, paying for this spec's hookups): the
// fit-scaled stage area and the below-stage panel stack — the `.stage-column`
// block that lived inline — plus the SPEC-096 chrome that hangs off the
// stage: the one-line arrange notice (aria-live), the confirm-to-release
// alertdialog for a rotated verbatim group, and the Position & Size dialog.
// A dumb component: every controller it needs arrives through `stageEdit`.
// Runtime prop declarations per the espree constraint (see AppShell.vue).
import { computed } from "vue";
import SlideStageArea from "./SlideStageArea.vue";
import SlidePanelStack from "./SlidePanelStack.vue";
import PositionSizeDialog from "./PositionSizeDialog.vue";
import UiButton from "../ui/UiButton.vue";

const props = defineProps({
  currentSlide: { type: Object, default: null },
  widthCm: { type: Number, required: true },
  heightCm: { type: Number, required: true },
  media: { type: Array, default: () => [] },
  theme: { type: Object, default: null },
  liveLabel: { type: String, default: "" },
  stageEdit: { type: Object, required: true },
  animations: { type: Object, required: true },
  transitions: { type: Object, required: true },
  notes: { type: Object, required: true },
  tables: { type: Object, required: true },
  mediaPanel: { type: Object, required: true },
  currentIndex: { type: Number, default: 0 },
  fields: { type: Array, default: () => [] },
});

defineEmits(["add"]);

const s = props.stageEdit;
const arrange = computed(() => s.arrange);
</script>

<template>
  <div class="stage-column">
    <SlideStageArea
      :current-slide="currentSlide ?? undefined"
      :width-cm="widthCm"
      :height-cm="heightCm"
      :media="media"
      :live-label="liveLabel"
      :stage-edit="stageEdit"
      :theme="theme"
      :animations="animations"
      :transitions="transitions"
      :fields="fields"
      @add="$emit('add')"
    />
    <div
      v-if="arrange.notice.value"
      class="arrange-notice"
      aria-live="polite"
      data-testid="arrange-notice"
    >
      {{ arrange.notice.value }}
    </div>
    <SlidePanelStack
      :notes="notes"
      :transitions="transitions"
      :animations="animations"
      :tables="tables"
      :media="mediaPanel"
      :slide-number="currentIndex + 1"
      :has-slide="!!currentSlide"
    />
    <!-- Releasing a rotated verbatim group is lossy — confirm (SPEC-096) -->
    <div
      v-if="arrange.pendingRelease.value"
      class="release-confirm"
      role="alertdialog"
      aria-label="Replace the original grouping"
      data-testid="group-release-confirm-box"
    >
      <p>Ungrouping discards this group's original rotation. Replace the original grouping?</p>
      <div class="confirm-actions">
        <UiButton
          variant="secondary"
          data-testid="group-release-cancel"
          @click="arrange.cancelRelease()"
        >
          Cancel
        </UiButton>
        <UiButton
          variant="primary"
          data-testid="group-release-confirm"
          @click="arrange.confirmRelease()"
        >
          Ungroup
        </UiButton>
      </div>
    </div>
    <PositionSizeDialog
      v-if="arrange.positionOpen.value && s.selection.anchorFrame.value"
      :frame="s.selection.anchorFrame.value"
      :unit="arrange.unit.value"
      @apply="arrange.applyPosition($event)"
      @dismiss="arrange.closePosition()"
    />
  </div>
</template>

<style scoped>
.stage-column {
  min-width: 0;
  min-height: 0;
  display: flex;
  flex-direction: column;
  position: relative;
}

.arrange-notice {
  padding: var(--space-1) var(--space-3);
  font-size: var(--text-sm);
  color: var(--text-secondary);
  background: var(--surface-raised);
  border-top: var(--border-width) solid var(--border-subtle);
}

.release-confirm {
  position: absolute;
  top: var(--space-5);
  left: 50%;
  transform: translateX(-50%);
  z-index: 40;
  padding: var(--space-4);
  background: var(--surface-raised);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
  box-shadow: var(--shadow-2);
  max-width: 26rem;
}

.release-confirm p {
  margin: 0 0 var(--space-3);
}

.confirm-actions {
  display: flex;
  justify-content: flex-end;
  gap: var(--space-2);
}
</style>
