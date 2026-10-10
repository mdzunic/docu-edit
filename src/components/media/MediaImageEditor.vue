<script setup lang="ts">
// The image editor's surface (SPEC-112, `media-edit`): the action bar, the
// notices, the tool rail, the active tool's options bar, the preview stage
// and the history list, over the session in views/media/useImageEditor
// (handed in as the `editor` prop — the controller-prop pattern the viewer
// uses). Every control is native and keyboard-operable; the logic is the
// composables', this file only lays it out.
// Runtime prop declarations — see AppShell.vue.
import { computed, nextTick, useTemplateRef, watch } from "vue";
import { asImageEditor } from "../../views/media/useImageEditor";
import UiButton from "../ui/UiButton.vue";
import EditActionBar from "./EditActionBar.vue";
import EditToolRail from "./EditToolRail.vue";
import EditCropOptions from "./EditCropOptions.vue";
import EditTransformOptions from "./EditTransformOptions.vue";
import EditAdjustOptions from "./EditAdjustOptions.vue";
import EditObjectOptions from "./EditObjectOptions.vue";
import EditStage from "./EditStage.vue";
import EditHistoryList from "./EditHistoryList.vue";
import EditNotices from "./EditNotices.vue";
import EditDialogs from "./EditDialogs.vue";

const props = defineProps({
  /** The useImageEditor controller. */
  editor: { type: Object, required: true },
});

const e = computed(() => asImageEditor(props.editor));
const ready = computed(() => e.value.phase.value === "ready");

// The Edit action that opened the editor is gone from the page, so focus
// would fall to <body>: a keyboard user lands on the active tool instead.
const root = useTemplateRef("root");
watch(
  ready,
  async (isReady) => {
    await nextTick();
    const tool = isReady ? root.value?.querySelector('[role="toolbar"] [aria-pressed="true"]') : null;
    if (tool instanceof HTMLElement) {
      tool.focus();
    }
  },
  { immediate: true },
);
</script>

<template>
  <section
    ref="root"
    class="image-editor"
    data-testid="media-edit"
    aria-label="Edit image"
  >
    <EditActionBar :editor="e" />
    <EditNotices :editor="e" />
    <div
      v-if="e.phase.value === 'size'"
      class="svg-size"
      data-testid="edit-svg-size"
    >
      <p>This SVG does not say how large it is. Choose the size to edit it at:</p>
      <label>
        Width (px)
        <input
          v-model.number="e.svgSize.value.widthPx"
          type="number"
          min="1"
          max="16384"
          data-testid="edit-svg-width"
        >
      </label>
      <label>
        Height (px)
        <input
          v-model.number="e.svgSize.value.heightPx"
          type="number"
          min="1"
          max="16384"
          data-testid="edit-svg-height"
        >
      </label>
      <UiButton
        variant="primary"
        data-testid="edit-svg-start"
        @click="e.startSvg()"
      >
        Start editing
      </UiButton>
    </div>
    <p
      v-else-if="e.phase.value === 'loading'"
      class="loading"
      role="status"
    >
      Preparing the image…
    </p>
    <div
      v-if="ready"
      class="edit-body"
    >
      <EditToolRail :editor="e" />
      <div class="edit-main">
        <EditCropOptions
          v-if="e.tool.value === 'crop'"
          :editor="e"
        />
        <EditTransformOptions
          v-else-if="e.tool.value === 'rotate'"
          :editor="e"
        />
        <EditAdjustOptions
          v-else-if="e.tool.value === 'adjust'"
          :editor="e"
        />
        <EditObjectOptions
          v-else
          :editor="e"
          :mode="e.tool.value"
        />
        <EditStage :editor="e" />
      </div>
      <EditHistoryList :editor="e" />
    </div>
    <EditDialogs :editor="e" />
  </section>
</template>

<style scoped>
.image-editor {
  flex: 1 1 auto;
  min-width: 0;
  min-height: 0;
  display: flex;
  flex-direction: column;
  background: var(--surface-sunken);
}

.edit-body {
  flex: 1 1 auto;
  min-height: 0;
  display: flex;
}

.edit-main {
  flex: 1 1 auto;
  min-width: 0;
  display: flex;
  flex-direction: column;
}

.svg-size {
  display: flex;
  flex-wrap: wrap;
  align-items: flex-end;
  gap: var(--space-3);
  padding: var(--space-4);
  color: var(--text-primary);
}

.svg-size p {
  flex-basis: 100%;
  margin: 0;
}

.svg-size label {
  display: flex;
  flex-direction: column;
  gap: var(--space-1);
  font-size: var(--text-sm);
}

.svg-size input {
  width: 7rem;
  font: inherit;
}

.loading {
  padding: var(--space-4);
  color: var(--text-secondary);
}
</style>
