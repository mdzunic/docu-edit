<script setup lang="ts">
// One conditional-format rule's preview (SPEC-129 "The side panel"): the
// manager's `cf-rule-preview` and the rule editor's live `cf-preview` — what
// it draws per kind is views/spreadsheet/cfPreview.ts's rulePreview (sample
// text styled by the dxf, a colour-scale strip, a data bar, the icon set's
// icons). Runtime prop declarations per the espree constraint (see
// AppShell.vue).
import { computed } from "vue";
import { asCfPreview } from "../../views/spreadsheet/cfPreview";
import CfIcon from "./CfIcon.vue";

const props = defineProps({
  /** cfPreview.ts's CfPreview. */
  preview: { type: Object, required: true },
  testid: { type: String, default: "cf-rule-preview" },
});

const p = computed(() => asCfPreview(props.preview));
</script>

<template>
  <span
    class="cf-preview"
    :data-testid="testid"
    :data-kind="p.kind"
    aria-hidden="true"
  >
    <span
      v-if="p.kind === 'format'"
      class="cf-preview-format"
      :style="p.box"
    ><span :style="p.text">{{ p.sample }}</span></span>
    <span
      v-else-if="p.kind === 'scale'"
      class="cf-preview-strip"
      :style="{ background: p.background }"
    />
    <span
      v-else-if="p.kind === 'bar'"
      class="cf-preview-bar"
    ><span
      class="cf-preview-bar-fill"
      :style="{ background: p.background }"
    /></span>
    <span
      v-else-if="p.kind === 'icons'"
      class="cf-preview-icons"
    >
      <CfIcon
        v-for="index in p.indices"
        :key="index"
        :set="p.set"
        :index="index"
        :size-px="14"
      />
    </span>
    <span
      v-else
      class="cf-preview-none"
    >—</span>
  </span>
</template>

<style scoped>
.cf-preview {
  display: inline-flex;
  align-items: center;
  flex: 0 0 auto;
  width: 4.5rem;
  height: 1.5rem;
  border: 1px solid var(--color-border, #d0d0d0);
  background: #ffffff;
  overflow: hidden;
}
.cf-preview-format {
  display: flex;
  align-items: center;
  justify-content: center;
  width: 100%;
  height: 100%;
  box-sizing: border-box;
  font-size: 0.75rem;
  color: #000000;
}
.cf-preview-strip {
  width: 100%;
  height: 100%;
}
.cf-preview-bar {
  display: flex;
  width: 100%;
  height: 100%;
  padding: 3px 0;
  box-sizing: border-box;
}
.cf-preview-bar-fill {
  width: 70%;
  height: 100%;
}
.cf-preview-icons {
  display: inline-flex;
  gap: 1px;
  padding: 0 2px;
}
.cf-preview-none {
  width: 100%;
  text-align: center;
  color: var(--color-text-muted, #666666);
}
</style>
