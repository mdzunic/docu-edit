<script setup lang="ts">
// The workbook's app-bar cluster (SPEC-116 on SPEC-058's app bar): the
// format name ("Excel workbook (.xlsx)", …), Download — the file that came
// in, part for part — and SPEC-060's Close. AppShell's #app-bar-actions slot
// is where every editor puts its cluster; the title and the format badge are
// AppShell's own props. Runtime declarations per the espree constraint.
import UiButton from "../ui/UiButton.vue";
import UiIconButton from "../ui/UiIconButton.vue";

const props = defineProps({
  /** e.g. "Excel workbook (.xlsx)"; empty for a missing workbook. */
  formatLabel: { type: String, default: "" },
});

const emit = defineEmits(["download", "close"]);
</script>

<template>
  <div class="workbook-actions">
    <template v-if="props.formatLabel">
      <span
        class="workbook-format"
        data-testid="spreadsheet-format"
      >{{ props.formatLabel }}</span>
      <UiButton
        data-testid="spreadsheet-download"
        @click="emit('download')"
      >
        Download
      </UiButton>
    </template>
    <UiIconButton
      class="close-button"
      icon="close"
      label="Close document"
      data-testid="close-document"
      @click="emit('close')"
    />
  </div>
</template>

<style scoped>
.workbook-actions {
  display: flex;
  flex: 0 1 auto;
  min-width: 0;
  gap: var(--space-3);
  align-items: center;
}

/* The format name in the caption face, as the deck's facts are — chrome to
   glance at; it is what shrinks first on a crowded bar. */
.workbook-format {
  flex: 0 1 auto;
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  color: var(--text-secondary);
  font-family: var(--font-caption);
  font-size: var(--text-sm);
  letter-spacing: var(--tracking-caption);
  text-transform: uppercase;
  white-space: nowrap;
}

/* On the brand bar the icon button takes the on-brand foreground and a
   translucent wash, the treatment every editor's Close gets. */
.workbook-actions .close-button {
  color: var(--text-on-brand);
}

.workbook-actions .close-button:not(:disabled):hover {
  color: var(--text-on-brand);
  background: color-mix(in srgb, var(--text-on-brand) 12%, transparent);
}

.workbook-actions .close-button:not(:disabled):active {
  background: color-mix(in srgb, var(--text-on-brand) 20%, transparent);
}
</style>
