<script setup lang="ts">
// The workbook's split Download (SPEC-121), PresentationDownloadButton's
// twin: the primary "Download" writes the workbook in its own format; the ▾
// menu offers "Download a copy as …" the other format. When the copy cannot
// be made yet the item stays, disabled, with the reason as its title and
// under it. The popup is a UiMenuSurface, so Escape and an outside click
// close it and focus returns to the caret. Runtime prop declarations per the
// espree constraint (see AppShell.vue).
import { nextTick, ref, useTemplateRef } from "vue";
import UiButton from "../ui/UiButton.vue";
import UiMenuSurface from "../ui/UiMenuSurface.vue";

const props = defineProps({
  /** e.g. "Download a copy as OpenDocument Spreadsheet (.ods)…". */
  copyLabel: { type: String, default: "" },
  /** Why the copy is unavailable; "" when it is available. */
  copyDisabledReason: { type: String, default: "" },
});

const emit = defineEmits(["download", "download-copy"]);

const open = ref(false);
const caret = useTemplateRef("caret");
const copyItem = useTemplateRef("copyItem");

async function toggle() {
  open.value = !open.value;
  if (open.value && !props.copyDisabledReason) {
    await nextTick();
    copyItem.value?.focus();
  }
}

function downloadCopy() {
  open.value = false;
  emit("download-copy");
}
</script>

<template>
  <div class="download-split">
    <UiButton
      data-testid="spreadsheet-download"
      @click="emit('download')"
    >
      Download
    </UiButton>
    <UiButton
      ref="caret"
      class="download-caret"
      data-testid="spreadsheet-download-menu"
      aria-haspopup="menu"
      :aria-expanded="open ? 'true' : 'false'"
      aria-label="More download options"
      title="More download options"
      @click="toggle"
    >
      ▾
    </UiButton>
    <UiMenuSurface
      :open="open"
      :anchor="caret"
      class="download-menu"
      role="menu"
      aria-label="Download options"
      data-testid="spreadsheet-download-menu-popup"
      @dismiss="open = false"
    >
      <button
        ref="copyItem"
        type="button"
        role="menuitem"
        class="download-menu-item"
        data-testid="spreadsheet-download-copy"
        tabindex="-1"
        :disabled="props.copyDisabledReason !== ''"
        :title="props.copyDisabledReason || undefined"
        :aria-describedby="props.copyDisabledReason ? 'spreadsheet-download-copy-reason' : undefined"
        @click="downloadCopy"
      >
        {{ props.copyLabel }}
      </button>
      <p
        v-if="props.copyDisabledReason"
        id="spreadsheet-download-copy-reason"
        class="download-menu-reason"
      >
        {{ props.copyDisabledReason }}
      </p>
    </UiMenuSurface>
  </div>
</template>

<style scoped>
.download-split {
  position: relative;
  display: flex;
  gap: var(--space-1);
  align-items: center;
}

.download-caret {
  padding-inline: var(--space-2);
}

.download-menu {
  position: absolute;
  right: 0;
  top: calc(100% + var(--space-1));
  z-index: 60;
  min-width: 18rem;
  padding: var(--space-1);
}

.download-menu-item {
  display: block;
  width: 100%;
  padding: var(--space-2) var(--space-3);
  border: 0;
  border-radius: var(--radius-sm);
  background: transparent;
  color: var(--text-primary);
  font: inherit;
  font-size: var(--text-md);
  text-align: left;
  cursor: pointer;
}

.download-menu-item:not(:disabled):hover,
.download-menu-item:focus-visible {
  background: var(--surface-sunken);
}

.download-menu-item:focus-visible {
  outline: var(--border-width) solid var(--focus-ring);
  outline-offset: calc(-1 * var(--border-width));
}

.download-menu-item:disabled {
  color: var(--text-disabled);
  cursor: default;
}

.download-menu-reason {
  margin: 0;
  padding: 0 var(--space-3) var(--space-2);
  color: var(--text-secondary);
  font-size: var(--text-xs);
  font-style: italic;
}
</style>
