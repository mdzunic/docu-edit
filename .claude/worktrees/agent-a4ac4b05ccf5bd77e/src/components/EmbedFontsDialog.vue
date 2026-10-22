<script setup lang="ts">
import BaseDialog from "./BaseDialog.vue";
import { asStringList } from "../views/document/useFonts";

// "Embed fonts in this file" (SPEC-048 AC-15/16/17/18): states the size cost
// before saving (AC-16) and names the faces skipped because their licence
// bits forbid embedding (AC-18). Confirm writes the document with the subset
// fonts embedded. Runtime prop/emit declarations (espree parses .vue scripts).
const props = defineProps({
  faceCount: { type: Number, default: 0 },
  sizeText: { type: String, default: "" },
  skipped: { type: Array, default: () => [] },
  busy: { type: Boolean, default: false },
});

const emit = defineEmits(["confirm", "close"]);
</script>

<template>
  <BaseDialog
    title="Embed fonts in this file"
    testid="embed-fonts-dialog"
    @dismiss="emit('close')"
  >
    <p
      class="dialog-message"
      data-testid="embed-fonts-size"
    >
      Embedding {{ faceCount }} font{{ faceCount === 1 ? "" : "s" }} adds about
      <strong>{{ sizeText }}</strong> to the file.
    </p>
    <div
      v-if="asStringList(props.skipped).length > 0"
      class="embed-skipped"
      data-testid="embed-fonts-skipped"
    >
      <p class="embed-skipped-title">
        These fonts can't be embedded — their licence forbids it:
      </p>
      <ul>
        <li
          v-for="name in asStringList(props.skipped)"
          :key="name"
        >
          {{ name }}
        </li>
      </ul>
    </div>
    <template #actions>
      <button
        type="button"
        data-testid="embed-fonts-cancel"
        @click="emit('close')"
      >
        Cancel
      </button>
      <button
        type="button"
        class="primary"
        data-testid="embed-fonts-confirm"
        :disabled="busy"
        @click="emit('confirm')"
      >
        Save with fonts embedded
      </button>
    </template>
  </BaseDialog>
</template>

<style scoped>
.dialog-message {
  margin: 0 0 var(--space-4);
  color: var(--text-primary);
  font-size: var(--text-lg);
}

.embed-skipped {
  font-size: var(--text-md);
  color: var(--text-secondary);
}

.embed-skipped-title {
  margin: 0 0 var(--space-2);
}

.embed-skipped ul {
  margin: 0;
  padding-left: var(--space-5);
}
</style>
