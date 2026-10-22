<script setup lang="ts">
// SPEC-056 equation editor host: the floating overlay that mounts the
// EquationEditor over the live document when an equation edit session is open.
// It takes the useEquations() return as a single `api` prop and destructures
// the session refs/handlers from it, so DocumentView's wiring stays a one-line
// `<EquationEditorHost :api="equations" />` (the size-budget contract) rather
// than plumbing every session ref and event through the template itself.
// Runtime prop declaration, not defineProps<{...}>() generics — the same
// espree constraint every other .vue here follows (see CommentsPane.vue); the
// cast below is that file's typed-access seam.
import EquationEditor from "./EquationEditor.vue";
import { asEquationsApi } from "../views/document/useEquations";

const props = defineProps({
  api: { type: Object, required: true },
});
const {
  equationEditState,
  equationLinearForm,
  equationAltText,
  equationOp,
  setEquationAltText,
  commitEquation,
  closeEquationEditor,
} = asEquationsApi(props.api);
</script>

<template>
  <div
    v-if="equationEditState"
    class="equation-editor-overlay"
    style="position: fixed; right: 1.5rem; bottom: 1.5rem; z-index: 60"
  >
    <EquationEditor
      :body="equationEditState.body"
      :linear="equationLinearForm()"
      :alt-text="equationAltText"
      @op="equationOp"
      @alt-text="setEquationAltText"
      @commit="commitEquation"
      @close="closeEquationEditor"
    />
  </div>
</template>
