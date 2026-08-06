<script setup lang="ts">
import { imageDisplayStyle } from "../lib/renderFormatting";
import { asImage } from "../lib/documentModel";

// One image block on the canvas (SPEC-024). Runtime prop declarations, not
// defineProps<{...}>() generics — see AppShell.vue's comment on the same
// espree constraint DocTableBlock.vue also follows. `data-section`/
// `data-block` sit on the anchor span (not the `<img>` itself) so
// lib/domCaret.ts's `[data-section][data-block]` lookups resolve to a
// stable element regardless of the image's own wrap-mode positioning
// (behind/infront images position the `<img>` absolutely relative to this
// anchor — see lib/renderFormatting.ts's imageDisplayStyle).
// The parent (DocumentView.vue) also binds data-section/data-block on this
// component's own `<component :is>` tag (as `undefined`, mirroring
// DocTableBlock's convention) so they never leak onto an unrelated element
// via Vue's attribute fallthrough — but unlike DocTableBlock (whose root
// element sets neither), this component's own root DOES set them, and
// fallthrough attrs (even undefined ones) take precedence over a
// component's own template bindings of the same name. inheritAttrs: false
// disables that fallthrough entirely, so the explicit bindings below win.
defineOptions({ inheritAttrs: false });

const props = defineProps({
  /** The Image block (documentModel.Image). */
  block: { type: Object, required: true },
  section: { type: Number, required: true },
  blockIndex: { type: Number, required: true },
  /** A data: URI resolved from the document's media (DocumentView.vue
   * looks this up — the block only carries a mediaId). */
  src: { type: String, required: true },
  selected: { type: Boolean, default: false },
});

defineEmits(["select"]);

const floatingWrap = ["behind", "infront"];
</script>

<template>
  <span
    class="doc-image-anchor"
    :class="{ 'doc-image-anchor-floating': floatingWrap.includes(block.wrap) }"
    :data-section="section"
    :data-block="blockIndex"
  >
    <img
      class="doc-image"
      :class="{ 'doc-image-selected': selected }"
      :src="props.src"
      :alt="block.altText || ''"
      :style="imageDisplayStyle(asImage(block))"
      data-testid="doc-image"
      draggable="false"
      @click.stop="$emit('select')"
    >
  </span>
</template>

<style scoped>
.doc-image-anchor {
  display: inline-block;
  line-height: 0;
}

.doc-image-anchor-floating {
  position: relative;
}

.doc-image {
  display: block;
  cursor: pointer;
}

.doc-image-selected {
  outline: 2px solid #1a73e8;
  outline-offset: 1px;
}
</style>
