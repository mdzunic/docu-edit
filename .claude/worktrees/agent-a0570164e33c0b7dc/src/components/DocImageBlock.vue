<script setup lang="ts">
import { onMounted, onUpdated, ref, useTemplateRef } from "vue";
import { drawingWrapSpacerStyle, imageDisplayStyle, trackChangeBlockHidden } from "../lib/renderFormatting";
import { pxToCm } from "../lib/core/units";
import { asImage, asTrackChangesDisplayMode } from "../lib/documentModel";

// One image block on the canvas (SPEC-024). Runtime prop declarations, not
// defineProps<{...}>() generics — see AppShell.vue's comment on the same
// espree constraint DocTableBlock.vue also follows. `data-section`/
// `data-block` sit on the anchor span (not the `<img>` itself) so
// lib/domCaret.ts's `[data-section][data-block]` lookups resolve to a
// stable element regardless of the image's own wrap-mode positioning
// (a floating image positions the `<img>` absolutely relative to this
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
  /** The review pane's display mode (SPEC-039 AC-9) — a tracked-inserted
   * image is hidden in "Original" (as if rejected), a tracked-deleted one
   * hidden in "No markup" (as if accepted). */
  displayMode: { type: String, default: "all" },
});

defineEmits(["select"]);

// Every non-inline wrap now positions absolutely from the anchor's own
// top-left (imageDisplayStyle), so all four floating modes need the
// zero-footprint anchor — not just behind/infront as before, when `square`
// was a plain CSS float with no offset to honour.
const floatingWrap = ["square", "topAndBottom", "behind", "infront"];

// The text-column width the wrap spacer needs (drawingWrapSpacerStyle):
// measured off the anchor's own block container, so an image inside a
// SPEC-030 column wraps within that column. Mirrors DocShapeBlock.vue's
// identical measure-in-component pattern.
const anchorEl = useTemplateRef("anchorEl");
const columnWidthCm = ref(0);

function measureColumnWidth() {
  const parent = anchorEl.value instanceof HTMLElement ? anchorEl.value.parentElement : null;
  const cm = parent ? pxToCm(parent.clientWidth) : 0;
  if (Math.abs(cm - columnWidthCm.value) > 0.01) {
    columnWidthCm.value = cm;
  }
}

onMounted(measureColumnWidth);
onUpdated(measureColumnWidth);

/** The in-flow float that reserves this image's text room (square /
 * topAndBottom) — null for every other wrap mode. The image itself stays an
 * absolute overlay, so this is what makes its wrap real rather than
 * decorative, exactly as it does for a shape. */
function wrapSpacerStyle() {
  return drawingWrapSpacerStyle(asImage(props.block), columnWidthCm.value);
}
</script>

<template>
  <span
    ref="anchorEl"
    class="doc-image-anchor"
    :class="{ 'doc-image-anchor-floating': floatingWrap.includes(block.wrap) }"
    :data-section="section"
    :data-block="blockIndex"
  >
    <!-- The wrap spacer: an invisible float in the TEXT flow that the
         surrounding lines yield to (square / topAndBottom). The image below
         stays an absolute overlay; this is what makes its wrap mode real. -->
    <span
      v-if="wrapSpacerStyle() && !trackChangeBlockHidden(block.changes, asTrackChangesDisplayMode(displayMode))"
      class="doc-image-wrap-spacer"
      data-testid="image-wrap-spacer"
      aria-hidden="true"
      :style="wrapSpacerStyle()"
    />
    <img
      v-show="!trackChangeBlockHidden(block.changes, asTrackChangesDisplayMode(displayMode))"
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

/* SPEC-078's zero-footprint floating anchor, the same one DocShapeBlock.vue
   uses and for the same reasons (see its own comment): a zero-height BLOCK,
   so the wrap spacer's float can escape into the page's text flow instead of
   being trapped in an inline-block's own formatting context, and so the
   anchor's top-left stays pinned at the anchor paragraph's top-left — the
   origin the image's `left`/`top` (offsetXCm/offsetYCm) measure from. An
   inline-block here added a baseline's worth of drift, which is exactly the
   "lands below where the file said" bug this shape of anchor exists to
   avoid. */
.doc-image-anchor-floating {
  display: block;
  height: 0;
  position: relative;
}

.doc-image-wrap-spacer {
  pointer-events: none;
}

.doc-image {
  display: block;
  cursor: pointer;
}

.doc-image-selected {
  outline: 2px solid var(--accent);
  outline-offset: 1px;
}
</style>
