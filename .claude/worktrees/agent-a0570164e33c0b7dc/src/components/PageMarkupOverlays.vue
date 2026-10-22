<script setup lang="ts">
// The three page-relative markup overlays that sit inside DocumentView's
// `.pages` container: the find-match highlights (SPEC-036), the comment-anchor
// highlights (SPEC-038) and the spell-check underlines (SPEC-035). They are
// purely presentational — each renders a list of pre-measured rects the view
// computes — so they were extracted out of the DocumentView backbone to keep
// it within its line budget (SPEC-064). The elements are declared as fragment
// roots (no wrapper) so they stay direct children of `.pages` and their
// `position: absolute; inset: 0` still resolves against it, exactly as before.
import { asHighlightRects } from "../lib/findReplace";
import { asCommentHighlightRects } from "../views/document/useComments";
import { asSpellUnderlineRects } from "../views/document/useSpellCheck";

// Runtime prop declarations (this script is espree-parsed — see DocPage.vue);
// the template narrows each `unknown[]` prop through the `as*` cast helpers,
// the same idiom DocumentRuler uses for its tab-stop array.
defineProps({
  findVisible: { type: Boolean, default: false },
  findHighlightRects: { type: Array, default: () => [] },
  commentHighlightRects: { type: Array, default: () => [] },
  spellUnderlineRects: { type: Array, default: () => [] },
});

const emit = defineEmits(["focus-thread"]);

function focusThreadFromHighlight(threadId = "") {
  emit("focus-thread", threadId);
}
</script>

<template>
  <div
    v-if="findVisible"
    class="find-highlight-layer print-exclude"
    data-testid="find-highlights"
    aria-hidden="true"
  >
    <div
      v-for="(rect, index) in asHighlightRects(findHighlightRects)"
      :key="index"
      class="find-highlight"
      :class="{ current: rect.current }"
      :data-current="rect.current ? 'true' : 'false'"
      :style="{
        left: `${rect.left}px`,
        top: `${rect.top}px`,
        width: `${rect.width}px`,
        height: `${rect.height}px`,
      }"
    />
  </div>
  <div
    v-if="commentHighlightRects.length > 0"
    class="comment-highlight-layer print-exclude"
    data-testid="comment-highlights"
  >
    <div
      v-for="(rect, index) in asCommentHighlightRects(commentHighlightRects)"
      :key="index"
      class="comment-highlight"
      :class="{ active: rect.active }"
      :data-testid="`comment-highlight-${rect.threadId}`"
      :style="{
        left: `${rect.left}px`,
        top: `${rect.top}px`,
        width: `${rect.width}px`,
        height: `${rect.height}px`,
      }"
      @click="focusThreadFromHighlight(rect.threadId)"
    />
  </div>
  <!-- The rects address a paragraph offset range, but they are NOT
     blocks: the address is spelled data-spell-* rather than
     data-section/data-block because that exact pair is
     BLOCK_SELECTOR (lib/domCaret.ts), the selector caret mapping,
     pagination and page navigation use to enumerate real block
     elements. Naming them the same way makes every flagged word
     masquerade as an extra slice of its own paragraph. -->
  <div
    v-if="spellUnderlineRects.length > 0"
    class="spell-underline-layer print-exclude"
    data-testid="spell-underlines"
    aria-hidden="true"
  >
    <div
      v-for="(rect, index) in asSpellUnderlineRects(spellUnderlineRects)"
      :key="index"
      class="spell-underline"
      data-testid="spell-underline"
      :data-spell-section="rect.section"
      :data-spell-block="rect.block"
      :data-spell-start="rect.start"
      :data-spell-end="rect.end"
      :data-word="rect.word"
      :data-lang="rect.lang"
      :style="{
        left: `${rect.left}px`,
        top: `${rect.top}px`,
        width: `${rect.width}px`,
        height: `${rect.height}px`,
      }"
    />
  </div>
</template>

<style scoped>
/* The highlight overlay sits inside the (positioned) .page, behind the text —
   .page-margin renders after it, so the text paints on top and stays
   readable. pointer-events: none lets clicks reach the text underneath. */
.find-highlight-layer {
  position: absolute;
  inset: 0;
  z-index: 0;
  pointer-events: none;
}

.find-highlight {
  position: absolute;
  background: rgb(255 214 0 / 0.45);
  border-radius: 1px;
}

.find-highlight.current {
  background: rgb(255 145 0 / 0.65);
  outline: 1px solid #f57c00;
}

/* The spell-check underline overlay (SPEC-035 AC-1). Unlike
   .find-highlight-layer (intentionally *behind* the text it highlights),
   this must paint ABOVE the page: DocPage's own .page root is itself
   `position: relative` with no z-index (stack level 0, CSS2.1 Appendix E
   step 6), the same tier a sibling z-index:0 layer occupies — ties within
   that tier resolve by DOM order, and .page (containing every page's
   opaque background and all its text) comes after this layer in the
   markup, so a bare z-index:0 here would paint BEHIND every page's
   background and be invisible (verified by screenshot while building this:
   z-index:0 here is empirically hidden). z-index:1 moves this layer into
   the next stacking tier (CSS2.1 step 7, positive z-index), which always
   paints after every stack-level-0 box regardless of DOM order — still well
   under the table/image handle layers' z-index:15. */
.spell-underline-layer {
  position: absolute;
  inset: 0;
  z-index: 1;
  pointer-events: none;
}

.spell-underline {
  position: absolute;
  /* Inert, and it must stay that way. Each rect covers a flagged word's whole
     glyph box (not a 3px sliver), and the layer paints above the text — so
     opting into pointer events here makes the overlay swallow every plain
     click, double-click and drag that starts on a misspelled word, leaving
     the caret unset and the selection empty. Right-clicks are hit-tested
     against these same rects in spellRectAtPoint() instead. */
  pointer-events: none;
  background-repeat: repeat-x;
  background-position: left bottom;
  background-size: 6px 4px;
  background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='6' height='4' viewBox='0 0 6 4'%3E%3Cpath d='M0 3 Q1.5 0 3 3 T6 3' fill='none' stroke='%23d93025' stroke-width='1'/%3E%3C/svg%3E");
}

/* The comment highlight overlay (SPEC-038): same layering convention as
   .find-highlight-layer, but each rect IS clickable (AC "clicking a
   highlighted range focuses its thread"), so pointer-events stays auto on
   the individual rects while the layer itself still passes clicks through
   the gaps between them. */
.comment-highlight-layer {
  position: absolute;
  inset: 0;
  /* Unlike .find-highlight-layer (z-index: 0, never needs to catch clicks),
     these rects must win hit-testing against the text underneath — the
     same z-index band the table handles (15, TableHandlesOverlay.vue) use
     for their own clickable overlay controls. */
  z-index: 5;
  pointer-events: none;
}

.comment-highlight {
  position: absolute;
  background: rgb(26 115 232 / 0.18);
  border-bottom: 2px solid rgb(26 115 232 / 0.55);
  pointer-events: auto;
  cursor: pointer;
}

.comment-highlight.active {
  background: rgb(26 115 232 / 0.32);
  border-bottom-color: #1a73e8;
}
</style>
