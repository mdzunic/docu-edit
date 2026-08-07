<script setup lang="ts">
// One rendered page of the paginated canvas (SPEC-028): a fixed-size page
// box whose content area renders the block entries the pagination layout
// assigned to it. A separate component so the entry v-for can carry v-memo
// (Vue only honours v-memo on the v-for element itself — nesting the entry
// loop inside DocumentView's page loop would silently disable it, and the
// memo is what keeps typing responsive on very large documents).
//
// Entries arrive fully precomputed (tag, style objects, run render lists,
// slice windows) from DocumentView's pagedViews computed, so this component
// holds no document logic of its own. Runtime prop/emit declarations, not
// the generic form — this script is parsed by espree (see AppShell.vue).
//
// Header/footer bands (SPEC-029): two more entry lists, rendered as
// absolutely-positioned bands inside the same page box rather than folded
// into the body's content flow — a header/footer never affects pagination
// (lib/pagination.ts is untouched by SPEC-029), it just repeats within the
// page's own top/bottom margin. `headerEditable`/`footerEditable` mark the
// one page currently being edited (only one at a time, see DocumentView's
// activeHeaderFooter) — every other page renders the identical content
// read-only, which is how "edits apply to every page in the section at
// once" (AC-10) falls out for free: they all read the same model content.
import DocBlockEntries from "./DocBlockEntries.vue";
import { noRenderEntries } from "../lib/pageRender";

const props = defineProps({
  pageStyle: { type: Object, required: true },
  contentStyle: { type: Object, required: true },
  /** Precomputed render entries (lib/pageRender.RenderEntry) — the typed
   * default seed gives the template its element type (see StylePicker). */
  entries: { type: Array, default: () => noRenderEntries() },
  /** A virtualized page (SPEC-028): the box renders with its exact layout
   * geometry but none of its content — the pagination controller measures
   * its blocks off-DOM instead. */
  virtual: { type: Boolean, default: false },
  cellRange: { type: Object, default: null },
  headerEntries: { type: Array, default: () => noRenderEntries() },
  footerEntries: { type: Array, default: () => noRenderEntries() },
  headerStyle: { type: Object, default: null },
  footerStyle: { type: Object, default: null },
  headerEditable: { type: Boolean, default: false },
  footerEditable: { type: Boolean, default: false },
  /** Body content dims while a header/footer is being edited (AC-1, AC-2). */
  dimmed: { type: Boolean, default: false },
});

defineEmits([
  "image-select",
  "dblclick-header",
  "dblclick-footer",
  "dblclick-body",
  "hf-beforeinput",
  "hf-keydown",
  "hf-compositionstart",
  "hf-compositionend",
]);
</script>

<template>
  <div
    class="page"
    data-testid="document-page"
    :data-virtual="virtual ? '1' : undefined"
    :style="pageStyle"
  >
    <div
      class="page-header-band"
      data-testid="page-header-band"
      :style="headerStyle"
      :contenteditable="headerEditable"
      :data-hf-editable="headerEditable ? '1' : undefined"
      spellcheck="false"
      @dblclick="$emit('dblclick-header')"
      @beforeinput.stop="headerEditable ? $emit('hf-beforeinput', $event) : undefined"
      @keydown.stop="headerEditable ? $emit('hf-keydown', $event) : undefined"
      @compositionstart.stop="headerEditable ? $emit('hf-compositionstart') : undefined"
      @compositionend.stop="headerEditable ? $emit('hf-compositionend', $event) : undefined"
    >
      <DocBlockEntries :entries="headerEntries" />
    </div>
    <div
      class="page-content"
      :style="contentStyle"
      @dblclick="$emit('dblclick-body')"
    >
      <div
        class="page-margin"
        data-testid="page-margin"
        :class="{ 'page-margin-dimmed': dimmed }"
      >
        <DocBlockEntries
          :entries="entries"
          :cell-range="cellRange"
          @image-select="(section, block) => $emit('image-select', section, block)"
        />
      </div>
    </div>
    <div
      class="page-footer-band"
      data-testid="page-footer-band"
      :style="footerStyle"
      :contenteditable="footerEditable"
      :data-hf-editable="footerEditable ? '1' : undefined"
      spellcheck="false"
      @dblclick="$emit('dblclick-footer')"
      @beforeinput.stop="footerEditable ? $emit('hf-beforeinput', $event) : undefined"
      @keydown.stop="footerEditable ? $emit('hf-keydown', $event) : undefined"
      @compositionstart.stop="footerEditable ? $emit('hf-compositionstart') : undefined"
      @compositionend.stop="footerEditable ? $emit('hf-compositionend', $event) : undefined"
    >
      <DocBlockEntries :entries="footerEntries" />
    </div>
  </div>
</template>

<style scoped>
/* A page-shaped surface sized by the layout's own geometry (inline style,
   real centimetres). 12px of margin is AC-2's visible gap between pages;
   overflow clips provisional overhang while measurement settles. */
.page {
  position: relative;
  margin: 0 auto 12px;
  background: #fff;
  box-shadow: 0 1px 3px rgb(0 0 0 / 0.3);
  box-sizing: border-box;
  overflow: hidden;
}

.page:last-child {
  margin-bottom: 0;
}

/* The content area: padding carries the page margins (inline style), so the
   blocks flow in normal document flow exactly where the engine placed
   them. */
.page-content {
  box-sizing: border-box;
  height: 100%;
}

/* Body content dims while a header/footer is being edited (SPEC-029
 * AC-1, AC-2) — opacity, not visibility, so the dimmed text stays a
 * legible reading reference the way Word's own dimming does.
 * `pointer-events: none` sits on the margin box (the entries' own
 * container), not `.page-content` itself, so a double-click still lands on
 * `.page-content`'s own dblclick-body listener above it in the DOM and
 * exits editing (AC-8) instead of being swallowed by the dimmed area. */
.page-margin-dimmed {
  opacity: 0.35;
  pointer-events: none;
}

/* Header/footer bands (SPEC-029): positioned within the page's own
   top/bottom margin — inline top/left/right/bottom come from
   DocumentView's headerStyle/footerStyle, computed from the section's
   margins and a fixed header/footer distance from the page edge. */
.page-header-band,
.page-footer-band {
  position: absolute;
  overflow: hidden;
  box-sizing: border-box;
  display: flex;
  flex-direction: column;
}

/* Header text sits at the bottom of the top margin, just above the body
   text's own top edge; footer text sits at the top of the bottom margin,
   just below the body text's own bottom edge — matching where Word
   positions header/footer content within the margin band. */
.page-header-band {
  justify-content: flex-end;
  padding-bottom: 0.2cm;
}

.page-footer-band {
  justify-content: flex-start;
  padding-top: 0.2cm;
}

.page-header-band[data-hf-editable],
.page-footer-band[data-hf-editable] {
  outline: 1px dashed #1a73e8;
  outline-offset: 2px;
}

/* Marks where the page margins end and content begins — an outline, not a
   border, so the rendered content geometry matches the engine's
   margin-box math exactly. */
.page-margin {
  box-sizing: border-box;
  height: 100%;
  outline: 1px dashed #c8c8c8;
  outline-offset: -1px;
}

</style>
