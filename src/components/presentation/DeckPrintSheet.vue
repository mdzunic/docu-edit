<script setup lang="ts">
// The deck's print/export sheet (SPEC-104): the pages planSheet() laid out,
// rendered off-screen — never the editor stage — with one
// `<SlideStage :interactive="false">` per slot, the "Frame slides" outlines,
// the ruled note lines, a notes page's speaker notes and the 9 pt page
// number. useDeckExport mounts a rolling window of pages here for PDF/PNG
// export (each page magnified so its slides render untransformed — see
// sheetPageViews) and every page for browser print, which shows only this
// sheet under @media print (App.vue's `body.is-deck-print` rules). It is
// teleported to <body> so those rules can hide everything else, and it is
// `aria-hidden` and `inert`: nothing here is for the screen. Runtime prop
// declarations per the espree constraint.
import { computed } from "vue";
import SlideStage from "./SlideStage.vue";
import { asDeck, asSheetPages, sheetPageViews } from "./deckPrintSheet";
import { now } from "../../lib/clock";

const props = defineProps({
  /** The useDeckExport() controller: its `deck` (read, never written), the
   * mounted `sheetPages` and the `sheetMode` they render in. */
  controller: { type: Object, required: true },
});

const c = props.controller;
const deck = computed(() => asDeck(c.deck.value));
const pages = computed(() =>
  deck.value ? sheetPageViews(deck.value, asSheetPages(c.sheetPages.value), c.sheetMode.value, now()) : [],
);
</script>

<template>
  <Teleport to="body">
    <div
      v-if="pages.length > 0"
      class="deck-print-sheet"
      data-testid="deck-print-sheet"
      aria-hidden="true"
      inert
    >
      <div
        v-for="page in pages"
        :key="page.index"
        class="deck-print-page"
        data-testid="deck-print-page"
        :data-page-index="page.index"
        :style="page.style"
      >
        <div
          v-for="slot in page.slots"
          :key="slot.key"
          class="deck-print-slot"
          data-testid="deck-print-slot"
          :style="slot.style"
        >
          <SlideStage
            :slide="deck.slides[slot.slideIndex]"
            :width-cm="deck.slideWidthCm"
            :height-cm="deck.slideHeightCm"
            :media="deck.media || []"
            :scale="slot.stageScale"
            :interactive="false"
            :raster-scale="2"
            :flat-rotation="true"
            :fields="slot.fields"
            :theme="deck.theme"
          />
        </div>
        <div
          v-for="frame in page.frames"
          :key="`frame-${frame.key}`"
          class="deck-print-frame"
          :style="frame.style"
        />
        <div
          v-for="line in page.noteLines"
          :key="`line-${line.key}`"
          class="deck-print-note-line"
          :style="line.style"
        />
        <div
          v-if="page.notes"
          class="deck-print-notes"
          :style="page.notes.style"
        >
          <p
            v-for="(paragraph, index) in page.notes.paragraphs"
            :key="index"
            class="deck-print-notes-paragraph"
          >
            {{ paragraph }}
          </p>
        </div>
        <div
          v-if="page.pageNumber"
          class="deck-print-page-number"
          data-testid="deck-print-page-number"
          :style="page.pageNumber.style"
        >
          {{ page.pageNumber.text }}
        </div>
      </div>
    </div>
  </Teleport>
</template>

<style scoped>
/* Off-screen, but laid out and painted for real: the exporter measures and
   reads it, and printing shows it (App.vue moves it back on screen). */
.deck-print-sheet {
  position: fixed;
  left: -20000px;
  top: 0;
  pointer-events: none;
  background: #ffffff;
  color: #000000;
  -webkit-print-color-adjust: exact;
  print-color-adjust: exact;
}

.deck-print-page {
  position: relative;
  overflow: hidden;
  background: #ffffff;
  break-after: page;
}

.deck-print-page:last-child {
  break-after: auto;
}

.deck-print-slot {
  position: absolute;
  overflow: hidden;
}

/* The paper is the backdrop; a slide casts no screen shadow on it. */
.deck-print-slot :deep(.stage-surface) {
  box-shadow: none;
}

.deck-print-frame {
  position: absolute;
  box-sizing: border-box;
}

.deck-print-note-line {
  position: absolute;
  height: 0;
}

.deck-print-notes {
  position: absolute;
  overflow: hidden;
  font-family: "Liberation Sans", Arial, sans-serif;
  color: #000000;
}

.deck-print-notes-paragraph {
  margin: 0;
  min-height: 1lh; /* an empty notes paragraph is still a blank line */
  white-space: pre-wrap;
  overflow-wrap: anywhere;
}

.deck-print-page-number {
  position: absolute;
  display: flex;
  align-items: center;
  justify-content: center;
  font-family: "Liberation Sans", Arial, sans-serif;
  color: #000000;
}
</style>
