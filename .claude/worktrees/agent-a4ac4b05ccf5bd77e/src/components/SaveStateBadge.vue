<script setup lang="ts">
// The app bar's "is my work safe" readout, shared by both products: the word
// processor's EditorSaveStatus and the deck header both render this, so the
// two say the same thing in the same place rather than one of them saying
// nothing at all.
//
// Runtime prop declarations — see AppShell.vue's comment on the espree
// constraint.
const props = defineProps({
  /** "saved" | "saving" | "error". Drives the resting-state dot at compact. */
  saveState: { type: String, default: "saved" },
  /** The words themselves, so each product keeps its own phrasing. */
  text: { type: String, required: true },
});
</script>

<template>
  <span
    class="save-state"
    data-testid="save-state"
    :data-save-state="props.saveState"
  >{{ props.text }}</span>
</template>

<style scoped>
.save-state {
  font-size: var(--text-md);
  /* Inherited, not fixed: the word processor renders this on the dark app bar
     (which sets --text-on-brand) and the deck header renders it on a light
     surface. Taking the host's own colour is what lets one component read
     correctly on both, and it is what the compact dot fills itself with. */
  color: inherit;
  white-space: nowrap;
  /* A status that changes width moves whatever sits beside it. "Saving…" and
     "All changes saved" alternate on every write, and the buttons after this
     one shifted by ~50px each time — enough that a button could move out from
     under a pointer between taking aim and clicking, which is exactly how
     Present started missing. Reserving the resting width keeps the row still;
     the error state is longer and one-way, so it may push. */
  display: inline-block;
  min-width: 7rem;
  text-align: right;
}

/* Compact tier (SPEC-068). "All changes saved" is ~110px of nowrap text that
   never shrinks, and on a phone-width bar it was winning that space from the
   document's own name. It is also the least urgent thing there: it says
   nothing changed. So at compact the RESTING state becomes a dot — the
   reassurance survives, its width does not — while "Saving…" and the error
   state keep their words, because those are the ones worth reading.
   `text-indent` rather than `display: none`: the text stays in the DOM and in
   the accessibility tree, so the status is still announced and still
   readable to anything that inspects it. */
@media (max-width: 639.98px) {
  .save-state[data-save-state="saved"] {
    /* The dot sizes itself; the reserved width above would keep the gap. */
    min-width: 0;
    width: var(--space-3);
    height: var(--space-3);
    /* A circle, not the pill token: at 10px square that token's 4px radius
       reads as a rounded box rather than a status dot. */
    border-radius: 50%;
    background: currentColor;
    overflow: hidden;
    text-indent: 100%;
  }
}
</style>
