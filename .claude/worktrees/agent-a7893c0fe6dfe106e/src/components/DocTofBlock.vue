<script setup lang="ts">
// A table of figures block on the canvas (SPEC-055): entries generated from
// the document's Caption-styled paragraphs of one label, each a clickable jump
// to its caption, with an "Update table of figures" control shown when the
// document's captions or pagination have drifted from what the block last
// generated. Mirrors DocTocBlock exactly — not directly editable (the parent
// marks it `contenteditable="false"` via `entry.uneditable`). Runtime
// prop/emit declarations, not the generic form — espree-parsed script.
import { asTof } from "../lib/documentModel";
import { TOF_PLACEHOLDER_TEXT } from "../lib/tof";

const props = defineProps({
  /** The TableOfFigures block (documentModel.TableOfFigures). */
  block: { type: Object, required: true },
  section: { type: Number, required: true },
  blockIndex: { type: Number, required: true },
  /** True when the document's current captions of this label no longer match
   * this block's entries (SPEC-055). */
  staleEntries: { type: Boolean, default: false },
  /** True when an entry's cached page number no longer matches its caption's
   * current page (SPEC-055). */
  stalePagination: { type: Boolean, default: false },
});

const emit = defineEmits(["action"]);

function jump(anchor = "") {
  emit("action", { type: "jump", anchor, section: props.section, blockIndex: props.blockIndex });
}

function updatePageNumbers() {
  emit("action", { type: "update-page-numbers", section: props.section, blockIndex: props.blockIndex });
}

function updateEntire() {
  emit("action", { type: "update-entire", section: props.section, blockIndex: props.blockIndex });
}

/** A click that isn't an entry link or update button — tracked as a "selected"
 * block so a typing attempt right after can be caught (same as DocTocBlock). */
function select() {
  emit("action", { type: "select", section: props.section, blockIndex: props.blockIndex });
}
</script>

<template>
  <div
    class="doc-tof"
    data-testid="tof-block"
    :data-section="section"
    :data-block="blockIndex"
    @click="select"
  >
    <div
      v-if="staleEntries || stalePagination"
      class="tof-update-bar print-exclude"
      data-testid="tof-update-bar"
    >
      <span class="tof-update-label">Update table of figures</span>
      <button
        type="button"
        class="tof-update-button"
        data-testid="tof-update-page-numbers"
        @mousedown.prevent
        @click.stop="updatePageNumbers"
      >
        Update page numbers only
      </button>
      <button
        type="button"
        class="tof-update-button"
        data-testid="tof-update-entire"
        @mousedown.prevent
        @click.stop="updateEntire"
      >
        Update entire table
      </button>
    </div>
    <p
      v-if="asTof(block).entries.length === 0"
      class="tof-placeholder"
      data-testid="tof-placeholder"
    >
      {{ TOF_PLACEHOLDER_TEXT }}
    </p>
    <p
      v-for="(entry, index) in asTof(block).entries"
      :key="`${entry.targetBookmark}:${index}`"
      class="tof-entry"
      data-testid="tof-entry"
    >
      <a
        href="#"
        class="tof-entry-link"
        data-tof-link=""
        @click.prevent.stop="jump(entry.targetBookmark)"
      >{{ entry.text || " " }}</a>
      <template v-if="asTof(block).options.pageNumbers && asTof(block).options.hyperlinks">
        <span
          v-if="asTof(block).options.dotLeader"
          class="tof-leader"
          aria-hidden="true"
        />
        <span
          v-else
          class="tof-spacer"
        />
        <span class="tof-page-number">{{ entry.page ?? "" }}</span>
      </template>
    </p>
  </div>
</template>

<style scoped>
.doc-tof {
  display: block;
  user-select: none;
}

.tof-update-bar {
  display: flex;
  align-items: center;
  gap: 0.5rem;
  margin-bottom: 0.5rem;
  padding: 0.375rem 0.5rem;
  background: #fef7e0;
  border: 1px solid #f2c94c;
  border-radius: 4px;
  font-size: 0.8125rem;
}

.tof-update-label {
  font-weight: 600;
  color: #5f5228;
  margin-inline-end: auto;
}

.tof-update-button {
  font: inherit;
  font-size: 0.75rem;
  padding: 0.25rem 0.5rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
  white-space: nowrap;
}

.tof-placeholder {
  margin: 0;
  color: #5f6368;
  font-style: italic;
}

.tof-entry {
  display: flex;
  align-items: baseline;
  margin: 0 0 0.15em;
  gap: 0.25em;
}

.tof-entry-link {
  color: inherit;
  text-decoration: none;
  cursor: pointer;
}

.tof-entry-link:hover {
  text-decoration: underline;
}

.tof-leader {
  flex: 1 1 auto;
  border-bottom: 1px dotted #5f6368;
  margin: 0 0.25em;
  align-self: end;
  min-width: 0.5em;
}

.tof-spacer {
  flex: 1 1 auto;
}

.tof-page-number {
  white-space: nowrap;
}
</style>
