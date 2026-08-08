<script setup lang="ts">
// A table of contents block on the canvas (SPEC-033): entries generated
// from the document's Heading 1-6 paragraphs, each a clickable jump to its
// heading (AC-8, AC-9), with an "Update table of contents" control shown
// when the document's headings or pagination have drifted from what the
// block last generated (AC-10, AC-11). Not directly editable — there is no
// contenteditable interior at all, only clickable entries and buttons (the
// parent marks this block `contenteditable="false"` via the shared
// `entry.uneditable` convention, same as a manual break). Runtime prop/emit
// declarations, not the generic form — this script is parsed by espree (see
// AppShell.vue's comment on the same constraint).
import { asToc } from "../lib/documentModel";
import { TOC_PLACEHOLDER_TEXT } from "../lib/toc";

const props = defineProps({
  /** The TableOfContents block (documentModel.TableOfContents). */
  block: { type: Object, required: true },
  section: { type: Number, required: true },
  blockIndex: { type: Number, required: true },
  /** True when the document's current Heading 1-6 paragraphs no longer
   * match this block's entries (SPEC-033 AC-10). */
  staleHeadings: { type: Boolean, default: false },
  /** True when an entry's cached page number no longer matches its
   * heading's current page (SPEC-033 AC-11). */
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

/** A click anywhere in the block that isn't an entry link or an update
 * button (SPEC-033 AC-14, AC-15): a native click never places a caret
 * inside a `contenteditable="false"` region, so the view tracks this as a
 * "selected" block instead — the same purpose SPEC-024's image selection
 * serves — to catch a typing attempt right after. */
function select() {
  emit("action", { type: "select", section: props.section, blockIndex: props.blockIndex });
}
</script>

<template>
  <div
    class="doc-toc"
    data-testid="toc-block"
    :data-section="section"
    :data-block="blockIndex"
    @click="select"
  >
    <div
      v-if="staleHeadings || stalePagination"
      class="toc-update-bar print-exclude"
      data-testid="toc-update-bar"
    >
      <span class="toc-update-label">Update table of contents</span>
      <button
        type="button"
        class="toc-update-button"
        data-testid="toc-update-page-numbers"
        @mousedown.prevent
        @click.stop="updatePageNumbers"
      >
        Update page numbers only
      </button>
      <button
        type="button"
        class="toc-update-button"
        data-testid="toc-update-entire"
        @mousedown.prevent
        @click.stop="updateEntire"
      >
        Update entire table
      </button>
    </div>
    <p
      v-if="asToc(block).entries.length === 0"
      class="toc-placeholder"
      data-testid="toc-placeholder"
    >
      {{ TOC_PLACEHOLDER_TEXT }}
    </p>
    <p
      v-for="(entry, index) in asToc(block).entries"
      :key="`${entry.anchor}:${index}`"
      class="toc-entry"
      data-testid="toc-entry"
      :style="{ paddingInlineStart: `${(entry.level - 1) * 1.27}cm` }"
    >
      <a
        href="#"
        class="toc-entry-link"
        data-toc-link=""
        @click.prevent.stop="jump(entry.anchor)"
      >{{ entry.text || " " }}</a>
      <template v-if="asToc(block).options.showPageNumbers && !asToc(block).options.hyperlinksOnly">
        <span
          class="toc-leader"
          aria-hidden="true"
        />
        <span class="toc-page-number">{{ entry.pageNumber ?? "" }}</span>
      </template>
    </p>
  </div>
</template>

<style scoped>
.doc-toc {
  display: block;
  user-select: none;
}

.toc-update-bar {
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

.toc-update-label {
  font-weight: 600;
  color: #5f5228;
  margin-inline-end: auto;
}

.toc-update-button {
  font: inherit;
  font-size: 0.75rem;
  padding: 0.25rem 0.5rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
  white-space: nowrap;
}

.toc-placeholder {
  margin: 0;
  color: #5f6368;
  font-style: italic;
}

.toc-entry {
  display: flex;
  align-items: baseline;
  margin: 0 0 0.15em;
  gap: 0.25em;
}

.toc-entry-link {
  color: inherit;
  text-decoration: none;
  cursor: pointer;
}

.toc-entry-link:hover {
  text-decoration: underline;
}

.toc-leader {
  flex: 1 1 auto;
  border-bottom: 1px dotted #5f6368;
  margin: 0 0.25em;
  align-self: end;
  min-width: 0.5em;
}

.toc-page-number {
  white-space: nowrap;
}
</style>
