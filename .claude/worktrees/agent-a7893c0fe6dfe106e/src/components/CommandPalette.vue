<script setup lang="ts">
// The command palette (SPEC-062): one accessible combobox/listbox dialog used
// by both the editor and the library, over two command sources. It renders
// and dispatches only — the searchable catalogue, the pure ranking, recents
// and enablement all arrive as props/data from the framework-free lib and the
// hosting view, so this component holds no product logic.
//
// Reuses BaseDialog for the modal shell it already owns — backdrop-click and
// Escape dismissal (AC-14, AC-25), initial focus and the Tab trap (AC-35).
// The combobox lives in its default slot: the input keeps focus while
// Up/Down move a virtual highlight (aria-activedescendant), never stealing the
// document selection (AC-15, AC-26 — the view keeps the model selection; this
// dialog only reads it).
//
// Runtime prop/emit declarations for the espree/ESLint constraint (see
// AppShell.vue).
import { computed, ref, watch, useTemplateRef, nextTick } from "vue";
import BaseDialog from "./BaseDialog.vue";
import AppIcon from "./AppIcon.vue";
import { searchCommands, asCommandMetas, asIdList, asReasonMap, noRanges } from "../lib/commandCatalogue";
import { shortcutLabel } from "../lib/shortcuts";

const props = defineProps({
  /** The commands this surface offers — the catalogue filtered to its scope
   * (document or library). */
  commands: { type: Array, default: () => [] },
  /** Extra, per-document virtual rows (the "Apply style: …" rows) merged into
   * the same search. Their ids are opaque to this component — it emits them
   * back on run and the view maps them to actions. */
  extraRows: { type: Array, default: () => [] },
  /** id -> "" when the command is enabled, or a short reason when not. Rows
   * without an entry are treated as enabled. */
  availability: { type: Object, default: () => ({}) },
  /** Recent command ids, most recent first (drives the empty-state list and
   * the recency tie-break). */
  recentIds: { type: Array, default: () => [] },
  /** Common command ids shown after recents in the empty state. */
  commonIds: { type: Array, default: () => [] },
  /** The dialog's accessible title. */
  title: { type: String, default: "Search commands" },
  testid: { type: String, default: "command-palette" },
});

const emit = defineEmits(["run", "dismiss"]);

const query = ref("");
const activeIndex = ref(0);
const listbox = useTemplateRef("listbox");

// Every searchable meta for this surface: the scope's catalogue plus the
// per-document virtual rows.
const allMeta = computed(() => [...asCommandMetas(props.commands), ...asCommandMetas(props.extraRows)]);

function reasonFor(id = "") {
  return asReasonMap(props.availability)[id] ?? "";
}

function toRow(meta = allMeta.value[0], ranges = noRanges()) {
  return {
    id: meta.id,
    label: meta.label,
    path: meta.path,
    icon: meta.icon ?? "",
    ranges,
    shortcut: shortcutLabel(meta.id),
    reason: reasonFor(meta.id),
  };
}

// The empty-query list: recents first (most recent first, up to 8), then the
// common commands not already shown (AC-20, AC-21).
function emptyStateRows() {
  const byId = new Map(allMeta.value.map((meta) => [meta.id, meta]));
  const rows = [];
  const seen = new Set();
  for (const id of asIdList(props.recentIds)) {
    const meta = byId.get(id);
    if (meta && !seen.has(id)) {
      seen.add(id);
      rows.push(toRow(meta, []));
    }
  }
  for (const id of asIdList(props.commonIds)) {
    const meta = byId.get(id);
    if (meta && !seen.has(id)) {
      seen.add(id);
      rows.push(toRow(meta, []));
    }
  }
  return rows;
}

// The visible rows. Enabled rows keep their ranked order; disabled rows are
// demoted after enabled ones of the same score (stable), so the palette never
// puts an unusable row at the top.
const rows = computed(() => {
  if (query.value.trim() === "") {
    return emptyStateRows();
  }
  const matches = searchCommands(query.value, allMeta.value, asIdList(props.recentIds));
  const ranked = matches.map((match) => toRow(match.meta, match.ranges));
  const enabled = ranked.filter((row) => row.reason === "");
  const disabled = ranked.filter((row) => row.reason !== "");
  return [...enabled, ...disabled];
});

const resultCount = computed(() => rows.value.length);

// Splits a label into highlighted / plain segments for the matched ranges.
function segmentsFor(row = rows.value[0]) {
  if (row.ranges.length === 0) {
    return [{ text: row.label, hit: false }];
  }
  const segments = [];
  let cursor = 0;
  for (const [start, end] of row.ranges) {
    if (start > cursor) {
      segments.push({ text: row.label.slice(cursor, start), hit: false });
    }
    segments.push({ text: row.label.slice(start, end), hit: true });
    cursor = end;
  }
  if (cursor < row.label.length) {
    segments.push({ text: row.label.slice(cursor), hit: false });
  }
  return segments;
}

function optionId(index = 0) {
  return `${props.testid}-option-${index}`;
}

const activeDescendant = computed(() =>
  rows.value.length > 0 ? optionId(activeIndex.value) : undefined,
);

// Reset the highlight to the top whenever the result set changes.
watch(rows, () => {
  activeIndex.value = 0;
});

function clampActive() {
  if (activeIndex.value < 0) {
    activeIndex.value = rows.value.length - 1;
  } else if (activeIndex.value >= rows.value.length) {
    activeIndex.value = 0;
  }
}

async function scrollActiveIntoView() {
  await nextTick();
  const container = listbox.value;
  if (!(container instanceof HTMLElement)) {
    return;
  }
  const option = container.querySelector(`[id="${optionId(activeIndex.value)}"]`);
  if (option instanceof HTMLElement && typeof option.scrollIntoView === "function") {
    option.scrollIntoView({ block: "nearest" });
  }
}

function move(delta = 1) {
  if (rows.value.length === 0) {
    return;
  }
  activeIndex.value += delta;
  clampActive();
  void scrollActiveIntoView();
}

function run(index = 0) {
  const row = rows.value[index];
  if (!row || row.reason !== "") {
    // Disabled rows are listed but not runnable (AC-18).
    return;
  }
  emit("run", row.id);
}

function handleInputKeydown(event = new KeyboardEvent("keydown")) {
  switch (event.key) {
    case "ArrowDown":
      event.preventDefault();
      move(1);
      break;
    case "ArrowUp":
      event.preventDefault();
      move(-1);
      break;
    case "Home":
      event.preventDefault();
      activeIndex.value = 0;
      void scrollActiveIntoView();
      break;
    case "End":
      event.preventDefault();
      activeIndex.value = rows.value.length - 1;
      void scrollActiveIntoView();
      break;
    case "Enter":
      event.preventDefault();
      run(activeIndex.value);
      break;
    default:
      break;
  }
}
</script>

<template>
  <BaseDialog
    :title="props.title"
    :testid="props.testid"
    width="34rem"
    @dismiss="emit('dismiss')"
  >
    <div class="palette">
      <input
        v-model="query"
        class="palette-input"
        type="text"
        role="combobox"
        aria-autocomplete="list"
        aria-expanded="true"
        :aria-controls="`${props.testid}-listbox`"
        :aria-activedescendant="activeDescendant"
        :aria-label="props.title"
        :data-testid="`${props.testid}-input`"
        placeholder="Type a command or style…"
        autofocus
        @keydown="handleInputKeydown"
      >
      <!-- Result count for assistive tech (AC-30). -->
      <div
        class="visually-hidden"
        aria-live="polite"
        :data-testid="`${props.testid}-count`"
      >
        {{ resultCount }} {{ resultCount === 1 ? "result" : "results" }}
      </div>
      <ul
        :id="`${props.testid}-listbox`"
        ref="listbox"
        class="palette-list"
        role="listbox"
        :aria-label="props.title"
      >
        <li
          v-for="(row, index) in rows"
          :id="optionId(index)"
          :key="row.id"
          class="palette-option"
          :class="{ active: index === activeIndex, disabled: row.reason !== '' }"
          role="option"
          :aria-selected="index === activeIndex ? 'true' : 'false'"
          :aria-disabled="row.reason !== '' ? 'true' : undefined"
          :data-testid="`${props.testid}-option`"
          :data-command-id="row.id"
          @click="run(index)"
          @mousemove="activeIndex = index"
        >
          <span
            class="palette-icon"
            aria-hidden="true"
          >
            <AppIcon
              v-if="row.icon"
              :name="row.icon"
            />
          </span>
          <span class="palette-label">
            <span
              v-if="row.path.length"
              class="palette-path"
            >{{ row.path.join(" → ") }} → </span>
            <template
              v-for="(segment, si) in segmentsFor(row)"
              :key="si"
            ><mark
              v-if="segment.hit"
              class="palette-hit"
            >{{ segment.text }}</mark><span v-else>{{ segment.text }}</span></template>
          </span>
          <span
            v-if="row.reason"
            class="palette-reason"
          >{{ row.reason }}</span>
          <span
            v-else-if="row.shortcut"
            class="palette-shortcut"
          >{{ row.shortcut }}</span>
        </li>
        <li
          v-if="rows.length === 0"
          class="palette-empty"
          :data-testid="`${props.testid}-no-results`"
        >
          No matching commands
        </li>
      </ul>
    </div>
  </BaseDialog>
</template>

<style scoped>
.palette {
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
}

.palette-input {
  width: 100%;
  box-sizing: border-box;
  padding: var(--space-3) var(--space-4);
  font-size: var(--text-md);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  color: var(--text-primary);
}

.palette-input:focus {
  outline: var(--border-width) solid var(--accent);
  outline-offset: calc(-1 * var(--border-width));
}

.palette-list {
  list-style: none;
  margin: 0;
  padding: 0;
  max-height: 24rem;
  overflow-y: auto;
}

.palette-option {
  display: flex;
  align-items: center;
  gap: var(--space-3);
  padding: var(--space-2) var(--space-3);
  border-radius: var(--radius-sm);
  cursor: pointer;
  color: var(--text-primary);
}

.palette-option.active {
  background: var(--accent-subtle);
}

.palette-option.disabled {
  cursor: default;
  color: var(--text-secondary);
}

.palette-icon {
  display: inline-flex;
  width: 1.25rem;
  flex: none;
  color: var(--text-secondary);
}

.palette-label {
  flex: 1 1 auto;
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.palette-path {
  color: var(--text-secondary);
}

.palette-hit {
  background: transparent;
  color: var(--accent);
  font-weight: var(--weight-bold);
}

.palette-shortcut {
  flex: none;
  color: var(--text-secondary);
  font-size: var(--text-sm);
  font-variant-numeric: tabular-nums;
}

.palette-reason {
  flex: none;
  color: var(--text-secondary);
  font-size: var(--text-sm);
  font-style: italic;
}

.palette-empty {
  padding: var(--space-4);
  color: var(--text-secondary);
  text-align: center;
}

.visually-hidden {
  position: absolute;
  width: 1px;
  height: 1px;
  padding: 0;
  margin: calc(-1 * var(--border-width));
  overflow: hidden;
  clip: rect(0, 0, 0, 0);
  white-space: nowrap;
  border: 0;
}
</style>
