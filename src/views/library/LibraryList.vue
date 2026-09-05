<script setup lang="ts">
// The library's list rendering (SPEC-067 AC-1): one row per visible entry —
// select checkbox, star, highlighted title, relative time, format badge and
// the overflow menu. Row semantics (modifier-click selects, plain click
// opens) live here; opening, menus and keyboard handling stay with
// HomeView, which owns the cross-view state they touch.
import HighlightedTitle from "../../components/library/HighlightedTitle.vue";
import RowMenu from "../../components/library/RowMenu.vue";
import { formatBadgeLabel, formatFileSize, formatRelativeTime } from "../../lib/format";
import { libraryRowMenuItems } from "../../lib/libraryMenu";

const props = defineProps({
  lib: { type: Object, required: true },
  menuFor: { type: String, default: "" },
  highlightId: { type: String, default: "" },
});

const emit = defineEmits(["open", "select-modifiers", "keydown", "menu-toggle", "menu-close", "menu-select"]);

function onRowClick(id = "", event = new MouseEvent("click")) {
  if (event.shiftKey || event.ctrlKey || event.metaKey) {
    event.preventDefault();
    emit("select-modifiers", {
      id,
      shift: event.shiftKey,
      ctrlOrMeta: event.ctrlKey || event.metaKey,
    });
    return;
  }
  emit("open", id);
}
</script>

<template>
  <table
    class="library-table"
    role="grid"
    data-testid="library-list"
    @keydown="emit('keydown', $event)"
  >
    <thead>
      <tr role="row">
        <th
          scope="col"
          role="columnheader"
          class="col-select"
        >
          <span class="visually-hidden">Select</span>
        </th>
        <th
          scope="col"
          role="columnheader"
          class="col-star"
        >
          <span class="visually-hidden">Starred</span>
        </th>
        <th
          scope="col"
          role="columnheader"
        >
          Title
        </th>
        <th
          scope="col"
          role="columnheader"
        >
          Last modified
        </th>
        <th
          scope="col"
          role="columnheader"
        >
          Format
        </th>
        <th
          scope="col"
          role="columnheader"
        >
          Size
        </th>
        <th
          scope="col"
          role="columnheader"
        >
          <span class="visually-hidden">Actions</span>
        </th>
      </tr>
    </thead>
    <tbody>
      <tr
        v-for="(entry, index) in props.lib.visible.value"
        :key="entry.id"
        class="library-row"
        role="row"
        data-testid="library-row"
        :data-doc-id="entry.id"
        :class="{ 'row-highlight': entry.id === props.highlightId, 'is-selected': props.lib.isSelected(entry.id) }"
        :tabindex="props.lib.rowTabindex(entry.id, index)"
        :aria-selected="props.lib.isSelected(entry.id)"
        @click="onRowClick(entry.id, $event)"
        @focus="props.lib.setFocusedId(entry.id)"
      >
        <td
          class="cell-select"
          role="gridcell"
          @click.stop
        >
          <input
            type="checkbox"
            data-testid="row-select"
            :checked="props.lib.isSelected(entry.id)"
            :aria-label="`Select ${entry.title}`"
            @change="props.lib.toggleSelection(entry.id)"
          >
        </td>
        <td
          class="cell-star"
          role="gridcell"
          @click.stop
        >
          <button
            type="button"
            class="star-button"
            data-testid="row-star"
            :class="{ 'is-starred': entry.starred }"
            :aria-pressed="entry.starred"
            :aria-label="entry.starred ? `Unstar ${entry.title}` : `Star ${entry.title}`"
            @click="props.lib.toggleStar(entry.id)"
          >
            {{ entry.starred ? "★" : "☆" }}
          </button>
        </td>
        <td
          class="cell-title"
          role="gridcell"
          data-testid="row-title"
        >
          <HighlightedTitle
            :title="entry.title"
            :search="props.lib.search.value"
          />
        </td>
        <td
          class="cell-modified"
          role="gridcell"
          data-testid="row-modified"
        >
          {{ formatRelativeTime(entry.lastModified, props.lib.now.value) }}
        </td>
        <td
          class="cell-format"
          role="gridcell"
        >
          <span
            class="badge"
            data-testid="row-format"
          >{{ formatBadgeLabel(entry.format) }}</span>
        </td>
        <td
          class="cell-size"
          role="gridcell"
          data-testid="row-size"
        >
          {{ formatFileSize(props.lib.sizeOf(entry.id)) }}
        </td>
        <td
          class="cell-actions"
          role="gridcell"
          @click.stop
        >
          <RowMenu
            :open="props.menuFor === entry.id"
            :label="`Actions for ${entry.title}`"
            :items="libraryRowMenuItems(entry.starred)"
            @toggle="emit('menu-toggle', entry.id)"
            @close="emit('menu-close')"
            @select="(action) => emit('menu-select', { id: entry.id, action })"
          />
        </td>
      </tr>
    </tbody>
  </table>
</template>

<style scoped>
.library-table {
  width: 100%;
  border-collapse: collapse;
  background: var(--surface-base);
  border-radius: var(--radius-md);
  box-shadow: var(--shadow-1);
}

/* The handoff's column headers: small uppercase caps over a heavier rule. */
.library-table th {
  text-align: left;
  font-size: var(--text-xs);
  font-weight: var(--weight-medium);
  letter-spacing: 0.08em;
  text-transform: uppercase;
  color: var(--text-secondary);
  padding: var(--space-3) var(--space-5);
  border-bottom: calc(2 * var(--border-width)) solid var(--border-default);
}

.library-table td {
  padding: var(--space-3) var(--space-5);
  border-bottom: var(--border-width) solid var(--border-subtle);
  color: var(--text-primary);
}

.col-select,
.col-star,
.cell-select,
.cell-star {
  width: 2.5rem;
  padding-right: 0;
}

.library-row {
  cursor: pointer;
}

.library-row:hover {
  background: var(--surface-raised);
}

.library-row:focus-visible {
  outline: var(--focus-ring-width, 2px) solid var(--border-focus, var(--text-accent));
  outline-offset: -2px;
}

.library-row.is-selected {
  background: var(--surface-selected);
}

.row-highlight {
  animation: row-flash var(--motion-base) var(--easing);
}

@keyframes row-flash {
  from {
    background: var(--surface-selected);
  }
  to {
    background: transparent;
  }
}

@media (prefers-reduced-motion: reduce) {
  .row-highlight {
    animation: none;
  }
}

.star-button {
  font-size: var(--text-lg);
  line-height: 1;
  border: none;
  background: transparent;
  cursor: pointer;
  color: var(--text-secondary);
}

.star-button.is-starred {
  color: var(--text-accent);
}

.cell-title {
  font-weight: var(--weight-medium);
  overflow-wrap: anywhere;
}

.cell-modified {
  color: var(--text-secondary);
  white-space: nowrap;
}

.cell-actions {
  position: relative;
  text-align: right;
  white-space: nowrap;
}

.badge {
  display: inline-block;
  padding: var(--space-1) var(--space-3);
  border-radius: var(--radius-pill);
  background: var(--surface-selected);
  color: var(--text-accent);
  font-size: var(--text-sm);
  font-weight: var(--weight-bold);
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

/* Compact tier (SPEC-068). Seven columns cannot share a phone's width: the
   table's own minimum was ~520px inside a ~340px column, so it overflowed a
   parent that does not scroll — the row menu sat off-screen entirely, and the
   title column, squeezed to ~66px, broke mid-word ("Untit / led / doc").
   LibraryGrid already had this breakpoint; the list never got one.

   So each row becomes two lines: the things you act on (select, star, title,
   menu) across the top, the metadata beneath. The markup is untouched — this
   is the same grid of the same cells, re-flowed — which is why every cell
   carries an explicit role above: blockifying a `td` drops the implicit table
   semantics that `role="grid"` needs, and the roles have to outlive the
   display change. */
@media (max-width: 639.98px) {
  .library-table,
  .library-table tbody {
    display: block;
  }

  /* The column headers name columns that no longer exist side by side. */
  .library-table thead {
    display: none;
  }

  .library-row {
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    gap: var(--space-2);
    padding: var(--space-3) var(--space-4);
    border-bottom: var(--border-width) solid var(--border-subtle);
  }

  /* Cells draw the row's border now, not their own. */
  .library-table td {
    padding: 0;
    border-bottom: none;
  }

  .cell-select,
  .cell-star {
    order: 1;
    width: auto;
    flex: none;
  }

  .cell-star {
    order: 2;
  }

  /* Grows to eat the first line, which is what pushes the metadata onto the
     second — `flex-basis: 0` so wrapping is decided before it grows. */
  .cell-title {
    order: 3;
    flex: 1 1 0;
    min-width: 0;
  }

  .cell-actions {
    order: 4;
    flex: none;
  }

  /* A zero-height flex item spanning the full width: the line break. Flexbox
     has no break-before, and the metadata cells' own widths are small enough
     that they would otherwise ride up onto the first line. */
  .library-row::after {
    content: "";
    order: 5;
    flex-basis: 100%;
    height: 0;
  }

  .cell-format {
    order: 6;
  }

  .cell-modified {
    order: 7;
  }

  .cell-size {
    order: 8;
    color: var(--text-secondary);
  }

  .cell-format,
  .cell-modified,
  .cell-size {
    flex: none;
    font-size: var(--text-sm);
  }
}
</style>
