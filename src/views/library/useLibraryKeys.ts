// The library's keyboard map (SPEC-067 AC-57..62), extracted from HomeView so
// the wiring hub stays within its line budget: arrows move the roving focus
// through the rows or cards, Enter opens, Space toggles the selection,
// Delete/Backspace trashes, F2 renames, Ctrl/Cmd+A selects everything the
// current filters show. Product-neutral — the actions it triggers are
// HomeView's to define.
import { nextTick } from "vue";
import type { LibraryController } from "../../composables/useLibrary";

export interface LibraryKeysDeps {
  lib: LibraryController;
  open: (id: string) => void;
  trash: (ids: string[]) => void;
  rename: (id: string) => void;
}

/** The content keydown handler for the list, grid and trash views. */
export function useLibraryKeys(deps: LibraryKeysDeps): (event: KeyboardEvent) => void {
  const { lib } = deps;

  async function focusDom(): Promise<void> {
    await nextTick();
    document.querySelector<HTMLElement>(`[data-doc-id="${lib.focusedId.value}"]`)?.focus();
  }

  return (event: KeyboardEvent): void => {
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "a") {
      event.preventDefault();
      lib.selectAll();
      return;
    }
    const key = event.key;
    const focused = lib.focusedId.value;
    if (key === "ArrowDown" || key === "ArrowRight") {
      event.preventDefault();
      lib.focusNext(1);
      void focusDom();
    } else if (key === "ArrowUp" || key === "ArrowLeft") {
      event.preventDefault();
      lib.focusNext(-1);
      void focusDom();
    } else if (key === "Enter") {
      if (focused) {
        event.preventDefault();
        deps.open(focused);
      }
    } else if (key === " ") {
      if (focused) {
        event.preventDefault();
        lib.toggleSelection(focused);
      }
    } else if (key === "Delete" || key === "Backspace") {
      if (!lib.showTrash.value && focused) {
        event.preventDefault();
        deps.trash(lib.hasSelection.value ? lib.selectedIds.value : [focused]);
      }
    } else if (key === "F2") {
      if (!lib.showTrash.value && focused) {
        event.preventDefault();
        deps.rename(focused);
      }
    }
  };
}
