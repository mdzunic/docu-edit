// The product-neutral history panel's typed transforms (SPEC-138), beside
// HistoryPanel.vue the way deckHistoryRender.ts sits beside the deck's: the
// component's script is linted by espree, so its runtime-declared
// `controller` arrives untyped and is narrowed here to the contract every
// history controller keeps — the deck's (useDeckHistory, through
// DeckHistoryPanel) and the workbook's (useWorkbookHistory).
import type { ComputedRef, Ref } from "vue";
import { formatTimestamp } from "../lib/format";
import type { VersionIndexEntry } from "../lib/core/versionChain";

/** What HistoryPanel reads and calls. A product's two facts — what a
 * version's counts say (`statLabel`) and what its download is called
 * (`downloadLabels`) — come from the controller. */
export interface HistoryPanelController {
  versions: Readonly<Ref<VersionIndexEntry[]>>;
  visibleVersions: Readonly<Ref<VersionIndexEntry[]>>;
  namedOnly: Ref<boolean>;
  pausedMessage: Readonly<Ref<string>>;
  previewModel: Readonly<Ref<unknown>>;
  previewVersionId: Readonly<Ref<string | null>>;
  preview(versionId: string): unknown;
  exitPreview(): void;
  restore(versionId: string): unknown;
  nameVersion(versionId: string, name: string): unknown;
  removeName(versionId: string): unknown;
  deleteVersion(versionId: string): unknown;
  downloadVersion(versionId: string): unknown;
  statLabel(entry: VersionIndexEntry): string;
  downloadLabels: Readonly<Ref<{ button: string; title: string }>>;
  /** True while this tab may not write (another tab holds the document):
   * Restore, naming and Delete are disabled. Absent: always writable. */
  readOnly?: Readonly<Ref<boolean>> | ComputedRef<boolean>;
  /** A sentence the panel shows about the last action (a version that could
   * not be opened). Absent or "": none. */
  notice?: Readonly<Ref<string>>;
}

export function asHistoryController(value: unknown): HistoryPanelController {
  return value as HistoryPanelController;
}

export function asEntries(value: unknown): VersionIndexEntry[] {
  return (value ?? []) as VersionIndexEntry[];
}

/** What took the snapshot, when it is worth saying. */
export function triggerLabel(entry: VersionIndexEntry): string {
  switch (entry.trigger) {
    case "pre-restore":
      return "Before restore";
    case "post-restore":
      return "After restore";
    default:
      return "";
  }
}

export function versionTime(entry: VersionIndexEntry): string {
  return formatTimestamp(new Date(entry.timestamp));
}
