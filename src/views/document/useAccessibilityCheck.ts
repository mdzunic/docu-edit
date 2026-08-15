// The accessibility checker (SPEC-043): the results pane's open state and
// issue list, the whole-document scan chunked through `scheduleIdleWork`
// (AC "runs on a 200-page document in under 3 seconds without blocking the
// UI... chunked through scheduleIdleWork, as spell check and find already
// are"), navigation to an issue's offending content, and the three
// mechanical fixes. Mirrors useTrackChanges.ts's shape (a pane's open
// state, a derived/scanned list, goToX navigation over applyCaretMove +
// scrollModelCaretIntoView).
import { ref, shallowRef, type Ref } from "vue";
import {
  buildAccessibilityScanTargets,
  createAccessibilityScanState,
  finalizeAccessibilityScan,
  scanAccessibilityTarget,
  sortAccessibilityIssues,
  type AccessibilityFixKind,
  type AccessibilityIssue,
  type AccessibilityIssueKind,
  type AccessibilityScanState,
  type AccessibilityScanTarget,
} from "../../lib/accessibilityCheck";
import { setRowIsHeader, type TableLocation } from "../../lib/tables";
import { applyParagraphStyle, caretSelection } from "../../lib/commands";
import { headingStyleId } from "../../lib/styles";
import { noIdleHandle, scheduleIdleWork, type IdleHandle } from "../../lib/core/idleScheduler";
import { compareCaret } from "../../lib/textEditing";
import type { ImageLocation } from "../../lib/images";
import type { EditorContext } from "./context";

/** Typed pass-through for an `AccessibilityIssue[]` arriving at
 * AccessibilityPane.vue as a runtime-declared Array prop (espree bars the
 * annotation there — see useTrackChanges.ts's asChangeEntries). */
export function asAccessibilityIssues(value: unknown): AccessibilityIssue[] {
  return value as AccessibilityIssue[];
}

/** Display labels for AccessibilityPane.vue's issue cards — kept as typed
 * lookup functions (rather than a plain object literal indexed by a bare
 * `string`) because that .vue script is espree-parsed and cannot itself
 * write the `Record<AccessibilityIssueKind, string>` annotation an object
 * literal would need to index safely. */
const KIND_LABELS: Record<AccessibilityIssueKind, string> = {
  "missing-alt-text": "Image missing alt text",
  "table-no-header-row": "Table has no header row",
  "table-merged-cells": "Table has merged cells",
  "heading-level-skipped": "Heading level skipped",
  "empty-heading": "Empty heading",
  "link-bare-url": "Link text is a bare URL",
  "link-click-here": 'Link text is "click here"',
  "low-contrast": "Low contrast text",
  "colour-alone": "Meaning carried by colour alone",
  "no-heading-structure": "No heading structure",
  "visual-heading": "Looks like a heading",
};

const FIX_LABELS: Record<AccessibilityFixKind, string> = {
  "add-alt-text": "Add alt text",
  "mark-header-row": "Mark header row",
  "fix-heading-level": "Fix heading level",
};

export function accessibilityKindLabel(kind: AccessibilityIssueKind): string {
  return KIND_LABELS[kind];
}

export function accessibilityFixLabel(fix: AccessibilityFixKind): string {
  return FIX_LABELS[fix];
}

// Same idle-slot budget spell check and find use (SPEC-035/SPEC-017): a
// short timeout so a busy main thread still gets a slot soon, a small
// per-slot budget so a large pass never eats a whole frame.
const IDLE_OPTIONS = { timeoutMs: 150, fallbackBudgetMs: 8, maxBudgetMs: 16 };

export interface AccessibilityCheckDeps extends EditorContext {
  /** The image toolbar's own selection/dialog (SPEC-024), reused so
   * "Add alt text" opens the exact same options dialog the image toolbar
   * does rather than inventing a second alt-text editing surface. */
  imageSelection: Ref<ImageLocation | null>;
  openImageOptions: () => void;
}

export function useAccessibilityCheck(deps: AccessibilityCheckDeps) {
  const { doc, applyEdit, applyCaretMove, scrollModelCaretIntoView, imageSelection, openImageOptions, onSessionClose } =
    deps;

  const accessibilityPaneOpen = ref(false);
  const accessibilityIssues = shallowRef<AccessibilityIssue[]>([]);
  const accessibilityScanning = ref(false);

  let idleHandle: IdleHandle | null = noIdleHandle();
  let scanTargets: AccessibilityScanTarget[] = [];
  let scanCursor = 0;
  let scanState: AccessibilityScanState = createAccessibilityScanState();
  let scanIssues: AccessibilityIssue[] = [];

  function cancelScan(): void {
    idleHandle?.cancel();
    idleHandle = null;
    accessibilityScanning.value = false;
  }

  function runSlot(): void {
    idleHandle = scheduleIdleWork((budgetMs) => {
      idleHandle = null;
      const model = doc.value;
      if (!model) {
        accessibilityScanning.value = false;
        return;
      }
      const started = performance.now();
      const startCursor = scanCursor;
      // The budget break holds regardless of progress, matching
      // usePagination's measureLoop: at least one target advances per slot,
      // but a slot never runs past its budget once it has advanced one.
      while (scanCursor < scanTargets.length) {
        if (scanCursor > startCursor && performance.now() - started > budgetMs) {
          break;
        }
        scanAccessibilityTarget(model, scanTargets[scanCursor], scanState, scanIssues);
        scanCursor += 1;
      }
      if (scanCursor >= scanTargets.length) {
        finalizeAccessibilityScan(scanState, scanIssues);
        accessibilityIssues.value = sortAccessibilityIssues(scanIssues);
        accessibilityScanning.value = false;
        return;
      }
      runSlot();
    }, IDLE_OPTIONS);
  }

  /** Starts a fresh whole-document scan (AC-1's "opens a results pane" and
   * AC-28's "applying a fix re-runs the check... and updates the list" both
   * drive through this — a full rescan is simple and always correct where a
   * scoped re-check of just the fixed item would have to reason about
   * cross-item state like the running heading level). */
  function rescan(): void {
    cancelScan();
    const model = doc.value;
    if (!model) {
      accessibilityIssues.value = [];
      return;
    }
    scanTargets = buildAccessibilityScanTargets(model);
    scanCursor = 0;
    scanState = createAccessibilityScanState();
    scanIssues = [];
    accessibilityScanning.value = true;
    runSlot();
  }

  function openAccessibilityPane(): void {
    accessibilityPaneOpen.value = true;
    rescan();
  }

  function closeAccessibilityPane(): void {
    accessibilityPaneOpen.value = false;
    cancelScan();
  }

  /** Selecting an issue (AC-25): scrolls to and highlights (selects) its
   * offending content — a collapsed caret at `at` alone would only point at
   * it, not highlight it. */
  async function goToAccessibilityIssue(issueId: string): Promise<void> {
    const issue = accessibilityIssues.value.find((entry) => entry.id === issueId);
    if (!issue) {
      return;
    }
    applyCaretMove(issue.at);
    if (compareCaret(issue.endAt, issue.at) !== 0) {
      applyCaretMove(issue.endAt, true);
    }
    await scrollModelCaretIntoView(issue.at, true);
  }

  /** Applies an issue's mechanical fix (AC-27), then re-runs the check
   * (AC-28). "Add alt text" instead opens the image's own options dialog —
   * the alt text's *words* are inherently a human's call, so the mechanical
   * part is jumping straight to the right field; the dialog's own Apply
   * handler (DocumentView's handleImageOptionsApply) commits the edit and
   * triggers the pane's rescan itself once the user actually types
   * something and applies it. */
  function applyAccessibilityFix(issueId: string): void {
    const model = doc.value;
    const issue = accessibilityIssues.value.find((entry) => entry.id === issueId);
    if (!model || !issue || !issue.fix) {
      return;
    }
    if (issue.fix === "add-alt-text") {
      imageSelection.value = { section: issue.at.section, block: issue.at.block };
      openImageOptions();
      return;
    }
    if (issue.fix === "mark-header-row") {
      const location: TableLocation = { section: issue.at.section, block: issue.at.block, row: 0, col: 0 };
      const selection = caretSelection(issue.at);
      const result = setRowIsHeader(model, selection, location, 0, true);
      applyEdit(result, selection, "mark header row");
      rescan();
      return;
    }
    if (issue.fix === "fix-heading-level" && issue.fixHeadingLevel !== undefined) {
      const selection = caretSelection(issue.at);
      const result = applyParagraphStyle(model, selection, headingStyleId(issue.fixHeadingLevel));
      applyEdit(result, selection, "fix heading level");
      rescan();
    }
  }

  onSessionClose(() => {
    cancelScan();
    accessibilityPaneOpen.value = false;
    accessibilityIssues.value = [];
  });

  return {
    accessibilityPaneOpen,
    accessibilityIssues,
    accessibilityScanning,
    openAccessibilityPane,
    closeAccessibilityPane,
    goToAccessibilityIssue,
    applyAccessibilityFix,
    rescanAccessibility: rescan,
  };
}

export type UseAccessibilityCheckReturn = ReturnType<typeof useAccessibilityCheck>;
