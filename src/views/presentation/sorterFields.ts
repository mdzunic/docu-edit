// The slide sorter's view-layer helpers (SPEC-099), lifted out of
// useSlideSorter so the controller stays within its composable budget: the
// typed pass-throughs the SlideSorter component reads its controller through
// (espree bars casts in .vue scripts), the rename input's focus ref, the
// event-target guards the sorter's window listeners share, and the honest
// wording of an import report. Nothing here holds state.
import type { SlideImportReport } from "../../lib/presentation/slideImport";
import type { SorterGroup } from "./useSlideSorter";

/** The import report as the sorter shows it. */
export interface SorterReport {
  title: string;
  lines: string[];
}

/** Typed pass-throughs for the sorter component's controller reads (espree
 * bars casts in .vue scripts). */
export function asSorterGroups(value: unknown): SorterGroup[] {
  return value as SorterGroup[];
}
export function asIndexes(value: unknown): number[] {
  return value as number[];
}

/** Focuses and selects a just-mounted rename input (a function ref). */
export function focusAndSelect(el: unknown): void {
  if (el instanceof HTMLInputElement && document.activeElement !== el) {
    el.focus();
    el.select();
  }
}

/** What the sorter says when a paste carries no slides (AC-9). */
export const NO_SLIDES_MESSAGE = "The clipboard doesn't hold any slides";

export function plural(n: number, noun: string): string {
  return `${n} ${noun}${n === 1 ? "" : "s"}`;
}

/** The honest summary of an import: what arrived (by title), what was reused,
 * and what could not come along. */
export function describeImport(report: SlideImportReport, titles: string[], verb: string): SorterReport {
  const named = titles.map((title) => `“${title}”`).join(", ");
  const lines = [`Kept ${plural(report.slides, "slide")}: ${named}.`];
  if (report.mediaReused > 0) {
    lines.push(`Reused ${plural(report.mediaReused, "image")} already in this deck.`);
  }
  if (report.rescaled) {
    lines.push("Resized to fit this deck's slide size.");
  }
  if (report.droppedFrames > 0) {
    lines.push(`Left out ${plural(report.droppedFrames, "object")} this deck can't hold: ${report.droppedFrameNames.join(", ")}.`);
  }
  if (report.lostTransitions > 0) {
    lines.push(`${plural(report.lostTransitions, "transition")} couldn't come along.`);
  }
  if (report.lostAnimations > 0) {
    lines.push(`${plural(report.lostAnimations, "animation")} couldn't come along.`);
  }
  if (report.droppedFrames + report.lostTransitions + report.lostAnimations === 0) {
    lines.push("Nothing was left out.");
  }
  return { title: `${verb} ${plural(report.slides, "slide")}`, lines };
}

export function isTypingTarget(target: EventTarget | null): boolean {
  return (
    target instanceof HTMLElement &&
    (target.tagName === "INPUT" || target.tagName === "TEXTAREA" || target.tagName === "SELECT" || target.isContentEditable)
  );
}

export function isInsideOverlay(target: EventTarget | null): boolean {
  return target instanceof Element && target.closest('[role="dialog"], [role="menu"]') !== null;
}
