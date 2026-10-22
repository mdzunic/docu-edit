// SPEC-135 "Round trip", across both formats on the spec135-annotations
// fixtures: a no-edit download is byte-identical part for part (equalBytes
// after unzip); add/edit/reply/resolve/link → save → reopen reproduces every
// modelled field (ids and retention excluded, equalStructural over the
// annotation arrays); writing the reopened model twice is byte-stable. The
// sanctioned deltas are pinned beside each format's wire tests:
// xlsx/annotations.spec.ts (a moved note's box keeps its size but not its
// position; an edited thread message loses its <mentions>) and
// ods/annotations.spec.ts (an ods cell that carried two links keeps the
// first once rewritten).
import { describe, expect, it } from "vitest";
import { equalBytes } from "../core/equality";
import { unzip } from "../core/zip";
import { equalStructural, firstStructuralDifference } from "../equality";
import { parseA1 } from "./address";
import { addNote, addThread, editNote, noteAt, replyToThread, setHyperlink, setNoteVisible, setThreadResolved, threadAt } from "./annotations";
import { buildSpec135AnnotationsOdsBytes, buildSpec135AnnotationsXlsxBytes } from "./fixtures/build";
import type { Sheet, SpreadsheetFormat, SpreadsheetModel } from "./model";
import { readSpreadsheetBytes } from "./open";
import { writeOds } from "./ods/writer";
import { writeXlsx } from "./xlsx/writer";

const FORMATS: readonly { format: Extract<SpreadsheetFormat, "xlsx" | "ods">; build: () => Promise<Uint8Array> }[] = [
  { format: "xlsx", build: buildSpec135AnnotationsXlsxBytes },
  { format: "ods", build: buildSpec135AnnotationsOdsBytes },
];

const write = (model: SpreadsheetModel): Promise<Uint8Array> => (model.format === "xlsx" ? writeXlsx(model) : writeOds(model));
const open = (bytes: Uint8Array, format: SpreadsheetFormat): Promise<SpreadsheetModel> => readSpreadsheetBytes(bytes, format, "SPEC-135");
const at = (a1: string) => parseA1(a1)!;

/** A sheet's annotations field for field — ids, persons' ids and retention left out. */
function fields(sheet: Sheet): unknown {
  return {
    notes: (sheet.notes ?? []).map(({ noteId: _id, source: _source, ...note }) => note),
    threads: (sheet.threads ?? []).map(({ threadId: _id, messages, ...thread }) => ({
      ...thread,
      messages: messages.map(({ messageId: _message, personId: _person, source: _source, ...message }) => message),
    })),
    hyperlinks: (sheet.hyperlinks ?? []).map(({ linkId: _id, source: _source, ...link }) => link),
  };
}

/** The names of the parts whose bytes differ (or exist on one side only). */
function changedParts(a: Map<string, Uint8Array>, b: Map<string, Uint8Array>): string[] {
  const names = new Set([...a.keys(), ...b.keys()]);
  return [...names].filter((name) => !a.has(name) || !b.has(name) || !equalBytes(a.get(name)!, b.get(name)!)).toSorted();
}

/** Every annotation command the round trip exercises, on Summary (xlsx
 * threads too — an .ods workbook has notes only). */
function edit(model: SpreadsheetModel): void {
  const summary = model.sheets[0]!;
  expect(addNote(model, summary.id, at("D4"), "Added in the app", "You")).not.toBeNull();
  expect(editNote(model, summary.id, noteAt(model.sheets[0]!, at("B2"))!.noteId, "Edited text")).not.toBeNull();
  expect(setNoteVisible(model, summary.id, noteAt(model.sheets[0]!, at("B4"))!.noteId, false)).toBe(true);
  // A2 holds text in both fixtures (an ods link wraps its cell's text).
  setHyperlink(model, summary.id, { top: 1, left: 0, bottom: 1, right: 0 }, { kind: "place", location: "Data!A10" }, "Go to the target");
  if (model.format === "xlsx") {
    const thread = threadAt(model.sheets[0]!, at("D2"))!;
    expect(replyToThread(model, summary.id, thread.threadId, "Replying as You", "You")).not.toBeNull();
    expect(setThreadResolved(model, summary.id, thread.threadId, true)).toBe(true);
    expect(addThread(model, summary.id, at("F3"), "A new thread", "You")).not.toBeNull();
  }
}

describe.each(FORMATS)("SPEC-135 annotations round trip ($format)", ({ format, build }) => {
  it("a no-edit download is byte-identical part for part", async () => {
    const bytes = await build();
    const model = await open(bytes, format);
    expect(model.sheets[0]!.notes?.length).toBeGreaterThan(0);
    expect(changedParts(await unzip(bytes), await unzip(await write(model)))).toEqual([]);
  });

  it("add/edit/reply/resolve/link → save → reopen reproduces every modelled field, and writing the reopened model twice is byte-stable", async () => {
    const model = await open(await build(), format);
    edit(model);
    const reopened = await open(await write(model), format);
    for (const [index, sheet] of model.sheets.entries()) {
      const want = fields(sheet);
      const got = fields(reopened.sheets[index]!);
      expect(equalStructural(got, want), firstStructuralDifference(got, want) ?? "").toBe(true);
      expect(got).toEqual(want);
    }
    if (format === "xlsx") {
      const names = (m: SpreadsheetModel) => (m.persons ?? []).map((person) => person.displayName).toSorted();
      expect(names(reopened)).toEqual(names(model));
    }
    const once = await unzip(await write(reopened));
    const twice = await unzip(await write(reopened));
    expect(changedParts(once, twice)).toEqual([]);
  });
});
