// SPEC-031 — footnotes and endnotes through the .odt reader and writer:
// text:note with text:note-citation and text:note-body inline at the
// reference, and text:notes-configuration numbering settings in styles.xml.
import { describe, it, expect } from "vitest";
import { parseOdt, type ParsedOdtDocument } from "./reader";
import { writeOdt } from "./writer";
import { unzip } from "../core/zip";
import {
  createDocument,
  createNoteRun,
  type DocumentModel,
  type NoteContent,
  type NoteKind,
  type Paragraph,
} from "../documentModel";
import { collectNoteRefs, noteById } from "../notes";
import { blockText } from "../textEditing";

const decoder = new TextDecoder();

// An .odt round trip resolves the writer's default-style font onto every
// run when reading back (same as the docx spec's constant).
const RESOLVED_DEFAULT_FORMATTING = { fontFamily: "Calibri", fontSizePt: 11 };

function paragraph(text: string): Paragraph {
  return {
    type: "paragraph",
    runs: text.length > 0 ? [{ text, formatting: { ...RESOLVED_DEFAULT_FORMATTING } }] : [],
    formatting: {},
  };
}

function note(noteId: string, kind: NoteKind, text: string): NoteContent {
  return { noteId, kind, blocks: [paragraph(text)] };
}

function modelWithNotes(notes: NoteContent[]): DocumentModel {
  const model = createDocument("spec-031-odt");
  const runs = notes.flatMap((entry, index) => [
    { text: `Body ${index} `, formatting: { ...RESOLVED_DEFAULT_FORMATTING } },
    createNoteRun(entry.noteId, entry.kind, { ...RESOLVED_DEFAULT_FORMATTING }),
  ]);
  model.sections = [{ blocks: [{ type: "paragraph", runs, formatting: {} }] }];
  model.notes = notes;
  return model;
}

async function writeAndReadBack(model: DocumentModel): Promise<{
  parsed: ParsedOdtDocument;
  contentXml: string;
  stylesXml: string;
}> {
  const { bytes } = await writeOdt(model);
  const parts = await unzip(bytes);
  const parsed = await parseOdt(bytes);
  return {
    parsed,
    contentXml: decoder.decode(parts.get("content.xml")!),
    stylesXml: decoder.decode(parts.get("styles.xml")!),
  };
}

function parsedModel(parsed: ParsedOdtDocument): DocumentModel {
  const model = createDocument("spec-031-odt-reparsed");
  model.sections = parsed.sections;
  model.notes = parsed.notes;
  return model;
}

describe("footnote round trip (SPEC-031 AC-23)", () => {
  it("writes text:note with citation and body and reads it back", async () => {
    const model = modelWithNotes([note("n1", "footnote", "First note"), note("n2", "footnote", "Second note")]);
    const { parsed, contentXml } = await writeAndReadBack(model);

    expect(contentXml).toContain('text:note-class="footnote"');
    expect(contentXml).toContain("<text:note-citation>1</text:note-citation>");
    expect(contentXml).toContain("<text:note-citation>2</text:note-citation>");
    expect(contentXml).toContain("<text:note-body>");

    const reparsed = parsedModel(parsed);
    const refs = collectNoteRefs(reparsed, "footnote");
    expect(refs).toHaveLength(2);
    expect(blockText(noteById(reparsed, refs[0].noteId)!.blocks[0])).toBe("First note");
    expect(blockText(noteById(reparsed, refs[1].noteId)!.blocks[0])).toBe("Second note");
    const body = reparsed.sections[0].blocks[0] as Paragraph;
    const marks = body.runs.filter((run) => run.note);
    expect(marks).toHaveLength(2);
    expect(marks[0].formatting.vertAlign).toBe("superscript");
  });

  it("keeps note ids stable across a round trip", async () => {
    const model = modelWithNotes([note("note-abc123", "footnote", "Stable")]);
    const { parsed } = await writeAndReadBack(model);
    const refs = collectNoteRefs(parsedModel(parsed), "footnote");
    expect(refs[0].noteId).toBe("note-abc123");
    expect(parsed.notes[0].noteId).toBe("note-abc123");
  });

  it("round-trips multi-paragraph note content", async () => {
    const content: NoteContent = {
      noteId: "n1",
      kind: "footnote",
      blocks: [paragraph("First paragraph"), paragraph("Second paragraph")],
    };
    const model = modelWithNotes([content]);
    const { parsed } = await writeAndReadBack(model);
    const readBack = parsed.notes.find((entry) => entry.noteId === "n1")!;
    expect(readBack.blocks).toHaveLength(2);
    expect(blockText(readBack.blocks[1])).toBe("Second paragraph");
  });
});

describe("endnote round trip (SPEC-031 AC-24)", () => {
  it("writes text:note-class endnote and reads it back", async () => {
    const model = modelWithNotes([note("e1", "endnote", "The endnote")]);
    const { parsed, contentXml } = await writeAndReadBack(model);
    expect(contentXml).toContain('text:note-class="endnote"');
    const reparsed = parsedModel(parsed);
    const refs = collectNoteRefs(reparsed, "endnote");
    expect(refs).toHaveLength(1);
    expect(blockText(noteById(reparsed, refs[0].noteId)!.blocks[0])).toBe("The endnote");
  });

  it("keeps both streams separate, with roman endnote citations", async () => {
    const model = modelWithNotes([note("f1", "footnote", "Foot"), note("e1", "endnote", "End")]);
    const { parsed, contentXml } = await writeAndReadBack(model);
    expect(contentXml).toContain("<text:note-citation>i</text:note-citation>");
    const reparsed = parsedModel(parsed);
    expect(collectNoteRefs(reparsed, "footnote")).toHaveLength(1);
    expect(collectNoteRefs(reparsed, "endnote")).toHaveLength(1);
  });
});

describe("note settings round trip (SPEC-031 AC-17..20)", () => {
  it("writes text:notes-configuration and reads it back", async () => {
    const model = modelWithNotes([note("n1", "footnote", "Note")]);
    model.footnoteSettings = { numberFormat: "lowerRoman", startAt: 3, restart: "page" };
    const { parsed, stylesXml } = await writeAndReadBack(model);
    expect(stylesXml).toContain('text:note-class="footnote"');
    expect(stylesXml).toContain('style:num-format="i"');
    expect(stylesXml).toContain('text:start-value="2"');
    expect(stylesXml).toContain('text:start-numbering-at="page"');
    expect(parsed.footnoteSettings).toEqual({ numberFormat: "lowerRoman", startAt: 3, restart: "page" });
  });

  it("maps a section restart to chapter and back", async () => {
    const model = modelWithNotes([note("n1", "footnote", "Note")]);
    model.footnoteSettings = { numberFormat: "lowerAlpha", restart: "section" };
    const { parsed, stylesXml } = await writeAndReadBack(model);
    expect(stylesXml).toContain('text:start-numbering-at="chapter"');
    expect(parsed.footnoteSettings).toEqual({ numberFormat: "lowerAlpha", restart: "section" });
  });

  it("emits no note markup for a document without notes", async () => {
    const model = createDocument("spec-031-odt-none");
    model.sections = [{ blocks: [paragraph("Plain")] }];
    const { parsed, contentXml, stylesXml } = await writeAndReadBack(model);
    expect(contentXml).not.toContain("text:note");
    expect(stylesXml).not.toContain("notes-configuration");
    expect(parsed.notes).toEqual([]);
    expect(parsed.footnoteSettings).toBeUndefined();
  });
});
