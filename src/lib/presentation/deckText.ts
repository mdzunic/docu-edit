// A deck's text for library search (SPEC-114): per slide, in deck order, the
// title first, then every other frame's paragraphs in z-order — text and
// shape frames, a table's cells row by row, a group's children — then the
// slide's SPEC-083 speaker notes. One line per paragraph. Images, media,
// charts and foreign frames carry no text this layer reads.
//
// Each slide's lines are recorded as a span, so an offset into the text
// names the slide it came from — how a search result opens the deck AT the
// slide holding its match. Layering rule (boundary.spec.ts): presentation
// siblings and core only.
import type { ExtractedText } from "../core/extractedText";
import type { PresentationModel, SlideFrame } from "./model";
import { paragraphText } from "./slideModel";

/** The text one slide contributed: [start, end) in the extracted text. */
export interface DeckTextSpan {
  start: number;
  end: number;
  slideIndex: number;
}

export interface DeckText extends ExtractedText {
  spans: DeckTextSpan[];
}

function frameLines(frame: SlideFrame): string[] {
  if (frame.kind === "text" || frame.kind === "shape") {
    return (frame.paragraphs ?? []).map(paragraphText);
  }
  if (frame.kind === "table") {
    return frame.rows.flatMap((row) => row.cells.filter((cell) => !cell.covered).flatMap((cell) => cell.paragraphs.map(paragraphText)));
  }
  if (frame.kind === "group") {
    return frame.children.flatMap(frameLines);
  }
  return [];
}

/** The deck's text, slide by slide, with each slide's span. A deck model is
 * read whole: its coverage is always `full`. */
export function deckText(model: PresentationModel): DeckText {
  const parts: string[] = [];
  const spans: DeckTextSpan[] = [];
  let length = 0;
  model.slides.forEach((slide, slideIndex) => {
    const title = slide.frames.find((frame) => frame.kind === "text" && frame.role === "title");
    const lines = [
      ...(title ? frameLines(title) : []),
      ...slide.frames.filter((frame) => frame !== title).flatMap(frameLines),
      ...(slide.notes ?? []),
    ].filter((line) => line.trim().length > 0);
    if (lines.length === 0) {
      return;
    }
    const text = lines.join("\n");
    const start = parts.length === 0 ? 0 : length + 1;
    parts.push(text);
    length = start + text.length;
    spans.push({ start, end: length, slideIndex });
  });
  return { text: parts.join("\n"), coverage: "full", spans };
}

/** The slide an offset into `extraction` came from; -1 for none. */
export function deckTextSlide(extraction: DeckText, offset: number): number {
  const span = extraction.spans.find((candidate) => offset >= candidate.start && offset <= candidate.end);
  return span ? span.slideIndex : -1;
}
