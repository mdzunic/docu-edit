// The one rule that keeps document text serialisable: every character this
// model holds must be legal in XML 1.0 character data.
//
// XML 1.0 does not merely require the five predefined entities to be escaped
// — it forbids most C0 control characters outright (§2.2 Char), and forbids
// them *even as numeric character references*, so there is no escape hatch on
// the write side. `escapeXml` therefore cannot fix this: it passes a U+0001
// straight through, the writer emits it into `word/document.xml`, and the
// package the user just saved is not well-formed XML. Word refuses it, and so
// does this app's own reader — with "This file isn't a Word document we can
// open", blaming the file for what the save did to it.
//
// Where illegal scalars actually come from, and therefore where this is
// applied: text typed or pasted into the canvas (lib/textEditing's insertText,
// the choke point every insertion path routes through), and `.md` files, whose
// reader parses plain text rather than XML. The XML-backed readers cannot
// carry one at all — DOMParser rejects the document, so a .docx/.odt/.pptx
// that contains one never parses in the first place.
//
// U+000B is the case that matters in practice, not a curiosity: it is what
// Word puts in the clipboard's plain-text flavour for a Shift+Enter line
// break, so pasting from Word into this editor was enough to produce an
// unopenable save.
//
// Every pattern below is written with \u escapes on purpose: these characters
// are invisible in a source file, and a literal one in a character class is
// unreviewable and easily mangled by an editor or a copy-paste.

/** U+000B VERTICAL TAB and U+000C FORM FEED both mean "break the line here" —
 * Word writes U+000B for a Shift+Enter break in its clipboard's plain-text
 * flavour — and this model spells a line break as a literal "\n" inside run
 * text (see Run's own doc comment). They are translated, not dropped. */
const LINE_BREAK_SCALARS = "\\u000B\\u000C";

/** Everything else XML 1.0 forbids in character data: the C0 controls other
 * than tab (U+0009), line feed (U+000A) and carriage return (U+000D), plus
 * the permanently-unassigned U+FFFE/U+FFFF. Dropped — they carry no meaning a
 * document can render, and there is nothing to translate them to. */
const ILLEGAL_SCALARS = "\\u0000-\\u0008\\u000E-\\u001F\\uFFFE\\uFFFF";

/** An unpaired surrogate is not a character at all: it cannot be encoded as
 * UTF-8, so it breaks serialisation before XML validity even applies. A high
 * surrogate not followed by a low one, or a low one not preceded by a high
 * one. */
const LONE_SURROGATE_SOURCE =
  "[\\uD800-\\uDBFF](?![\\uDC00-\\uDFFF])|(?<![\\uD800-\\uDBFF])[\\uDC00-\\uDFFF]";

// Built once. The `test` copies are deliberately non-global: a /g regex
// carries lastIndex between calls, which makes `test` alternate true/false on
// the same input.
const NEEDS_FIXING = new RegExp(
  `[${LINE_BREAK_SCALARS}${ILLEGAL_SCALARS}]|${LONE_SURROGATE_SOURCE}`,
);
const LINE_BREAK_RE = new RegExp(`[${LINE_BREAK_SCALARS}]`, "g");
const ILLEGAL_RE = new RegExp(`[${ILLEGAL_SCALARS}]`, "g");
const LONE_SURROGATE_RE = new RegExp(LONE_SURROGATE_SOURCE, "g");

/** Whether `text` contains anything `sanitizeXmlText` would change. Lets a
 * caller skip an allocation on the overwhelmingly common clean string. */
export function hasIllegalXmlScalars(text: string): boolean {
  return NEEDS_FIXING.test(text);
}

/** `text` with every scalar XML 1.0 cannot carry removed, and the two that
 * mean "line break" translated to "\n". Returns the input unchanged (the same
 * reference) when there is nothing to fix, which is nearly always. */
export function sanitizeXmlText(text: string): string {
  if (!NEEDS_FIXING.test(text)) {
    return text;
  }
  return text
    .replace(LINE_BREAK_RE, "\n")
    .replace(ILLEGAL_RE, "")
    .replace(LONE_SURROGATE_RE, "");
}
