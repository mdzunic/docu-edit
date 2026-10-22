// The span-diff helper of SPEC-124's retained-reference specs: an adjusted
// part may differ from its source only inside attribute values and text nodes
// (the spans rewriteRetainedReferences splices), plus whole elements a test
// names as removed. The XML is cut into markup and value chunks — a tag's
// text up to and including each attribute's opening quote is markup, the
// value up to its closing quote a value chunk; text between tags (and a CDATA
// section) is a value chunk; comments and processing instructions are markup
// — and the two chunk sequences must agree on every markup chunk, in order.
// Works on fragments (an ods sourceXml, an envelope prefix) as well as parts.

interface Chunk {
  markup: boolean;
  text: string;
}

function chunks(xml: string): Chunk[] {
  const out: Chunk[] = [];
  let i = 0;
  const push = (markup: boolean, text: string): void => {
    if (text !== "" || !markup) {
      out.push({ markup, text });
    }
  };
  while (i < xml.length) {
    if (xml[i] !== "<") {
      const next = xml.indexOf("<", i);
      const end = next < 0 ? xml.length : next;
      push(false, xml.slice(i, end));
      i = end;
      continue;
    }
    for (const [open, close, markup] of [
      ["<!--", "-->", true],
      ["<![CDATA[", "]]>", false],
      ["<?", "?>", true],
    ] as const) {
      if (xml.startsWith(open, i)) {
        const end = xml.indexOf(close, i);
        const stop = end < 0 ? xml.length : end + close.length;
        push(markup, xml.slice(i, stop));
        i = stop;
        break;
      }
    }
    if (xml[i] !== "<" || xml.startsWith("<!--", i) || xml.startsWith("<![CDATA[", i) || xml.startsWith("<?", i)) {
      continue;
    }
    // A tag: markup through each opening quote, the value, then on.
    let start = i;
    let j = i + 1;
    while (j < xml.length && xml[j] !== ">") {
      const ch = xml[j]!;
      if (ch === '"' || ch === "'") {
        push(true, xml.slice(start, j + 1));
        const close = xml.indexOf(ch, j + 1);
        const end = close < 0 ? xml.length : close;
        push(false, xml.slice(j + 1, end));
        start = end;
        j = end + 1;
        continue;
      }
      j += 1;
    }
    push(true, xml.slice(start, Math.min(j + 1, xml.length)));
    i = j + 1;
  }
  return out;
}

/** One attribute value or text node that differs. */
export interface ChangedSpan {
  before: string;
  after: string;
}

/** The attribute values and text nodes `after` changed in `before` — in
 * document order — after `removed` (substrings of `before`, each removed
 * once) are taken out. Throws when anything else differs. */
export function changedSpans(before: string, after: string, removed: readonly string[] = []): ChangedSpan[] {
  let source = before;
  for (const element of removed) {
    const at = source.indexOf(element);
    if (at < 0) {
      throw new Error(`span diff: the removed text is not in the source: ${element.slice(0, 80)}`);
    }
    source = source.slice(0, at) + source.slice(at + element.length);
  }
  const a = chunks(source);
  const b = chunks(after);
  if (a.length !== b.length) {
    throw new Error(`span diff: the markup differs (${a.length} chunks before, ${b.length} after)`);
  }
  const changed: ChangedSpan[] = [];
  for (let k = 0; k < a.length; k += 1) {
    const x = a[k]!;
    const y = b[k]!;
    if (x.markup !== y.markup || (x.markup && x.text !== y.text)) {
      throw new Error(`span diff: markup differs at chunk ${k}: ${JSON.stringify(x.text.slice(0, 80))} vs ${JSON.stringify(y.text.slice(0, 80))}`);
    }
    if (!x.markup && x.text !== y.text) {
      changed.push({ before: x.text, after: y.text });
    }
  }
  return changed;
}
