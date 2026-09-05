// SPEC-056 OMML (Office Math Markup Language) serialiser and parser for the
// `.docx` path. An inline equation is `m:oMath`; a display equation is
// `m:oMathPara` wrapping one `m:oMath`; both live directly inside `w:p` with no
// drawing wrapper. A construct outside the closed set is preserved verbatim as
// an `opaqueMath` node (SPEC-008), its `source` re-emitted byte-for-byte.
import type { EquationSpec, MathNode } from "../documentModel";
import { childElements } from "../core/xml";
import { escapeXml, serializeElement } from "../core/xml";
import { linearForm } from "../math/linearForm";
import { MATH_NAMESPACE } from "./xml";

const M = MATH_NAMESPACE;

// ---- parse: OMML element → EquationSpec ----------------------------------

/** Parse an `m:oMath` or `m:oMathPara` element into an EquationSpec. */
export function parseOmmlEquation(el: Element): EquationSpec {
  const display = el.localName === "oMathPara" ? "block" : "inline";
  const oMath =
    el.localName === "oMathPara" ? childElements(el, M, "oMath")[0] ?? el : el;
  const body = parseRow(oMath);
  return { display, body };
}

function parseRow(parent: Element): MathNode[] {
  const out: MathNode[] = [];
  for (const element of Array.from(parent.children)) {
    if (element.namespaceURI !== M) continue;
    const node = parseNode(element);
    if (node) out.push(...node);
  }
  return out;
}

/** Parse the row content of an argument element (m:e, m:num, m:den, …). */
function argRow(parent: Element | undefined): MathNode[] {
  return parent ? parseRow(parent) : [];
}

function child(el: Element, name: string): Element | undefined {
  return childElements(el, M, name)[0];
}

function parseNode(el: Element): MathNode[] | null {
  switch (el.localName) {
    case "r":
      return parseRun(el);
    case "f": {
      const noBar = child(child(el, "fPr") ?? el, "type")?.getAttributeNS(M, "val") === "noBar";
      return [
        {
          kind: "frac",
          num: argRow(child(el, "num")),
          den: argRow(child(el, "den")),
          bar: !noBar,
        },
      ];
    }
    case "sSup":
      return [{ kind: "script", base: argRow(child(el, "e")), sup: argRow(child(el, "sup")) }];
    case "sSub":
      return [{ kind: "script", base: argRow(child(el, "e")), sub: argRow(child(el, "sub")) }];
    case "sSubSup":
      return [
        {
          kind: "script",
          base: argRow(child(el, "e")),
          sub: argRow(child(el, "sub")),
          sup: argRow(child(el, "sup")),
        },
      ];
    case "rad": {
      const degHide = child(child(el, "radPr") ?? el, "degHide")?.getAttributeNS(M, "val");
      const index = degHide === "1" || degHide === "true" ? undefined : argRow(child(el, "deg"));
      return [
        {
          kind: "radical",
          radicand: argRow(child(el, "e")),
          ...(index && index.length ? { index } : {}),
        },
      ];
    }
    case "nary": {
      const chr = child(child(el, "naryPr") ?? el, "chr")?.getAttributeNS(M, "val") ?? "∫";
      return [
        {
          kind: "nary",
          op: chr,
          lower: argRow(child(el, "sub")),
          upper: argRow(child(el, "sup")),
          body: argRow(child(el, "e")),
        },
      ];
    }
    case "d": {
      const pr = child(el, "dPr");
      const open = child(pr ?? el, "begChr")?.getAttributeNS(M, "val") ?? "(";
      const close = child(pr ?? el, "endChr")?.getAttributeNS(M, "val") ?? ")";
      const bodies = childElements(el, M, "e").map((e) => parseRow(e));
      // Multiple m:e children are a separator-joined list; join with the body
      // of the first for the common single-argument case.
      const body = bodies.length <= 1 ? bodies[0] ?? [] : bodies.flat();
      return [{ kind: "delim", open, close, body }];
    }
    case "m": {
      const rows = childElements(el, M, "mr").map((mr) =>
        childElements(mr, M, "e").map((e) => parseRow(e)),
      );
      return [{ kind: "matrix", rows }];
    }
    case "acc": {
      const chr = child(child(el, "accPr") ?? el, "chr")?.getAttributeNS(M, "val") ?? "̂";
      return [{ kind: "accent", accent: chr, base: argRow(child(el, "e")) }];
    }
    case "func": {
      const fName = child(el, "fName");
      const name = fName ? linearForm(parseRow(fName)).replace(/\\/g, "") : "";
      return [{ kind: "func", name, body: argRow(child(el, "e")) }];
    }
    case "oMath":
    case "oMathPara":
      return parseRow(el);
    default:
      // Outside the closed set: preserve verbatim (SPEC-008).
      return [{ kind: "opaqueMath", source: serializeElement(el), linear: el.textContent ?? "" }];
  }
}

function parseRun(el: Element): MathNode[] {
  const rPr = child(el, "rPr");
  const text = childElements(el, M, "t")
    .map((t) => t.textContent ?? "")
    .join("");
  if (!text) return [];
  const nor = rPr ? child(rPr, "nor") : undefined;
  if (nor) {
    return [{ kind: "text", text }];
  }
  const sty = rPr ? child(rPr, "sty")?.getAttributeNS(M, "val") : undefined;
  const variant = sty === "p" ? "upright" : sty === "i" ? "italic" : defaultVariant(text);
  return [{ kind: "sym", text, variant }];
}

function defaultVariant(text: string): "italic" | "upright" {
  // OMML defaults a maths run to italic; a digit or operator run is upright.
  return /^[A-Za-zΑ-ϖ]$/.test(text) ? "italic" : "upright";
}

// ---- serialise: EquationSpec → OMML string -------------------------------

/** Serialise an equation to an `m:oMath`/`m:oMathPara` string for a `w:p`. */
export function serializeOmmlEquation(eq: EquationSpec): string {
  const inner = eq.body.map(nodeXml).join("");
  const oMath = `<m:oMath>${inner}</m:oMath>`;
  return eq.display === "block" ? `<m:oMathPara>${oMath}</m:oMathPara>` : oMath;
}

function rowXml(nodes: MathNode[]): string {
  return nodes.map(nodeXml).join("");
}

function nodeXml(node: MathNode): string {
  switch (node.kind) {
    case "sym": {
      const sty = node.variant === "italic" ? "i" : "p";
      return `<m:r><m:rPr><m:sty m:val="${sty}"/></m:rPr><m:t xml:space="preserve">${escapeXml(node.text)}</m:t></m:r>`;
    }
    case "text":
      return `<m:r><m:rPr><m:nor/></m:rPr><m:t xml:space="preserve">${escapeXml(node.text)}</m:t></m:r>`;
    case "frac": {
      const pr = node.bar ? "" : `<m:fPr><m:type m:val="noBar"/></m:fPr>`;
      return `<m:f>${pr}<m:num>${rowXml(node.num)}</m:num><m:den>${rowXml(node.den)}</m:den></m:f>`;
    }
    case "script": {
      if (node.sub && node.sup) {
        return `<m:sSubSup><m:e>${rowXml(node.base)}</m:e><m:sub>${rowXml(node.sub)}</m:sub><m:sup>${rowXml(node.sup)}</m:sup></m:sSubSup>`;
      }
      if (node.sub) {
        return `<m:sSub><m:e>${rowXml(node.base)}</m:e><m:sub>${rowXml(node.sub)}</m:sub></m:sSub>`;
      }
      return `<m:sSup><m:e>${rowXml(node.base)}</m:e><m:sup>${rowXml(node.sup ?? [])}</m:sup></m:sSup>`;
    }
    case "radical": {
      const hasIndex = node.index && node.index.length > 0;
      const pr = hasIndex ? "" : `<m:radPr><m:degHide m:val="1"/></m:radPr>`;
      const deg = hasIndex ? `<m:deg>${rowXml(node.index ?? [])}</m:deg>` : `<m:deg/>`;
      return `<m:rad>${pr}${deg}<m:e>${rowXml(node.radicand)}</m:e></m:rad>`;
    }
    case "nary": {
      const pr = `<m:naryPr><m:chr m:val="${escapeXml(node.op)}"/></m:naryPr>`;
      const sub = `<m:sub>${rowXml(node.lower ?? [])}</m:sub>`;
      const sup = `<m:sup>${rowXml(node.upper ?? [])}</m:sup>`;
      return `<m:nary>${pr}${sub}${sup}<m:e>${rowXml(node.body)}</m:e></m:nary>`;
    }
    case "delim": {
      const pr = `<m:dPr><m:begChr m:val="${escapeXml(node.open)}"/><m:endChr m:val="${escapeXml(node.close)}"/></m:dPr>`;
      return `<m:d>${pr}<m:e>${rowXml(node.body)}</m:e></m:d>`;
    }
    case "matrix": {
      const rows = node.rows
        .map((row) => `<m:mr>${row.map((cell) => `<m:e>${rowXml(cell)}</m:e>`).join("")}</m:mr>`)
        .join("");
      return `<m:m>${rows}</m:m>`;
    }
    case "accent": {
      const pr = `<m:accPr><m:chr m:val="${escapeXml(node.accent)}"/></m:accPr>`;
      return `<m:acc>${pr}<m:e>${rowXml(node.base)}</m:e></m:acc>`;
    }
    case "func":
      return `<m:func><m:fName><m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t xml:space="preserve">${escapeXml(node.name)}</m:t></m:r></m:fName><m:e>${rowXml(node.body)}</m:e></m:func>`;
    case "opaqueMath":
      // Byte-for-byte preservation of the original OMML subtree (SPEC-008).
      return node.source;
  }
}
