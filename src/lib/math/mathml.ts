// SPEC-056 Presentation-MathML serialiser and parser for the `.odt` equation
// path. An equation is a `<math>` element (the body of an ODF formula
// content.xml); its `display` attribute carries EquationSpec.display. The
// mapping is deliberately structurally reversible: `parseMathML` recognises
// exactly what `serializeMathML` emits, so a round trip reproduces the model
// (compared with structural equality that ignores `id`s). A construct outside
// the closed set is preserved verbatim as an `opaqueMath` node (SPEC-008),
// its `source` re-emitted as-is.
import type { EquationSpec, MathNode } from "../documentModel";
import { escapeXml, serializeElement } from "../core/xml";

export const MATHML_NAMESPACE = "http://www.w3.org/1998/Math/MathML";

// U+2061 FUNCTION APPLICATION — the invisible operator between a function name
// and its argument.
const APPLY_FUNCTION = "⁡";

// ---- serialise: EquationSpec → `<math>` string ---------------------------

/** Serialise an equation to a `<math>` element string (the body of an ODF
 * formula content.xml). */
export function serializeMathML(eq: EquationSpec): string {
  return `<math xmlns="${MATHML_NAMESPACE}" display="${eq.display}">${rowMathML(eq.body)}</math>`;
}

/** A row of nodes serialises to a bare concatenation of its nodes' elements. */
function rowMathML(nodes: MathNode[]): string {
  return nodes.map(nodeMathML).join("");
}

/** A slot needing a single MathML child always wraps its row in `<mrow>`, even
 * for a single node, so parsing is uniform. */
function slotMathML(nodes: MathNode[]): string {
  return `<mrow>${rowMathML(nodes)}</mrow>`;
}

function nodeMathML(node: MathNode): string {
  switch (node.kind) {
    case "sym": {
      const t = escapeXml(node.text);
      if (node.variant === "italic") {
        return `<mi>${t}</mi>`;
      }
      return /^[0-9.,]+$/.test(node.text) ? `<mn>${t}</mn>` : `<mo>${t}</mo>`;
    }
    case "text":
      return `<mtext>${escapeXml(node.text)}</mtext>`;
    case "frac": {
      const thin = node.bar ? "" : ` linethickness="0"`;
      return `<mfrac${thin}>${slotMathML(node.num)}${slotMathML(node.den)}</mfrac>`;
    }
    case "script": {
      if (node.sub && node.sup) {
        return `<msubsup>${slotMathML(node.base)}${slotMathML(node.sub)}${slotMathML(node.sup)}</msubsup>`;
      }
      if (node.sub) {
        return `<msub>${slotMathML(node.base)}${slotMathML(node.sub)}</msub>`;
      }
      return `<msup>${slotMathML(node.base)}${slotMathML(node.sup ?? [])}</msup>`;
    }
    case "radical": {
      if (node.index && node.index.length > 0) {
        return `<mroot>${slotMathML(node.radicand)}${slotMathML(node.index)}</mroot>`;
      }
      return `<msqrt>${slotMathML(node.radicand)}</msqrt>`;
    }
    case "nary": {
      const munderover = `<munderover><mo>${escapeXml(node.op)}</mo>${slotMathML(node.lower ?? [])}${slotMathML(node.upper ?? [])}</munderover>`;
      return `<mrow>${munderover}${slotMathML(node.body)}</mrow>`;
    }
    case "delim": {
      const open =
        node.open === "" ? "" : `<mo fence="true" stretchy="true">${escapeXml(node.open)}</mo>`;
      const close =
        node.close === "" ? "" : `<mo fence="true" stretchy="true">${escapeXml(node.close)}</mo>`;
      return `<mrow>${open}${slotMathML(node.body)}${close}</mrow>`;
    }
    case "matrix": {
      const rows = node.rows
        .map(
          (row) => `<mtr>${row.map((cell) => `<mtd>${slotMathML(cell)}</mtd>`).join("")}</mtr>`,
        )
        .join("");
      return `<mtable>${rows}</mtable>`;
    }
    case "accent":
      return `<mover accent="true">${slotMathML(node.base)}<mo>${escapeXml(node.accent)}</mo></mover>`;
    case "func":
      return `<mrow class="fn"><mi mathvariant="normal">${escapeXml(node.name)}</mi><mo>${APPLY_FUNCTION}</mo>${slotMathML(node.body)}</mrow>`;
    case "opaqueMath":
      // Verbatim preservation of the original MathML subtree (SPEC-008).
      return node.source;
  }
}

// ---- parse: `<math>` Element → EquationSpec ------------------------------

/** Parse a `<math>` Element into an EquationSpec. Never throws — an
 * unrecognised element becomes an `opaqueMath` node. */
export function parseMathML(mathEl: Element): EquationSpec {
  const display = mathEl.getAttribute("display") === "block" ? "block" : "inline";
  return { display, body: parseRow(elementChildren(mathEl)) };
}

function elementChildren(el: Element): Element[] {
  return Array.from(el.children);
}

function textOf(el: Element | undefined): string {
  return el?.textContent ?? "";
}

function isEl(el: Element | undefined, localName: string): boolean {
  return !!el && el.namespaceURI === MATHML_NAMESPACE && el.localName === localName;
}

function isFenceMo(el: Element | undefined): boolean {
  return isEl(el, "mo") && el!.getAttribute("fence") === "true";
}

function isFuncMrow(el: Element): boolean {
  const kids = elementChildren(el);
  return (
    isEl(kids[0], "mi") &&
    kids[0].getAttribute("mathvariant") === "normal" &&
    isEl(kids[1], "mo") &&
    textOf(kids[1]).includes(APPLY_FUNCTION)
  );
}

function isNaryMrow(el: Element): boolean {
  const kids = elementChildren(el);
  if (!isEl(kids[0], "munderover")) {
    return false;
  }
  return isEl(elementChildren(kids[0])[0], "mo");
}

function isDelimMrow(el: Element): boolean {
  const kids = elementChildren(el);
  if (kids.length === 0) {
    return false;
  }
  return isFenceMo(kids[0]) || isFenceMo(kids[kids.length - 1]);
}

function isSpecialMrow(el: Element): boolean {
  return isEl(el, "mrow") && (isFuncMrow(el) || isNaryMrow(el) || isDelimMrow(el));
}

/** Parse a slot's children back into a row. When the slot is a single plain
 * `<mrow>` wrapper (not a delim/nary/func node that also happens to be an
 * mrow), descend into it; otherwise parse the children directly. */
function parseSlot(children: Array<Element | undefined>): MathNode[] {
  const kids = children.filter((c): c is Element => !!c);
  if (kids.length === 1 && isEl(kids[0], "mrow") && !isSpecialMrow(kids[0])) {
    return parseRow(elementChildren(kids[0]));
  }
  return parseRow(kids);
}

function parseRow(children: Array<Element | undefined>): MathNode[] {
  const out: MathNode[] = [];
  for (const child of children) {
    if (child) {
      out.push(...parseNode(child));
    }
  }
  return out;
}

function parseNode(el: Element): MathNode[] {
  switch (el.localName) {
    case "mi": {
      const variant = el.getAttribute("mathvariant") === "normal" ? "upright" : "italic";
      return [{ kind: "sym", text: textOf(el), variant }];
    }
    case "mn":
    case "mo":
      return [{ kind: "sym", text: textOf(el), variant: "upright" }];
    case "mtext":
      return [{ kind: "text", text: textOf(el) }];
    case "mfrac": {
      const kids = elementChildren(el);
      return [
        {
          kind: "frac",
          num: parseSlot([kids[0]]),
          den: parseSlot([kids[1]]),
          bar: el.getAttribute("linethickness") !== "0",
        },
      ];
    }
    case "msup": {
      const kids = elementChildren(el);
      return [{ kind: "script", base: parseSlot([kids[0]]), sup: parseSlot([kids[1]]) }];
    }
    case "msub": {
      const kids = elementChildren(el);
      return [{ kind: "script", base: parseSlot([kids[0]]), sub: parseSlot([kids[1]]) }];
    }
    case "msubsup": {
      const kids = elementChildren(el);
      return [
        {
          kind: "script",
          base: parseSlot([kids[0]]),
          sub: parseSlot([kids[1]]),
          sup: parseSlot([kids[2]]),
        },
      ];
    }
    case "msqrt":
      return [{ kind: "radical", radicand: parseSlot(elementChildren(el)) }];
    case "mroot": {
      const kids = elementChildren(el);
      const index = parseSlot([kids[1]]);
      return [
        {
          kind: "radical",
          radicand: parseSlot([kids[0]]),
          ...(index.length > 0 ? { index } : {}),
        },
      ];
    }
    case "mtable": {
      const rows = elementChildren(el)
        .filter((mtr) => isEl(mtr, "mtr"))
        .map((mtr) =>
          elementChildren(mtr)
            .filter((mtd) => isEl(mtd, "mtd"))
            .map((mtd) => parseSlot(elementChildren(mtd))),
        );
      return [{ kind: "matrix", rows }];
    }
    case "mover": {
      const kids = elementChildren(el);
      return [
        {
          kind: "accent",
          accent: textOf(kids[kids.length - 1]),
          base: parseSlot([kids[0]]),
        },
      ];
    }
    case "mrow":
      return parseMrow(el);
    default:
      // Outside the closed set: preserve verbatim (SPEC-008).
      return [{ kind: "opaqueMath", source: serializeElement(el), linear: el.textContent ?? "" }];
  }
}

function parseMrow(el: Element): MathNode[] {
  const kids = elementChildren(el);
  // func before the generic delim/nary rules.
  if (isFuncMrow(el)) {
    return [
      {
        kind: "func",
        name: textOf(kids[0]),
        body: kids.length > 2 ? parseSlot([kids[2]]) : [],
      },
    ];
  }
  // nary (munderover-first mrow) before delim.
  if (isNaryMrow(el)) {
    const munderover = kids[0];
    const mkids = elementChildren(munderover);
    const lower = parseSlot([mkids[1]]);
    const upper = parseSlot([mkids[2]]);
    return [
      {
        kind: "nary",
        op: textOf(mkids[0]),
        ...(lower.length > 0 ? { lower } : {}),
        ...(upper.length > 0 ? { upper } : {}),
        body: parseSlot(kids.slice(1)),
      },
    ];
  }
  if (isDelimMrow(el)) {
    let start = 0;
    let end = kids.length;
    let open = "";
    let close = "";
    if (isFenceMo(kids[0])) {
      open = textOf(kids[0]);
      start = 1;
    }
    if (end - 1 >= start && isFenceMo(kids[end - 1])) {
      close = textOf(kids[end - 1]);
      end = end - 1;
    }
    return [{ kind: "delim", open, close, body: parseSlot(kids.slice(start, end)) }];
  }
  // A plain grouping mrow at row level — descend.
  return parseRow(kids);
}
