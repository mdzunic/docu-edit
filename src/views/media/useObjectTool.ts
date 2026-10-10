// The redact and annotate tools (SPEC-112): rectangles and shapes drawn over
// the image in output pixels. Drawing is a gesture — one undo entry per
// stroke, labelled as the history list reads it ("Redact", "Add arrow") —
// and every object is also reachable without a pointer: "Add" places one at
// a default spot, and the object list is focusable, an arrow key nudging the
// focused object 1 px (10 with Shift) and Delete removing it.
import { computed, ref, type ComputedRef, type Ref } from "vue";
import { canAddObject, removeAt, replaceAt } from "../../lib/core/imageEditHistory";
import {
  annotationNoun,
  DEFAULT_REDACT_COLOR,
  describeEdit,
  describeEditDetail,
  isObjectEdit,
  MAX_FREEHAND_POINTS,
  translateEdit,
  type AnnotationKind,
  type AnnotationShape,
  type ImageEdit,
} from "../../lib/core/imageEdits";
import type { EditSession } from "./editSession";

export const OBJECT_STEP_PX = 1;
export const OBJECT_SHIFT_STEP_PX = 10;

/** A drag shorter than this (output px) on both axes draws nothing. */
const MIN_DRAG_PX = 2;

export interface EditObject {
  /** The object's index in the history list. */
  index: number;
  kind: "redact" | "annotate";
  label: string;
  detail: string;
}

export interface ObjectTool {
  redactColor: Ref<string>;
  shape: Ref<AnnotationKind>;
  color: Ref<string>;
  strokeWidth: Ref<number>;
  opacity: Ref<number>;
  text: Ref<string>;
  fontSize: Ref<number>;
  objects: ComputedRef<EditObject[]>;
  selected: Ref<number | null>;
  /** Whether the stack bound (MAX_EDIT_STACK) leaves no room for another. */
  full: ComputedRef<boolean>;
  /** Places one object of the active tool at a default spot. */
  add: (tool: "redact" | "annotate") => void;
  startDraw: (tool: "redact" | "annotate", x: number, y: number) => void;
  drawTo: (x: number, y: number) => void;
  endDraw: () => void;
  nudge: (index: number, dx: number, dy: number) => void;
  remove: (index: number) => void;
  /** Arrow keys and Delete on a focused object; true when handled. */
  handleKey: (index: number, event: KeyboardEvent) => boolean;
}

function noun(edit: ImageEdit): string {
  return edit.kind === "annotate" ? annotationNoun(edit.shape.shape) : "redaction";
}

export function useObjectTool(session: EditSession): ObjectTool {
  const redactColor = ref(DEFAULT_REDACT_COLOR);
  const shape = ref<AnnotationKind>("arrow");
  const color = ref("#e53935");
  const strokeWidth = ref(4);
  const opacity = ref(1);
  const text = ref("Note");
  const fontSize = ref(24);
  const selected = ref<number | null>(null);

  const objects = computed(() =>
    session.history.value.flatMap((edit, index) =>
      isObjectEdit(edit) ? [{ index, kind: edit.kind, label: describeEdit(edit), detail: describeEditDetail(edit) }] : [],
    ),
  );
  const full = computed(() => !canAddObject(session.history.value));

  function annotation(kind: AnnotationKind, x1: number, y1: number, x2: number, y2: number, points?: Array<{ x: number; y: number }>): ImageEdit {
    const box = { x: Math.min(x1, x2), y: Math.min(y1, y2), width: Math.abs(x2 - x1), height: Math.abs(y2 - y1) };
    let drawn: AnnotationShape;
    if (kind === "arrow" || kind === "line") {
      drawn = { shape: kind, x1, y1, x2, y2 };
    } else if (kind === "freehand") {
      drawn = { shape: "freehand", points: points ?? [{ x: x1, y: y1 }, { x: x2, y: y2 }] };
    } else if (kind === "text") {
      drawn = { shape: "text", x: x1, y: y1, text: text.value.trim() || "Note", fontSize: fontSize.value };
    } else {
      drawn = { shape: kind, ...box };
    }
    const fill = kind === "highlight" ? Math.min(opacity.value, 0.5) : opacity.value;
    return { kind: "annotate", shape: drawn, color: color.value, strokeWidth: strokeWidth.value, opacity: fill };
  }

  function redaction(x1: number, y1: number, x2: number, y2: number): ImageEdit {
    return {
      kind: "redact",
      x: Math.round(Math.min(x1, x2)),
      y: Math.round(Math.min(y1, y2)),
      width: Math.max(1, Math.round(Math.abs(x2 - x1))),
      height: Math.max(1, Math.round(Math.abs(y2 - y1))),
      color: redactColor.value,
    };
  }

  function add(tool: "redact" | "annotate"): void {
    const geometry = session.geometry.value;
    if (!geometry || full.value) return;
    const w = geometry.widthPx;
    const h = geometry.heightPx;
    const edit =
      tool === "redact"
        ? redaction(w * 0.375, h * 0.375, w * 0.625, h * 0.625)
        : annotation(shape.value, w * 0.25, h * 0.25, w * 0.75, h * 0.75);
    session.apply(describeEdit(edit), [...session.history.value, edit]);
    selected.value = session.history.value.length - 1;
  }

  let draw: { tool: "redact" | "annotate"; kind: AnnotationKind; x: number; y: number; index: number; points: Array<{ x: number; y: number }> } | null = null;

  function startDraw(tool: "redact" | "annotate", x: number, y: number): void {
    if (full.value || !session.geometry.value) return;
    const edit = tool === "redact" ? redaction(x, y, x, y) : annotation(shape.value, x, y, x, y, [{ x, y }]);
    session.begin(describeEdit(edit));
    session.update([...session.history.value, edit]);
    draw = { tool, kind: shape.value, x, y, index: session.history.value.length - 1, points: [{ x, y }] };
  }

  function drawTo(x: number, y: number): void {
    if (!draw) return;
    if (draw.kind === "freehand" && draw.points.length < MAX_FREEHAND_POINTS) {
      draw.points.push({ x, y });
    }
    const edit = draw.tool === "redact" ? redaction(draw.x, draw.y, x, y) : annotation(draw.kind, draw.x, draw.y, x, y, [...draw.points]);
    session.update(replaceAt(session.history.value, draw.index, edit));
  }

  function endDraw(): void {
    if (!draw) return;
    const edit = session.history.value[draw.index];
    if (!edit || isDegenerate(edit)) {
      session.cancel();
    } else {
      selected.value = draw.index;
      session.end();
    }
    draw = null;
  }

  function nudge(index: number, dx: number, dy: number): void {
    const edit = session.history.value[index];
    if (!edit || !isObjectEdit(edit)) return;
    session.apply(`Move ${noun(edit)}`, replaceAt(session.history.value, index, translateEdit(edit, dx, dy)));
  }

  function remove(index: number): void {
    const edit = session.history.value[index];
    if (!edit || !isObjectEdit(edit)) return;
    session.apply(`Delete ${noun(edit)}`, removeAt(session.history.value, index));
    selected.value = null;
  }

  function handleKey(index: number, event: KeyboardEvent): boolean {
    if (event.key === "Delete" || event.key === "Backspace") {
      remove(index);
      return true;
    }
    const step = event.shiftKey ? OBJECT_SHIFT_STEP_PX : OBJECT_STEP_PX;
    const delta: Record<string, [number, number]> = {
      ArrowLeft: [-step, 0],
      ArrowRight: [step, 0],
      ArrowUp: [0, -step],
      ArrowDown: [0, step],
    };
    const move = delta[event.key];
    if (!move || event.altKey || event.metaKey || event.ctrlKey) {
      return false;
    }
    nudge(index, move[0], move[1]);
    return true;
  }

  return {
    redactColor,
    shape,
    color,
    strokeWidth,
    opacity,
    text,
    fontSize,
    objects,
    selected,
    full,
    add,
    startDraw,
    drawTo,
    endDraw,
    nudge,
    remove,
    handleKey,
  };
}

/** A stroke that draws nothing worth keeping: a click rather than a drag
 * (text is placed by a click, so it always counts). */
function isDegenerate(edit: ImageEdit): boolean {
  if (edit.kind === "redact") {
    return edit.width < MIN_DRAG_PX && edit.height < MIN_DRAG_PX;
  }
  if (edit.kind !== "annotate") {
    return false;
  }
  const shape = edit.shape;
  switch (shape.shape) {
    case "text":
      return false;
    case "freehand":
      return shape.points.length < 2;
    case "arrow":
    case "line":
      return Math.abs(shape.x2 - shape.x1) < MIN_DRAG_PX && Math.abs(shape.y2 - shape.y1) < MIN_DRAG_PX;
    default:
      return shape.width < MIN_DRAG_PX && shape.height < MIN_DRAG_PX;
  }
}
