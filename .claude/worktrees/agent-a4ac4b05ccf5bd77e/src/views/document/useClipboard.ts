// Cut, copy and paste (SPEC-016). Fully controlled like typing: the native
// ClipboardEvents are cancelled and re-implemented over the model, because only
// these synchronous handlers can carry the private document-model MIME type
// (navigator.clipboard.read/write accept a fixed type allowlist), and the
// browser must never splice its own HTML into the canvas. Also owns the paste
// options chip (re-applying a paste under a different interpretation) and the
// Ctrl/Cmd+Shift+V plain-paste path.
//
// The context menu's cut/copy/paste handshake uses this module's flavours but
// lives with the context menu in DocumentView; `serializeFragment` is passed
// in where the menu needs to recognise its own copy.
import { nextTick, ref } from "vue";
import {
  asPasteMode,
  clipboardFlavours,
  DOC_MODEL_MIME,
  executePaste,
  extractFragment,
  IMAGES_UNSUPPORTED_MESSAGE,
  noClipboardSnapshot,
  noPastePayload,
  parseFragmentPayload,
  type ClipboardSnapshot,
  type PastePayload,
} from "../../lib/clipboard";
import { importHtml } from "../../lib/htmlPaste";
import { caretClientRect } from "../../lib/caretGeometry";
import { deleteRange, isCollapsed, noPendingFormat } from "../../lib/commands";
import { isAcceptedImageType } from "../../lib/core/media";
import { undo as undoDocument } from "../../lib/undoHistory";
import type { ModelSelection } from "../../lib/commands";
import type { DocumentModel, PendingFormat } from "../../lib/documentModel";
import type { EditorContext } from "./context";

/** The clipboard flavours a paste event carries into `performPaste`. */
export interface PasteFlavours {
  internal: string;
  html: string;
  text: string;
  hasImage: boolean;
}

export interface ClipboardDeps extends EditorContext {
  getPendingFormat: () => PendingFormat | null;
  /** The image-insert path (SPEC-024) a paste with a raw image File routes
   * into — passed in because image insertion lives with the images feature. */
  insertImageFromFile: (file: File, selection: ModelSelection | null) => void;
}

function clipboardHasImage(data: DataTransfer): boolean {
  const items = data.items;
  if (!items) {
    return false;
  }
  for (let i = 0; i < items.length; i += 1) {
    if (items[i].kind === "file" && items[i].type.startsWith("image/")) {
      return true;
    }
  }
  return false;
}

/** The first clipboard file item of an accepted image type, or null. */
function clipboardImageFile(data: DataTransfer): File | null {
  const items = data.items;
  if (!items) {
    return null;
  }
  for (let i = 0; i < items.length; i += 1) {
    if (items[i].kind === "file" && isAcceptedImageType(items[i].type) && typeof items[i].getAsFile === "function") {
      return items[i].getAsFile();
    }
  }
  return null;
}

export function useClipboard(deps: ClipboardDeps) {
  const {
    doc,
    contentRoot,
    editable,
    documentId,
    currentSelection,
    applyEdit,
    addToast,
    getCaret,
    getPendingFormat,
    insertImageFromFile,
    onSessionClose,
  } = deps;

  // `plainPasteNext` is armed by the Ctrl/Cmd+Shift+V keydown (AC-27): where
  // the browser fires a `paste` event for that combination the flag routes it
  // down the plain-text path, and where it does not (macOS — see
  // schedulePlainPasteFallback) it survives long enough for the fallback to
  // notice. `pastePayload`/`pastePendingFormat` describe the last applied paste
  // so the options chip can re-interpret it; `lastCopy` is the local snapshot
  // backing the context menu's async paste. Plain variables — nothing renders
  // from them; the chip renders from the `pasteChip` ref below.
  let plainPasteNext = false;
  let pastePayload = noPastePayload();
  let pastePendingFormat = noPendingFormat();
  let lastCopy = noClipboardSnapshot();
  const pasteChip = ref({ visible: false, left: 0, top: 0, applied: "source" as string });

  /** Captures the selection's fragment and writes all three clipboard flavours
   * (AC-7..9) and remembers the snapshot locally for the context menu's paste. */
  function writeClipboardFlavours(
    data: DataTransfer,
    model: DocumentModel,
    selection: ModelSelection,
  ): void {
    if (!model || !selection) {
      return;
    }
    const fragment = extractFragment(model, selection);
    const flavours = clipboardFlavours(fragment);
    data.setData("text/plain", flavours.text);
    data.setData("text/html", flavours.html);
    data.setData(DOC_MODEL_MIME, flavours.payload);
    lastCopy = { text: flavours.text, html: flavours.html, fragment };
  }

  /** Ctrl/Cmd+C and the context menu's Copy (AC-2, AC-5). Copying mutates
   * nothing, so it works in a read-only (conflict-bar) session too. */
  function handleCopy(event: ClipboardEvent): void {
    const root = contentRoot.value as HTMLElement | null;
    const model = doc.value;
    const data = event.clipboardData;
    if (!root || !model || !data) {
      return;
    }
    const selection = currentSelection(root);
    if (!selection || isCollapsed(selection)) {
      return;
    }
    event.preventDefault();
    writeClipboardFlavours(data, model, selection);
  }

  /** Ctrl/Cmd+X and the context menu's Cut (AC-1, AC-4): the copy flavours
   * plus deleting the selection, recorded as a single "cut" undo step. */
  function handleCut(event: ClipboardEvent): void {
    const root = contentRoot.value as HTMLElement | null;
    const model = doc.value;
    const data = event.clipboardData;
    if (!root || !model || !data || !editable.value) {
      return;
    }
    const selection = currentSelection(root);
    if (!selection || isCollapsed(selection)) {
      return;
    }
    event.preventDefault();
    writeClipboardFlavours(data, model, selection);
    applyEdit(deleteRange(model, selection), selection, "cut");
  }

  /** Ctrl/Cmd+V and the native menu's Paste (AC-3, AC-6): reads the flavours
   * off the event and routes them through performPaste. Always prevented — the
   * browser must never splice its own HTML into the canvas. */
  function handlePaste(event: ClipboardEvent): void {
    const root = contentRoot.value as HTMLElement | null;
    const model = doc.value;
    if (!root || !model || !editable.value) {
      return;
    }
    event.preventDefault();
    const data = event.clipboardData;
    if (!data) {
      return;
    }
    const plainOnly = plainPasteNext;
    plainPasteNext = false;
    const selection = currentSelection(root);
    if (!selection) {
      return;
    }
    const imageFile = clipboardImageFile(data);
    if (imageFile) {
      void insertImageFromFile(imageFile, selection);
      return;
    }
    performPaste(
      {
        internal: data.getData(DOC_MODEL_MIME),
        html: data.getData("text/html"),
        text: data.getData("text/plain"),
        hasImage: clipboardHasImage(data),
      },
      selection,
      plainOnly,
    );
  }

  /** Ctrl/Cmd+Shift+V's fallback path: Chromium only turns that combination
   * into a clipboard command on Windows/Linux; on macOS it never reaches the
   * page, so the armed flag would sit there. Read text ourselves as a fallback. */
  function schedulePlainPasteFallback(): void {
    window.setTimeout(() => {
      if (!plainPasteNext) {
        // The paste event arrived and consumed the flag.
        return;
      }
      plainPasteNext = false;
      void pastePlainFromClipboard();
    }, 0);
  }

  async function pastePlainFromClipboard(): Promise<void> {
    let text = "";
    try {
      text = await navigator.clipboard.readText();
    } catch {
      return;
    }
    const root = contentRoot.value as HTMLElement | null;
    const model = doc.value;
    if (!root || !model || !editable.value || text === "") {
      return;
    }
    const selection = currentSelection(root);
    if (!selection) {
      return;
    }
    performPaste({ internal: "", html: "", text, hasImage: false }, selection, true);
  }

  /** One paste, whatever triggered it: the internal payload wins (exact
   * editor-to-editor fidelity), then the sanitised HTML import, then plain
   * text; `plainOnly` forces the text path but keeps any richer fragment so
   * the chip can still offer "Keep source formatting". */
  function performPaste(
    flavours: PasteFlavours,
    selection: ModelSelection | null,
    plainOnly = false,
  ): void {
    const model = doc.value;
    if (!model || !selection) {
      return;
    }
    let fragment = flavours.internal ? parseFragmentPayload(flavours.internal) : null;
    let imagesDropped = false;
    if (!fragment && flavours.html) {
      const imported = importHtml(flavours.html);
      imagesDropped = imported.imagesDropped;
      if (imported.blocks.length > 0) {
        fragment = imported.blocks;
      }
    }
    if (imagesDropped || flavours.hasImage) {
      addToast(IMAGES_UNSUPPORTED_MESSAGE);
    }
    if (!fragment && !flavours.text) {
      return;
    }
    const payload: PastePayload = { fragment, text: flavours.text };
    const mode = plainOnly || !fragment ? "text" : "source";
    const pendingFormat = getPendingFormat();
    const result = executePaste(model, selection, payload, mode, pendingFormat);
    pastePayload = payload;
    pastePendingFormat = pendingFormat ? { ...pendingFormat } : noPendingFormat();
    applyEdit(result, selection, "paste", { keepPasteChip: true });
    void openPasteChip(mode);
  }

  /** Shows the chip pinned under the post-paste caret (AC-29). */
  async function openPasteChip(applied = "source"): Promise<void> {
    await nextTick();
    const root = contentRoot.value as HTMLElement | null;
    const caret = getCaret();
    const rect = root && caret ? caretClientRect(root, caret) : null;
    pasteChip.value = {
      visible: true,
      left: rect ? Math.max(8, rect.left) : 16,
      top: rect ? rect.bottom + 8 : 16,
      applied,
    };
  }

  function dismissPasteChip(): void {
    if (pasteChip.value.visible) {
      pasteChip.value = { visible: false, left: 0, top: 0, applied: "source" };
    }
    pastePayload = noPastePayload();
    pastePendingFormat = noPendingFormat();
  }

  /** A chip choice (AC-30..32): re-applies the same paste under the chosen
   * interpretation by replacing the paste's undo step — pop it, restore the
   * pre-paste state, re-run the paste from the stored payload — so the paste
   * plus the re-application stay one undo step (AC-33). */
  function handlePasteOption(mode = ""): void {
    const root = contentRoot.value as HTMLElement | null;
    const payload = pastePayload;
    if (!root || !editable.value || !payload || !pasteChip.value.visible) {
      return;
    }
    const chosen = asPasteMode(mode);
    if (chosen === pasteChip.value.applied) {
      root.focus();
      return;
    }
    const snapshot = undoDocument(documentId.value);
    if (!snapshot) {
      return;
    }
    doc.value = snapshot.model;
    const result = executePaste(
      snapshot.model,
      snapshot.selection,
      payload,
      chosen,
      pastePendingFormat,
    );
    applyEdit(result, snapshot.selection, "paste", { keepPasteChip: true });
    void openPasteChip(chosen);
    root.focus();
  }

  /** Ctrl/Cmd+Shift+V: arm the plain-text flag and let the browser fire the
   * paste event it produces for this combo (macOS falls back via the timer). */
  function armPlainPaste(): void {
    plainPasteNext = true;
    schedulePlainPasteFallback();
  }

  /** A plain Ctrl/Cmd+V clears a stale plain-paste flag (armed but never
   * consumed because an empty clipboard fired no paste event). */
  function clearPlainPaste(): void {
    plainPasteNext = false;
  }

  /** Ends a clipboard session: chip hidden, flags and payload cleared. */
  function resetClipboard(): void {
    plainPasteNext = false;
    dismissPasteChip();
  }

  /** The last copy/cut snapshot — the context menu's paste reads it to
   * recognise its own copy (the async Clipboard API cannot see the private
   * MIME type) and re-insert the fragment with exact fidelity. */
  function getLastCopy(): ClipboardSnapshot | null {
    return lastCopy;
  }

  onSessionClose(resetClipboard);

  return {
    pasteChip,
    handleCopy,
    handleCut,
    handlePaste,
    handlePasteOption,
    dismissPasteChip,
    armPlainPaste,
    clearPlainPaste,
    resetClipboard,
    performPaste,
    getLastCopy,
  };
}
