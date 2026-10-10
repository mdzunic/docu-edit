// The browser side of the cell clipboard (SPEC-125 "Flavours"): reading the
// async Clipboard API — what the cell menu's Paste and Paste special can
// reach — and finding the payload attribute in a text/html flavour. Copy,
// Cut and the keyboard's Paste go through the synchronous ClipboardEvent
// handlers in useSheetClipboard instead.
//
// - readClipboard asks for the raw HTML (`unsanitized: ["text/html"]`, where
//   the browser supports it — elsewhere the option is ignored) and returns
//   the text/html and text/plain flavours, or null when the API is missing
//   or the read is refused.
// - payloadAttribute parses the HTML on a detached DOMParser document (so
//   nothing in it executes or loads) and returns the attribute of its
//   `table[data-docu-edit-cells]`, or null.
import { PAYLOAD_ATTRIBUTE } from "../../lib/spreadsheet/clipboard";

export interface ClipboardFlavours {
  html: string;
  text: string;
}

/** The async API's read with Chromium's raw-HTML option (unknown options
 * are ignored where unsupported). */
interface RawRead {
  read(options?: { unsanitized?: string[] }): Promise<ClipboardItems>;
}

/** The clipboard's text/html and text/plain (see the module header). */
export async function readClipboard(clipboard: Clipboard | undefined): Promise<ClipboardFlavours | null> {
  if (!clipboard) {
    return null;
  }
  try {
    if (typeof clipboard.read === "function") {
      const items = await (clipboard as unknown as RawRead).read({ unsanitized: ["text/html"] });
      const flavours: ClipboardFlavours = { html: "", text: "" };
      for (const item of items) {
        for (const type of item.types) {
          if (type === "text/html" && flavours.html === "") {
            flavours.html = await (await item.getType(type)).text();
          } else if (type === "text/plain" && flavours.text === "") {
            flavours.text = await (await item.getType(type)).text();
          }
        }
      }
      return flavours;
    }
    if (typeof clipboard.readText === "function") {
      return { html: "", text: await clipboard.readText() };
    }
  } catch {
    return null;
  }
  return null;
}

/** The payload attribute of the HTML's `table[data-docu-edit-cells]`, or
 * null (see the module header). */
export function payloadAttribute(html: string): string | null {
  if (!html.includes(PAYLOAD_ATTRIBUTE) || typeof DOMParser === "undefined") {
    return null;
  }
  const table = new DOMParser().parseFromString(html, "text/html").querySelector(`table[${PAYLOAD_ATTRIBUTE}]`);
  return table?.getAttribute(PAYLOAD_ATTRIBUTE) ?? null;
}

let cuts = 0;

/** A fresh cut token: which cut a pasted payload is, in this view. */
export function newCutToken(): string {
  cuts += 1;
  const random = typeof crypto !== "undefined" && typeof crypto.randomUUID === "function" ? crypto.randomUUID() : String(Math.random()).slice(2);
  return `cut-${cuts}-${random}`;
}
