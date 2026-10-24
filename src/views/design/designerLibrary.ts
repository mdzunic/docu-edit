// What the deck designer (SPEC-115) needs from outside itself, as narrow
// interfaces the session takes — so useDeckDesigner.spec drives it against a
// stubbed gateway and an in-memory library — with the real ones here: the
// store's text documents and decks, and the model gateway, reached only
// through composables/useModelGateway.ts by an import() taken on the first
// request.
import type { ModelRequest, ModelResult } from "../../lib/ai/modelGateway";
import {
  getDocument,
  getPresentation,
  listDocuments,
  listTrashedDocuments,
  storePresentation,
  trashDocument,
} from "../../lib/documentStore";
import type { PresentationModel } from "../../lib/presentation/model";
import type { OutlineSource } from "../../lib/design/outlineDesigner";
import { outlineFromDeck, outlineFromDocument } from "./outlineSource";

/** The gateway calls the designer makes — `stream` only: the plan step
 * streams its reply so a slow local model shows progress. */
export interface DesignerGateway {
  stream(request: ModelRequest, onDelta: (chunk: string) => void): Promise<ModelResult>;
}

/** A library entry the New flow can take as its source. */
export interface DesignerSourceEntry {
  id: string;
  title: string;
  kind: "document" | "deck";
}

/** The library as the designer uses it: what may be a source (and its
 * outline), the titles a new deck is numbered against, storing and trashing. */
export interface DesignerLibrary {
  sources(): DesignerSourceEntry[];
  outlineOf(id: string): OutlineSource | undefined;
  titles(): string[];
  store(model: PresentationModel): void;
  trash(id: string): void;
}

/** The real library: the store's text documents and decks. Reads a source
 * and never writes to it. */
export function storeLibrary(): DesignerLibrary {
  return {
    sources() {
      return listDocuments()
        .filter((entry) => entry.kind === "text" || entry.kind === "presentation")
        .map((entry) => ({ id: entry.id, title: entry.title, kind: entry.kind === "presentation" ? "deck" : "document" }));
    },
    outlineOf(id) {
      const deck = getPresentation(id);
      if (deck) {
        return outlineFromDeck(deck);
      }
      const doc = getDocument(id);
      return doc ? outlineFromDocument(doc) : undefined;
    },
    titles() {
      // Every record's title, trashed ones too — as renameDocument numbers.
      return [...listDocuments(), ...listTrashedDocuments()].map((entry) => entry.title);
    },
    store: storePresentation,
    trash: trashDocument,
  };
}

/** The model gateway, loaded on demand. */
export async function loadModelGateway(): Promise<DesignerGateway> {
  const { useModelGateway } = await import("../../composables/useModelGateway");
  return useModelGateway();
}
