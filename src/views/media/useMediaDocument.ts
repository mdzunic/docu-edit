// The media viewer's load path (SPEC-107): resolves the routed id to its
// MediaModel, guards the kind in both directions (a text or presentation id
// hands off to its own view, mirroring the other views' guards), and owns
// the blob-URL lifecycle for the PDF embed — one URL.createObjectURL per
// open, revoked in onScopeDispose so a closed viewer never leaks the bytes.
import { onScopeDispose, shallowRef, watch, type ComputedRef, type Ref, ref } from "vue";
import type { Router } from "vue-router";
import { documentKindOf, getMedia, getStoredRecord, initLibrary } from "../../lib/documentStore";
import { editorPathFor } from "../../lib/documentKinds";
import type { MediaModel } from "../../lib/media/model";
import { mediaSourceBytes } from "../../lib/media/download";

export interface MediaDocumentDeps {
  mediaId: ComputedRef<string>;
  router: Router;
}

export interface MediaDocumentController {
  media: Ref<MediaModel | null>;
  /** blob: URL over the verbatim source bytes — the PDF embed's src and the
   * image viewer's, both served from memory, never re-encoded. */
  blobUrl: Ref<string>;
  /** Epoch ms the library record was last modified ("stored date"). */
  storedAt: Ref<number>;
  close: () => void;
}

export function useMediaDocument(deps: MediaDocumentDeps): MediaDocumentController {
  const media = shallowRef<MediaModel | null>(null);
  const blobUrl = ref("");
  const storedAt = ref(0);

  function revokeBlobUrl(): void {
    if (blobUrl.value) {
      URL.revokeObjectURL(blobUrl.value);
      blobUrl.value = "";
    }
  }

  function load(): void {
    const id = deps.mediaId.value;
    // Wrong-kind ids hand off to their own view, exactly as /doc/:id and
    // /slides/:id do (the route comes from the kind registry).
    const kind = documentKindOf(id);
    if (kind !== undefined && kind !== "media") {
      void deps.router.replace(editorPathFor(kind, id));
      return;
    }
    const model = getMedia(id) ?? null;
    storedAt.value = getStoredRecord(id)?.lastModified ?? 0;
    if (model !== null && model === media.value && blobUrl.value) {
      // The same document is already open (the post-initLibrary re-sync):
      // one object URL per open, so the embed never loads a URL that is
      // then revoked under it.
      return;
    }
    media.value = model;
    revokeBlobUrl();
    if (model) {
      const bytes = mediaSourceBytes(model);
      blobUrl.value = URL.createObjectURL(
        new Blob([bytes as unknown as BlobPart], { type: model.contentType }),
      );
    }
  }

  // main.ts awaits initLibrary() before mounting, so the re-sync after it
  // resolves matters only when the view is mounted directly (tests).
  load();
  void initLibrary().then(load);
  watch(deps.mediaId, load);

  onScopeDispose(revokeBlobUrl);

  function close(): void {
    // Mirrors the other views' close (SPEC-060): the library scrolls to and
    // highlights the row that was just closed.
    void deps.router.push({ path: "/", state: { closedDocId: deps.mediaId.value } });
  }

  return { media, blobUrl, storedAt, close };
}
