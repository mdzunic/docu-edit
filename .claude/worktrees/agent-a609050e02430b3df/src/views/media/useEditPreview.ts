// The image editor's on-screen preview (SPEC-112): renderEditedImage — the
// same function every export calls — run over the session's downscaled
// buffer at the stage's size, then painted into the visible canvas. Renders
// are coalesced to one per frame and a stale result (the stack moved on
// while it rendered) is dropped, never painted. `scale` maps output pixels
// to the canvas's CSS pixels for the overlays (crop handles, drawing).
import { computed, onScopeDispose, ref, watch, type Ref } from "vue";
import { resolveEditGeometry } from "../../lib/core/imageEdits";
import { previewTargetSize } from "../../lib/media/renderEdits";
import type { ImageEditor } from "./useImageEditor";

/** The stage size assumed before it is measured (and under jsdom). */
const FALLBACK_BOX = { widthPx: 800, heightPx: 600 };

export function useEditPreview(editor: ImageEditor, canvas: Ref<HTMLCanvasElement | null>) {
  const box = ref({ ...FALLBACK_BOX });
  /** How many renders have been painted — a test hook for "the preview
   * caught up". */
  const painted = ref(0);
  const geometry = computed(() => {
    const source = editor.preview.value;
    return source ? resolveEditGeometry(editor.previewStack.value, source.naturalPx) : null;
  });
  const size = computed(() =>
    geometry.value
      ? previewTargetSize({ widthPx: geometry.value.widthPx, heightPx: geometry.value.heightPx }, box.value)
      : { widthPx: 1, heightPx: 1 },
  );
  /** Canvas CSS pixels per output pixel. */
  const scale = computed(() => (geometry.value ? size.value.widthPx / geometry.value.widthPx : 1));

  let frame = 0;
  let token = 0;

  async function render(): Promise<void> {
    frame = 0;
    const source = editor.preview.value;
    const target = canvas.value;
    if (!source || !target) return;
    const mine = ++token;
    const want = size.value;
    try {
      const result = await editor.steps.render(source, editor.previewStack.value, want);
      if (mine === token) {
        target.width = want.widthPx;
        target.height = want.heightPx;
        target.getContext("2d")?.drawImage(result, 0, 0);
        painted.value += 1;
      }
      editor.steps.release(result);
    } catch {
      // A preview that cannot draw leaves the last good frame up; the
      // full-resolution actions report their own failures.
    }
  }

  function schedule(): void {
    if (frame !== 0) return;
    frame =
      typeof requestAnimationFrame === "function"
        ? requestAnimationFrame(() => void render())
        : (setTimeout(() => void render(), 0) as unknown as number);
  }

  function setBox(widthPx: number, heightPx: number): void {
    if (widthPx > 0 && heightPx > 0) {
      box.value = { widthPx: Math.floor(widthPx), heightPx: Math.floor(heightPx) };
    }
  }

  watch([() => editor.preview.value, () => editor.previewStack.value, box, canvas], schedule, { immediate: true });
  onScopeDispose(() => {
    token += 1;
    if (frame !== 0 && typeof cancelAnimationFrame === "function") cancelAnimationFrame(frame);
  });

  return { geometry, size, scale, painted, setBox };
}
