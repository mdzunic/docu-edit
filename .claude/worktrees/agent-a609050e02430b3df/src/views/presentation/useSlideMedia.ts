// The Media panel editing session (SPEC-095) — the presentation view's wiring
// for the audio/video side panel and the stage's inline playback: the persisted
// open/closed state (a `presentationMediaOpen` preference), the selected media
// frame's modelled flags (start mode, loop, mute, volume, full screen), the
// "Replace poster…" picker, and the editor's play/stop activation (nothing
// auto-plays here; activating the badge swaps in native controls, and changing
// slides or deselecting stops it). Each flag edit sequences record → mutate →
// commit exactly like the Transitions/Animations panels; the model mutations
// live framework-free in frames.ts.
import { computed, type ComputedRef, type Ref } from "vue";
import type { MediaFrame, PresentationModel, Slide } from "../../lib/presentation/model";
import { editMediaFrame, setMediaPoster } from "../../lib/presentation/frames";
import { bytesToBase64 } from "../../lib/core/packageParts";
import type { MediaPart } from "../../lib/core/media";
import { PRESENTATION_IMAGE_ACCEPT, PRESENTATION_IMAGE_MIME_TYPES } from "./useFrameInsertion";
import { usePreferences } from "../../composables/usePreferences";

export interface MediaDeps {
  model: Ref<PresentationModel | null>;
  currentSlide: ComputedRef<Slide | null>;
  currentIndex: Ref<number>;
  record: (label?: string) => void;
  commit: () => void;
  /** Marks `model.media` changed so `commit` re-slices it (AC-13/31). */
  bumpMedia: () => void;
  selectedFrameId: Ref<string | null>;
  selectFrame: (id: string | null) => void;
  reportError: (message: string) => void;
}

export function useSlideMedia(deps: MediaDeps) {
  const { preferences, setPreference } = usePreferences();

  const open = computed(() => preferences.presentationMediaOpen);
  function toggle(): void {
    setPreference("presentationMediaOpen", !open.value);
  }

  /** The selected media frame, or null when the selection is not a clip. */
  const frame = computed<MediaFrame | null>(() => {
    const slide = deps.currentSlide.value;
    const id = deps.selectedFrameId.value;
    if (!slide || !id) {
      return null;
    }
    const found = slide.frames.find((f) => f.id === id);
    return found && found.kind === "media" ? found : null;
  });
  const isMedia = computed(() => frame.value !== null);
  const slideNumber = computed(() => deps.currentIndex.value + 1);

  /** record → mutate the selected clip's flags → commit, one undo entry. */
  function edit(label: string, mutate: (f: MediaFrame) => void): void {
    const model = deps.model.value;
    const f = frame.value;
    if (!model || !f) {
      return;
    }
    deps.record(label);
    editMediaFrame(model, deps.currentIndex.value, f.id, mutate);
    deps.commit();
  }

  // --- modelled flags (value + setter per control) --------------------------

  const startAutomatically = computed(() => frame.value?.autoStart === true);
  function setStartMode(auto: boolean): void {
    edit("Media start mode", (f) => {
      if (auto) {
        f.autoStart = true;
      } else {
        delete f.autoStart;
      }
    });
  }

  const loop = computed(() => frame.value?.loop === true);
  function setLoop(on: boolean): void {
    edit("Media loop", (f) => {
      if (on) {
        f.loop = true;
      } else {
        delete f.loop;
      }
    });
  }

  const muted = computed(() => frame.value?.muted === true);
  function setMuted(on: boolean): void {
    edit("Media mute", (f) => {
      if (on) {
        f.muted = true;
      } else {
        delete f.muted;
      }
    });
  }

  const volume = computed(() => frame.value?.volume ?? 100);
  function setVolume(value: number): void {
    const clamped = Math.max(0, Math.min(100, Math.round(value)));
    edit("Media volume", (f) => {
      // Absent = 100 (full); store only a non-default level so a clip at full
      // volume stays byte-minimal.
      if (clamped === 100) {
        delete f.volume;
      } else {
        f.volume = clamped;
      }
    });
  }

  const fullScreen = computed(() => frame.value?.fullScreen === true);
  function setFullScreen(on: boolean): void {
    edit("Media full screen", (f) => {
      if (on) {
        f.fullScreen = true;
      } else {
        delete f.fullScreen;
      }
    });
  }

  // --- Replace poster… ------------------------------------------------------

  const posterAccept = PRESENTATION_IMAGE_ACCEPT;
  async function replacePoster(file: File): Promise<void> {
    const model = deps.model.value;
    const f = frame.value;
    if (!model || !f) {
      return;
    }
    const contentType = file.type;
    if (!(PRESENTATION_IMAGE_MIME_TYPES as readonly string[]).includes(contentType)) {
      deps.reportError("That file isn't a supported image (PNG, JPEG, GIF or WebP).");
      return;
    }
    let buffer: Uint8Array;
    try {
      buffer = new Uint8Array(await file.arrayBuffer());
    } catch {
      deps.reportError("That poster image couldn't be read.");
      return;
    }
    const poster: MediaPart = { mediaId: "", contentType, data: bytesToBase64(buffer) };
    deps.record("Replace poster");
    if (!setMediaPoster(model, deps.currentIndex.value, f.id, poster)) {
      deps.reportError("That poster couldn't be set.");
      return;
    }
    deps.bumpMedia();
    deps.commit();
  }

  return {
    open,
    toggle,
    frame,
    isMedia,
    slideNumber,
    startAutomatically,
    setStartMode,
    loop,
    setLoop,
    muted,
    setMuted,
    volume,
    setVolume,
    fullScreen,
    setFullScreen,
    posterAccept,
    replacePoster,
  };
}
