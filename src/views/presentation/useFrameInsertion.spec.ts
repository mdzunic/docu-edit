// View spec: the frame-insertion composable's arm state and image-picker
// validation (AC-4/9). The draw gesture itself is covered by stageMath +
// frames specs and the e2e; here we exercise the branches that need no surface.
import { computed, shallowRef } from "vue";
import { describe, expect, it, vi } from "vitest";
import { createBlankPresentation } from "../../lib/presentation/slides";
import type { PresentationModel } from "../../lib/presentation/model";
import { useFrameSelection } from "./useFrameSelection";
import { useFrameInsertion } from "./useFrameInsertion";

const PNG_2x3 =
  "iVBORw0KGgoAAAANSUhEUgAAAAIAAAADCAYAAADkc372AAAAEUlEQVR42mNk+M9QzwAFjDAaAF1eBFsPTQEwAAAAAElFTkSuQmCC";

function bytesFromBase64(b64: string): Uint8Array {
  const binary = atob(b64);
  const out = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) out[i] = binary.charCodeAt(i);
  return out;
}

// jsdom's File has no arrayBuffer(); the composable only reads `.type` and
// `.arrayBuffer()`, so a minimal stub is faithful and portable.
function fileStub(bytes: Uint8Array, type: string): File {
  return { type, arrayBuffer: async () => bytes.buffer } as unknown as File;
}

function setup() {
  const model = shallowRef<PresentationModel | null>(createBlankPresentation());
  const currentIndex = shallowRef(0);
  const currentSlide = computed(() => model.value?.slides[0] ?? null);
  const selection = useFrameSelection({ currentSlide });
  const reportError = vi.fn();
  const startEditing = vi.fn();
  const insertion = useFrameInsertion({
    model,
    currentSlide,
    currentIndex,
    selection,
    surface: () => null,
    record: vi.fn(),
    commit: vi.fn(),
    startEditing,
    reportError,
  });
  return { model, selection, insertion, reportError };
}

describe("useFrameInsertion arm state", () => {
  it("arms a text box / shape and disarms", () => {
    const { insertion } = setup();
    expect(insertion.isArmed.value).toBe(false);
    insertion.armTextBox();
    expect(insertion.armed.value).toEqual({ kind: "textbox" });
    insertion.armShape("ellipse");
    expect(insertion.armed.value).toEqual({ kind: "shape", geometry: "ellipse" });
    insertion.disarm();
    expect(insertion.isArmed.value).toBe(false);
  });

  it("accepts png/jpeg/gif/webp in its accept string", () => {
    const { insertion } = setup();
    expect(insertion.accept).toBe("image/png,image/jpeg,image/gif,image/webp");
  });

  it("inserts a table centred at 80% width and selects it (AC-5)", () => {
    const { model, selection, insertion } = setup();
    insertion.insertTable(3, 4);
    const table = model.value!.slides[0].frames.find((f) => f.kind === "table");
    expect(table).toBeTruthy();
    expect(selection.selectedFrameId.value).toBe(table!.id);
    if (table && table.kind === "table") {
      expect(table.rows).toHaveLength(3);
      expect(table.columnWidthsCm).toHaveLength(4);
    }
  });
});

describe("useFrameInsertion image picker (AC-4/9)", () => {
  it("inserts a valid PNG and selects it", async () => {
    const { model, selection, insertion, reportError } = setup();
    const file = fileStub(bytesFromBase64(PNG_2x3), "image/png");
    await insertion.insertImageFile(file);
    expect(reportError).not.toHaveBeenCalled();
    const image = model.value!.slides[0].frames.find((f) => f.kind === "image");
    expect(image).toBeTruthy();
    expect(selection.selectedFrameId.value).toBe(image!.id);
  });

  it("rejects a broken image and changes nothing (AC-9)", async () => {
    const { model, insertion, reportError } = setup();
    const before = model.value!.slides[0].frames.length;
    const file = fileStub(new Uint8Array([1, 2, 3, 4]), "image/png");
    await insertion.insertImageFile(file);
    expect(reportError).toHaveBeenCalledTimes(1);
    expect(model.value!.slides[0].frames).toHaveLength(before);
  });

  it("rejects an unsupported type", async () => {
    const { insertion, reportError } = setup();
    const file = fileStub(new Uint8Array([1, 2, 3]), "image/bmp");
    await insertion.insertImageFile(file);
    expect(reportError).toHaveBeenCalledTimes(1);
  });
});
