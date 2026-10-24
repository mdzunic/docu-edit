<script setup lang="ts">
// The assembly workspace at /assemble (SPEC-109): a left tray of the chosen
// images in order, the centre preview of the sheet at fit-to-window, a right
// inspector (the layout controls above SPEC-108's SheetSetupPanel) and a
// bottom output bar (format, estimate, Save to library, Download). A wiring
// hub like the other views: the workspace's logic is useAssembly's (tray,
// plan, persistence, output) and useAssemblyPreview's, and the surfaces are
// the components/assembly/* components. It takes no id and needs no kind
// guard; the library opens it through "Combine…" and "New › Combined
// image", each of which starts the plan it then restores.
import { computed, ref } from "vue";
import { useRouter } from "vue-router";
import AppShell from "../components/AppShell.vue";
import UiIconButton from "../components/ui/UiIconButton.vue";
import AssemblyTray from "../components/assembly/AssemblyTray.vue";
import AssemblyPreview from "../components/assembly/AssemblyPreview.vue";
import AssemblyInspector from "../components/assembly/AssemblyInspector.vue";
import AssemblyOutputBar from "../components/assembly/AssemblyOutputBar.vue";
import AssemblyPicker from "../components/assembly/AssemblyPicker.vue";
import { useTheme } from "../composables/useTheme";
import { useAssembly } from "./assembly/useAssembly";
import { useAssemblyPreview } from "./assembly/useAssemblyPreview";
import { asAssemblySteps } from "./assembly/useAssemblyOutput";
import { asPdfAssemblySteps } from "./assemble/usePdfAssembly";
import { MAX_ASSEMBLY_ITEMS } from "../lib/core/assemblyLayout";

const props = defineProps({
  /** Test seam: the browser-bound output steps (canvas, decoding). */
  steps: { type: Object, default: undefined },
  /** Test seam: the PDF output's steps (SPEC-110). */
  pdfSteps: { type: Object, default: undefined },
});

const router = useRouter();

// No menu bar here — the shared theme composable keeps the workspace themed
// by the app tokens and following the OS live (SPEC-059).
useTheme();

const assembly = useAssembly({ steps: asAssemblySteps(props.steps), pdfSteps: asPdfAssemblySteps(props.pdfSteps) });
// For a PDF the tray says how each image will be embedded.
const encodings = computed(() => (assembly.isPdf.value ? assembly.pdf.encodings.value : []));
const preview = useAssemblyPreview({ assembly, decode: asAssemblySteps(props.steps)?.decode });

const pickerOpen = ref(false);
const room = computed(() => MAX_ASSEMBLY_ITEMS - assembly.tray.entries.value.length);

function addFromLibrary(ids = [""]) {
  pickerOpen.value = false;
  assembly.tray.add(ids);
}

function close() {
  assembly.flush();
  void router.push("/");
}
</script>

<template>
  <AppShell document-title="Combine images">
    <template #app-bar-actions>
      <UiIconButton
        class="assembly-close"
        icon="close"
        label="Close and return to the library"
        data-testid="close-document"
        @click="close"
      />
    </template>
    <div
      class="assembly-workspace"
      data-testid="assembly-workspace"
    >
      <p
        v-if="assembly.notice.value"
        class="assembly-notice"
        data-testid="assembly-notice"
        role="status"
      >
        {{ assembly.notice.value }}
      </p>
      <aside
        class="assembly-left"
        aria-label="Tray"
      >
        <AssemblyTray
          :tray="assembly.tray"
          :encodings="encodings"
          @add="pickerOpen = true"
        />
      </aside>
      <div class="assembly-centre">
        <AssemblyPreview
          :assembly="assembly"
          :preview="preview"
        />
      </div>
      <aside
        class="assembly-right"
        aria-label="Inspector"
      >
        <AssemblyInspector :assembly="assembly" />
      </aside>
      <footer class="assembly-bottom">
        <AssemblyOutputBar :assembly="assembly" />
      </footer>
    </div>
    <AssemblyPicker
      v-if="pickerOpen"
      :room="room"
      @add="addFromLibrary"
      @close="pickerOpen = false"
    />
  </AppShell>
</template>

<style scoped>
/* On the brand bar the close control takes the on-brand foreground (the
   treatment the media viewer's Close gets). */
.assembly-close {
  color: var(--text-on-brand);
}

.assembly-close:not(:disabled):hover {
  color: var(--text-on-brand);
  background: color-mix(in srgb, var(--text-on-brand) 12%, transparent);
}

.assembly-workspace {
  height: 100%;
  min-height: 0;
  display: grid;
  grid-template-columns: 17rem minmax(0, 1fr) 22rem;
  grid-template-rows: auto minmax(0, 1fr) auto;
  grid-template-areas:
    "notice notice notice"
    "left centre right"
    "bottom bottom bottom";
  background: var(--surface-canvas);
}

.assembly-notice {
  grid-area: notice;
  margin: 0;
  padding: var(--space-2) var(--space-4);
  background: var(--surface-selected);
  color: var(--text-primary);
  font-size: var(--text-sm);
}

.assembly-left,
.assembly-right {
  min-height: 0;
  overflow-y: auto;
  padding: var(--space-4);
  background: var(--surface-base);
}

.assembly-left {
  grid-area: left;
  display: flex;
  flex-direction: column;
  border-right: var(--border-width) solid var(--border-subtle);
}

.assembly-right {
  grid-area: right;
  border-left: var(--border-width) solid var(--border-subtle);
}

.assembly-centre {
  grid-area: centre;
  min-width: 0;
  min-height: 0;
  display: flex;
  padding: var(--space-4);
}

.assembly-centre > * {
  flex: 1 1 auto;
}

.assembly-bottom {
  grid-area: bottom;
  padding: var(--space-3) var(--space-4);
  border-top: var(--border-width) solid var(--border-subtle);
  background: var(--surface-base);
}

@media (max-width: 64rem) {
  .assembly-workspace {
    height: auto;
    grid-template-columns: minmax(0, 1fr);
    grid-template-rows: auto;
    grid-template-areas:
      "notice"
      "centre"
      "left"
      "right"
      "bottom";
  }

  .assembly-left,
  .assembly-right {
    border: 0;
    border-top: var(--border-width) solid var(--border-subtle);
  }
}
</style>
