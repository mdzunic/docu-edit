<script setup lang="ts">
// The slides menu bar (SPEC-049's model, for the presentation product). The
// word processor has had one since SPEC-049 while the deck editor had only
// strips of buttons, so the same application offered two different ways to
// find a command — and the same action two names, "Download" here against
// "Export" there. This is a second, discoverable route to commands the strips
// already carry, not a parallel implementation: every item emits a command id
// that usePresentationMenu routes to the composable the strip calls.
//
// Structure and keyboard behaviour come from MenuBar, which is product-neutral
// (it renders titles and popups and owns the roving focus; the contents are
// slots). That is also what gives the deck the compact-tier menu sheet for
// free. Runtime prop declarations — see AppShell.vue on the espree constraint.
import MenuBar from "../MenuBar.vue";
import MenuItem from "../MenuItem.vue";
import MenuCheckItem from "../MenuCheckItem.vue";
import MenuSubmenu from "../MenuSubmenu.vue";
import { computed, useTemplateRef } from "vue";
import { asPresentationMenuState } from "../../lib/presentation/menuState";

const props = defineProps({
  /** One object rather than a dozen props — see PresentationMenuState. */
  state: { type: Object, default: () => ({}) },
});

const emit = defineEmits(["command", "insert-image-file", "insert-media-file"]);

// A computed, not a snapshot: the prop changes as panels open and the
// selection changes, and the template auto-unwraps this so items stay `s.x`.
const s = computed(() => asPresentationMenuState(props.state));

const MENUS = [
  { id: "file", label: "File" },
  { id: "edit", label: "Edit" },
  { id: "view", label: "View" },
  { id: "insert", label: "Insert" },
  { id: "format", label: "Format" },
  { id: "slide", label: "Slide" },
];

// The shape palette the Insert strip offers, trimmed to the ones worth naming
// in a menu; the strip keeps the full gallery.
const SHAPES = [
  { id: "rect", label: "Rectangle" },
  { id: "roundRect", label: "Rounded rectangle" },
  { id: "ellipse", label: "Ellipse" },
  { id: "triangle", label: "Triangle" },
  { id: "diamond", label: "Diamond" },
  { id: "arrowRight", label: "Arrow" },
  { id: "star5", label: "Star" },
  { id: "line", label: "Line" },
];

const CHARTS = [
  { id: "bar", label: "Bar" },
  { id: "column", label: "Column" },
  { id: "line", label: "Line" },
  { id: "pie", label: "Pie" },
];

const ALIGNMENTS = [
  { id: "left", label: "Left" },
  { id: "center", label: "Centre" },
  { id: "right", label: "Right" },
  { id: "justify", label: "Justified" },
];

const ARRANGE = [
  { id: "front", label: "Bring to front" },
  { id: "forward", label: "Bring forward" },
  { id: "backward", label: "Send backward" },
  { id: "back", label: "Send to back" },
];

// Image/Media open the OS picker from the menu, exactly as the word
// processor's Insert ▸ Image… does. Hidden from the tab order and the
// accessibility tree: the menu item is the accessible control (see
// src/fileInputs.spec.ts, which enforces that on every picker).
const imageInput = useTemplateRef("imageInput");
const mediaInput = useTemplateRef("mediaInput");

// Default values rather than annotations: ESLint parses .vue scripts with
// espree, which rejects TS-only syntax (see AppShell.vue).
function pickedFile(event = new Event("change")) {
  const target = event.target;
  const file = target instanceof HTMLInputElement ? target.files?.[0] : undefined;
  if (target instanceof HTMLInputElement) {
    // Cleared so picking the same file twice still fires a change event.
    target.value = "";
  }
  return file;
}

function onImagePicked(event = new Event("change")) {
  const file = pickedFile(event);
  if (file) {
    emit("insert-image-file", file);
  }
}

function onMediaPicked(event = new Event("change")) {
  const file = pickedFile(event);
  if (file) {
    emit("insert-media-file", file);
  }
}
</script>

<template>
  <MenuBar
    :menus="MENUS"
    aria-label="Presentation menu"
  >
    <template #file>
      <MenuSubmenu
        label="Export"
        testid="slides-file-menu-export"
      >
        <MenuItem
          label="PowerPoint (.pptx)"
          testid="slides-export-native"
          @select="emit('command', 'file.export')"
        />
        <MenuItem
          v-if="s.otherFormatLabel"
          :label="s.otherFormatLabel"
          testid="slides-export-copy"
          @select="emit('command', 'file.export-copy')"
        />
        <MenuItem
          label="Web page (.html)"
          testid="slides-export-html"
          @select="emit('command', 'file.export-html')"
        />
      </MenuSubmenu>
      <hr class="menu-separator">
      <MenuItem
        label="Deck setup…"
        testid="slides-menu-deck-setup"
        @select="emit('command', 'file.deck-setup')"
      />
      <MenuItem
        label="Preferences…"
        testid="slides-menu-preferences"
        @select="emit('command', 'app.preferences')"
      />
      <hr class="menu-separator">
      <MenuItem
        label="Close"
        testid="slides-menu-close"
        @select="emit('command', 'file.close')"
      />
    </template>

    <template #edit>
      <MenuItem
        label="Undo"
        testid="slides-menu-undo"
        :disabled="!s.canUndo"
        @select="emit('command', 'edit.undo')"
      />
      <MenuItem
        label="Redo"
        testid="slides-menu-redo"
        :disabled="!s.canRedo"
        @select="emit('command', 'edit.redo')"
      />
    </template>

    <template #view>
      <MenuCheckItem
        label="Design"
        testid="slides-menu-design"
        :checked="s.designOpen"
        @select="emit('command', 'view.design')"
      />
      <MenuCheckItem
        label="Transition"
        testid="slides-menu-transition"
        :checked="s.transitionOpen"
        @select="emit('command', 'view.transition')"
      />
      <MenuCheckItem
        label="Animate"
        testid="slides-menu-animate"
        :checked="s.animateOpen"
        @select="emit('command', 'view.animate')"
      />
      <MenuCheckItem
        label="Layout"
        testid="slides-menu-layout"
        :checked="s.layoutOpen"
        @select="emit('command', 'view.layout')"
      />
      <MenuCheckItem
        label="Table"
        testid="slides-menu-table-panel"
        :checked="s.tableOpen"
        @select="emit('command', 'view.table')"
      />
      <MenuCheckItem
        label="Media"
        testid="slides-menu-media-panel"
        :checked="s.mediaOpen"
        @select="emit('command', 'view.media')"
      />
      <MenuCheckItem
        label="Header &amp; footer"
        testid="slides-menu-header-footer"
        :checked="s.headerFooterOpen"
        @select="emit('command', 'view.header-footer')"
      />
      <hr class="menu-separator">
      <MenuCheckItem
        label="Speaker notes"
        testid="slides-menu-notes"
        :checked="s.notesOpen"
        @select="emit('command', 'view.notes')"
      />
      <hr class="menu-separator">
      <MenuItem
        label="Start slideshow"
        testid="slides-menu-present"
        :disabled="!s.canPresent"
        @select="emit('command', 'view.present')"
      />
    </template>

    <template #insert>
      <MenuItem
        label="Text box"
        testid="slides-menu-text-box"
        @select="emit('command', 'insert.text-box')"
      />
      <MenuItem
        label="Image…"
        testid="slides-menu-image"
        @select="imageInput?.click()"
      />
      <MenuItem
        label="Audio or video…"
        testid="slides-menu-media"
        @select="mediaInput?.click()"
      />
      <hr class="menu-separator">
      <MenuSubmenu
        label="Shape"
        testid="slides-menu-shape"
      >
        <MenuItem
          v-for="shape in SHAPES"
          :key="shape.id"
          :label="shape.label"
          :testid="`slides-menu-shape-${shape.id}`"
          @select="emit('command', `insert.shape:${shape.id}`)"
        />
      </MenuSubmenu>
      <MenuSubmenu
        label="Chart"
        testid="slides-menu-chart"
      >
        <MenuItem
          v-for="chart in CHARTS"
          :key="chart.id"
          :label="chart.label"
          :testid="`slides-menu-chart-${chart.id}`"
          @select="emit('command', `insert.chart:${chart.id}`)"
        />
      </MenuSubmenu>
    </template>

    <template #format>
      <MenuItem
        label="Bold"
        testid="slides-menu-bold"
        @select="emit('command', 'format.bold')"
      />
      <MenuItem
        label="Italic"
        testid="slides-menu-italic"
        @select="emit('command', 'format.italic')"
      />
      <MenuItem
        label="Underline"
        testid="slides-menu-underline"
        @select="emit('command', 'format.underline')"
      />
      <hr class="menu-separator">
      <MenuSubmenu
        label="Align text"
        testid="slides-menu-align-text"
      >
        <MenuItem
          v-for="alignment in ALIGNMENTS"
          :key="alignment.id"
          :label="alignment.label"
          :testid="`slides-menu-align-${alignment.id}`"
          @select="emit('command', `format.align:${alignment.id}`)"
        />
      </MenuSubmenu>
      <MenuItem
        label="Bulleted list"
        testid="slides-menu-bullet"
        @select="emit('command', 'format.bullet')"
      />
      <MenuItem
        label="Increase indent"
        testid="slides-menu-indent"
        @select="emit('command', 'format.indent')"
      />
      <MenuItem
        label="Decrease indent"
        testid="slides-menu-outdent"
        @select="emit('command', 'format.outdent')"
      />
    </template>

    <template #slide>
      <MenuItem
        label="New slide"
        testid="slides-menu-new-slide"
        @select="emit('command', 'insert.slide')"
      />
      <hr class="menu-separator">
      <MenuSubmenu
        label="Arrange"
        testid="slides-menu-arrange"
      >
        <MenuItem
          v-for="order in ARRANGE"
          :key="order.id"
          :label="order.label"
          :testid="`slides-menu-arrange-${order.id}`"
          :disabled="!s.hasSelection"
          @select="emit('command', `arrange:${order.id}`)"
        />
      </MenuSubmenu>
      <MenuItem
        label="Group"
        testid="slides-menu-group"
        :disabled="!s.canGroup"
        @select="emit('command', 'format.group')"
      />
      <MenuItem
        label="Ungroup"
        testid="slides-menu-ungroup"
        :disabled="!s.canUngroup"
        @select="emit('command', 'format.ungroup')"
      />
      <hr class="menu-separator">
      <MenuItem
        label="Position and size…"
        testid="slides-menu-position"
        :disabled="!s.hasSelection"
        @select="emit('command', 'format.position')"
      />
    </template>
  </MenuBar>

  <input
    ref="imageInput"
    type="file"
    tabindex="-1"
    aria-hidden="true"
    class="menu-file-input"
    data-testid="slides-menu-image-input"
    accept="image/png,image/jpeg,image/gif,image/svg+xml,image/webp"
    @change="onImagePicked"
  >
  <input
    ref="mediaInput"
    type="file"
    tabindex="-1"
    aria-hidden="true"
    class="menu-file-input"
    data-testid="slides-menu-media-input"
    accept="audio/*,video/*"
    @change="onMediaPicked"
  >
</template>

<style scoped>
.menu-file-input {
  display: none;
}
</style>
