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
// free. It renders `embedded`: the deck gives its menu no row of its own —
// SlideRibbon hosts the titles at the head of the one chrome row, so the bar
// draws no ground or border (see SlideRibbon on why the deck merged them).
// Runtime prop declarations — see AppShell.vue on the espree constraint.
import MenuBar from "../MenuBar.vue";
import MenuItem from "../MenuItem.vue";
import MenuCheckItem from "../MenuCheckItem.vue";
import MenuSubmenu from "../MenuSubmenu.vue";
import ExperimentGate from "../ExperimentGate.vue";
import { computed } from "vue";
import { asPresentationMenuState } from "../../lib/presentation/menuState";

// SPEC-115: Slide › "Redesign…" is the deck designer's, so it renders — and
// its code loads — only while the deck-designer experiment is on.
const loadRedesignItem = () => import("../../views/design/RedesignMenuItem.vue");

const props = defineProps({
  /** One object rather than a dozen props — see PresentationMenuState. */
  state: { type: Object, default: () => ({}) },
});

const emit = defineEmits(["command"]);

// A computed, not a snapshot: the prop changes as panels open and the
// selection changes, and the template auto-unwraps this so items stay `s.x`.
const s = computed(() => asPresentationMenuState(props.state));

// Four menus, not six. "Insert" and "Format" are the ribbon's two strip tabs,
// which sit in the same row — so the bar used to offer INSERT twice, a menu
// and a tab, in the same condensed uppercase face a thumb's width apart, and
// FORMAT beside a Home tab holding the identical commands. Every command the
// two menus carried is on the strip that replaced them: Bold/Italic/Underline,
// the four alignments, bullets and both indents on Home; text box, image,
// media, the shape gallery and the charts on Insert — each of which offers
// MORE than the menu did (the full shape gallery, tables, arrange, align,
// group, style). So the deck's bar is deliberately shorter than the word
// processor's, which has no ribbon to absorb them.
const MENUS = [
  { id: "file", label: "File" },
  { id: "edit", label: "Edit" },
  { id: "view", label: "View" },
  { id: "slide", label: "Slide" },
];

const ARRANGE = [
  { id: "front", label: "Bring to front" },
  { id: "forward", label: "Bring forward" },
  { id: "backward", label: "Send backward" },
  { id: "back", label: "Send to back" },
];
</script>

<template>
  <MenuBar
    :menus="MENUS"
    aria-label="Presentation menu"
    embedded
  >
    <template #file>
      <MenuSubmenu
        label="Export"
        icon="download"
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
        icon="settings"
        testid="slides-menu-preferences"
        @select="emit('command', 'app.preferences')"
      />
      <hr class="menu-separator">
      <MenuItem
        label="Close"
        icon="close"
        testid="slides-menu-close"
        @select="emit('command', 'file.close')"
      />
    </template>

    <template #edit>
      <MenuItem
        label="Undo"
        icon="undo"
        testid="slides-menu-undo"
        :disabled="!s.canUndo"
        @select="emit('command', 'edit.undo')"
      />
      <MenuItem
        label="Redo"
        icon="redo"
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

    <template #slide>
      <MenuItem
        label="New slide"
        testid="slides-menu-new-slide"
        @select="emit('command', 'insert.slide')"
      />
      <ExperimentGate
        experiment="deck-designer"
        :load="loadRedesignItem"
        :disabled="!s.canRedesign"
        @select="emit('command', 'slide.redesign')"
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
</template>
