<script setup lang="ts">
import { nextTick, onMounted, useTemplateRef } from "vue";

// The shared modal shell — backdrop, centred panel, title and actions row —
// extracted from the near-identical copies that grew one per dialog. The base
// owns dismissal (backdrop click, Escape anywhere inside the panel) and
// keyboard containment (initial focus, a Tab/Shift+Tab trap); the wrapping
// dialog owns its body, its buttons and its outward emit names, mapping
// @dismiss onto whatever it emitted before (cancel or close) so callers never
// change. Runtime prop declarations, not defineProps<{...}>() generics — see
// AppShell.vue's comment on the same espree constraint.
const props = defineProps({
  /** Rendered as the h2 title and — through aria-labelledby — the dialog's
   * accessible name. */
  title: { type: String, required: true },
  /** data-testid, applied to the backdrop (where every dialog carried it);
   * also prefixes the title's element id (`${testid}-title`). */
  testid: { type: String, required: true },
  /** The panel's min-width — the one knob the per-dialog .dialog copies
   * actually varied (max-width was a constant 90vw). 20rem is the most
   * common value. */
  width: { type: String, default: "20rem" },
  /** Optional data-testid applied to the h2 title itself (in addition to
   * `testid` on the backdrop) — ConversionReportDialog's spec asserts one
   * on its title. "" (the default) renders no attribute at all. */
  titleTestid: { type: String, default: "" },
});

const emit = defineEmits(["dismiss"]);

const panel = useTemplateRef("panel");

// The standard focusable-elements selector, trimmed to what dialogs contain.
const FOCUSABLE =
  'a[href], button:not([disabled]), input:not([disabled]), ' +
  'select:not([disabled]), textarea:not([disabled]), ' +
  '[tabindex]:not([tabindex="-1"])';

function focusables() {
  return panel.value ? Array.from(panel.value.querySelectorAll(FOCUSABLE)) : [];
}

// Initial focus: the [autofocus] element if the dialog marked one (migrated
// dialogs put the attribute on whatever their own onMounted used to focus),
// else the panel itself so Escape and the Tab trap work immediately.
onMounted(async () => {
  await nextTick();
  const marked = panel.value?.querySelector("[autofocus]");
  const target = marked instanceof HTMLElement ? marked : panel.value;
  target?.focus();
});

function handleKeydown(event = new KeyboardEvent("keydown")) {
  if (event.key === "Escape") {
    // Stop propagation, as the strictest existing dialogs did — no global
    // Escape handler should also react while a modal is up.
    event.stopPropagation();
    emit("dismiss");
    return;
  }
  if (event.key !== "Tab") {
    return;
  }
  // Focus trap: Tab wraps from the last focusable to the first, Shift+Tab
  // from the first (or from the panel itself, right after fallback focus)
  // to the last.
  const items = focusables();
  const first = items[0];
  const last = items[items.length - 1];
  if (!(first instanceof HTMLElement) || !(last instanceof HTMLElement)) {
    event.preventDefault();
    return;
  }
  const active = document.activeElement;
  if (event.shiftKey && (active === first || active === panel.value)) {
    event.preventDefault();
    last.focus();
  } else if (!event.shiftKey && active === last) {
    event.preventDefault();
    first.focus();
  }
}
</script>

<template>
  <!-- Deliberately NOT teleported in v1: the dialog unit specs find their
       elements through wrapper-scoped queries that a Teleport would escape,
       for zero user-visible gain today — dialogs already mount near the view
       root, so the backdrop's z-index 50 wins as-is. When Teleport does
       become worth it, it is a one-line change in this one file — which is
       the point of having a base. -->
  <div
    class="dialog-backdrop"
    :data-testid="props.testid"
    @click.self="emit('dismiss')"
  >
    <div
      ref="panel"
      class="dialog"
      role="dialog"
      aria-modal="true"
      :aria-labelledby="`${props.testid}-title`"
      tabindex="-1"
      :style="{ '--dialog-width': props.width }"
      @keydown="handleKeydown"
    >
      <!-- Header-row variant: some dialogs (mail merge, special characters,
           the template gallery) put controls — a ✕ close button, an import
           button — on the title line. The optional #header-extra slot opts
           into a flex header row holding the h2 and that content; without
           it the h2 renders exactly as before, so plain dialogs see no
           structural change. -->
      <div
        v-if="$slots['header-extra']"
        class="dialog-header"
      >
        <h2
          :id="`${props.testid}-title`"
          class="dialog-title"
          :data-testid="props.titleTestid || undefined"
        >
          {{ props.title }}
        </h2>
        <slot name="header-extra" />
      </div>
      <h2
        v-else
        :id="`${props.testid}-title`"
        class="dialog-title"
        :data-testid="props.titleTestid || undefined"
      >
        {{ props.title }}
      </h2>
      <slot />
      <div
        v-if="$slots.actions"
        class="dialog-actions"
      >
        <slot name="actions" />
      </div>
    </div>
  </div>
</template>

<style scoped>
.dialog-backdrop {
  position: fixed;
  inset: 0;
  background: rgb(0 0 0 / 0.4);
  display: flex;
  align-items: center;
  justify-content: center;
  z-index: 50;
}

.dialog {
  background: #fff;
  border-radius: 6px;
  box-shadow: 0 4px 16px rgb(0 0 0 / 0.3);
  padding: 1.25rem 1.5rem;
  min-width: var(--dialog-width, 20rem);
  max-width: 90vw;
  box-sizing: border-box;
}

.dialog-title {
  margin: 0 0 1rem;
  font-size: 1rem;
  font-weight: 600;
  color: #202124;
}

/* The #header-extra variant: title left, the dialog's own controls right.
   The row carries the title's usual 1rem bottom spacing (dialogs that had
   a tighter header override it via :deep). */
.dialog-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-bottom: 1rem;
}

.dialog-header .dialog-title {
  margin: 0;
}

.dialog-actions {
  display: flex;
  justify-content: flex-end;
  gap: 0.5rem;
  margin-top: 1.25rem;
}
</style>
