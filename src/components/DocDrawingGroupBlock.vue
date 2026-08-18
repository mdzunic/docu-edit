<script setup lang="ts">
import { asDrawingGroup, isDrawingGroup, isShape } from "../lib/documentModel";
import { groupDisplayStyle, groupMemberWrapperStyle, trackChangeBlockHidden } from "../lib/renderFormatting";
import { asTrackChangesDisplayMode } from "../lib/documentModel";
import { shapeClipPath } from "../lib/core/shapeGeometry";
import DocGroupMemberVisual from "./DocGroupMemberVisual.vue";

// One native drawing-group container on the page (SPEC-057 "Rendering and
// test surface"). Recursive: a member that is itself a container renders as a
// nested <DocDrawingGroupBlock member>. Paints each member's visual
// (DocGroupMemberVisual) and lays a transparent proxy over it — the container
// proxy carries role="group", each member proxy nests as
// `<base>-member-<j>` — the elements every selection click, keyboard
// traversal and e2e assertion goes through. The container's own box is not a
// hit surface (an empty region selects nothing), so only member proxies catch
// clicks; the container proxy is focus/ARIA only. Selection depth ("one level
// per click") is resolved by useShapes from the full member path this emits.
// Runtime prop declarations per the espree constraint (see DocImageBlock.vue).
defineOptions({ inheritAttrs: false });

const props = defineProps({
  /** The DrawingGroup block. */
  block: { type: Object, required: true },
  section: { type: Number, required: true },
  blockIndex: { type: Number, required: true },
  /** This drawing's 0-based index in document order (top-level only). */
  drawingIndex: { type: Number, default: 0 },
  /** The data-testid base — top-level derives `drawing-<index>`; a nested
   * container is passed `<parentBase>-member-<j>`. */
  testidBase: { type: String, default: "" },
  /** Nested-member mode: the parent has already positioned/sized this box, so
   * the anchor/positioning wrapper is skipped. */
  member: { type: Boolean, default: false },
  /** The selected drawing's path RELATIVE to this container: `[]` = this
   * container, `[j,...]` = a member, `null` = nothing here selected. */
  selectedPath: { type: Array, default: null },
  /** Whether a member proxy of THIS container is focusable — a member is
   * focusable only once its container is entered (SPEC-057 AC-40/41). */
  parentEntered: { type: Boolean, default: false },
  displayMode: { type: String, default: "all" },
  listDefinitions: { type: Array, default: () => [] },
  /** The document model — image members resolve their media from it. */
  model: { type: Object, default: null },
});

const emit = defineEmits(["select", "edit-text"]);

/** The selected path relative to this container (`[]` = container, `null` =
 * nothing here) as a real array or null — the prop is a plain Object/Array. */
function relativePath() {
  const path = props.selectedPath;
  return Array.isArray(path) ? path : null;
}

/** The container is entered (its direct members are focusable) when a member
 * of it, or something deeper, is selected. */
function entered() {
  const path = relativePath();
  return !!path && path.length >= 1;
}

/** This container's own proxy is selected when the relative path is empty. */
function containerSelected() {
  const path = relativePath();
  return !!path && path.length === 0;
}

function baseTestid() {
  return props.member ? props.testidBase : `drawing-${props.drawingIndex}`;
}

function memberTestid(index = 0) {
  return `${baseTestid()}-member-${index}`;
}

/** The selection path relative to member `index`, or undefined. */
function childSelectedPath(index = 0) {
  const path = relativePath();
  if (path && path.length >= 1 && path[0] === index) {
    return path.slice(1);
  }
  return undefined;
}

/** A leaf member is selected when the relative path is exactly `[index]`. */
function memberSelected(index = 0) {
  const child = childSelectedPath(index);
  return Array.isArray(child) && child.length === 0;
}

/** Whether THIS container's own proxy is focusable: a top-level container is
 * always tabbable; a nested container follows the member rule. */
function containerFocusable() {
  return props.member ? props.parentEntered : true;
}

const floatingWrap = ["square", "topAndBottom", "behind", "infront"];

function memberAt(index = 0) {
  return asDrawingGroup(props.block).members[index];
}

function memberWrapperStyle(index = 0) {
  return groupMemberWrapperStyle(memberAt(index));
}

/** A shape member's proxy clips to its painted region so a click on empty
 * interior falls through to whatever is beneath (SPEC-057, mirroring
 * DocShapeBlock's proxyStyle). Other kinds fill their rectangle. */
function memberProxyStyle(index = 0) {
  const member = memberAt(index);
  if (isShape(member)) {
    const clip = shapeClipPath(member.geometry, !!member.fill);
    return clip ? { clipPath: clip } : {};
  }
  return {};
}

function memberRole(index = 0) {
  const member = memberAt(index);
  if (isDrawingGroup(member)) {
    return "group";
  }
  if (isShape(member) && member.text) {
    return "group";
  }
  return "img";
}

function memberLabel(index = 0) {
  const member = memberAt(index);
  if (member.altText) {
    return member.altText;
  }
  if (isDrawingGroup(member)) {
    return `Group of ${member.members.length} drawings`;
  }
  return isShape(member) ? "Shape" : "";
}

function containerLabel() {
  const group = asDrawingGroup(props.block);
  return group.altText || `Group of ${group.members.length} drawings`;
}
</script>

<template>
  <span
    class="doc-group-outer"
    :class="{ 'doc-shape-anchor': !member, 'doc-shape-anchor-floating': !member && floatingWrap.includes(block.wrap) }"
    :data-section="!member ? section : undefined"
    :data-block="!member ? blockIndex : undefined"
  >
    <span
      v-show="member || !trackChangeBlockHidden(block.changes, asTrackChangesDisplayMode(displayMode))"
      class="doc-group"
      :class="{ 'doc-group-selected': containerSelected() }"
      :style="{ ...(member ? { position: 'absolute', inset: '0' } : groupDisplayStyle(asDrawingGroup(block))), '--group-rotation': `${block.rotationDeg ?? 0}deg` }"
    >
      <!-- Rotated content layer: the container turns about its box centre, so
           members (visuals + proxies) turn with it; the container proxy below
           stays unrotated as the measured/focus surface. -->
      <span class="doc-group-content">
        <span
          v-for="(m, j) in asDrawingGroup(block).members"
          :key="j"
          class="doc-group-member"
          :style="memberWrapperStyle(j)"
        >
          <DocDrawingGroupBlock
            v-if="isDrawingGroup(m)"
            member
            :block="m"
            :section="section"
            :block-index="blockIndex"
            :testid-base="memberTestid(j)"
            :selected-path="childSelectedPath(j)"
            :parent-entered="entered()"
            :display-mode="displayMode"
            :list-definitions="listDefinitions"
            :model="model"
            @select="(payload) => emit('select', { path: [j, ...payload.path], shiftKey: payload.shiftKey })"
            @edit-text="(payload) => emit('edit-text', { path: [j, ...payload.path] })"
          />
          <template v-else>
            <DocGroupMemberVisual
              :block="m"
              :model="model"
              :list-definitions="listDefinitions"
            />
            <span
              class="doc-group-member-proxy"
              :class="{ 'doc-group-member-selected': memberSelected(j) }"
              :style="memberProxyStyle(j)"
              :data-testid="memberTestid(j)"
              :role="memberRole(j)"
              :aria-label="memberLabel(j)"
              :tabindex="entered() ? 0 : -1"
              @click.stop="(event) => emit('select', { path: [j], shiftKey: event.shiftKey })"
              @dblclick.stop="emit('edit-text', { path: [j] })"
            />
          </template>
        </span>
      </span>
      <span
        class="doc-group-proxy"
        :data-testid="baseTestid()"
        role="group"
        :aria-label="containerLabel()"
        :tabindex="containerFocusable() ? 0 : -1"
      />
    </span>
  </span>
</template>

<style scoped>
.doc-group-outer {
  display: inline-block;
  line-height: 0;
}

.doc-shape-anchor-floating {
  position: relative;
}

.doc-group {
  display: block;
}

.doc-group-content {
  display: block;
  position: absolute;
  inset: 0;
  transform: rotate(var(--group-rotation, 0deg));
}

.doc-group-member {
  line-height: normal;
}

.doc-group-member-proxy {
  display: block;
  position: absolute;
  inset: 0;
  background: transparent;
  outline: none;
  cursor: pointer;
}

/* Selection frames turn with their drawing, exactly as DocShapeBlock's do:
   the proxy stays axis-aligned (it is the hit/measure surface), a rotated
   pseudo-element traces the outline. */
.doc-group-member-selected::after {
  content: "";
  position: absolute;
  inset: -3px;
  border: 2px solid #1a73e8;
  transform: rotate(var(--member-rotation, 0deg));
  pointer-events: none;
}

.doc-group-proxy {
  display: block;
  position: absolute;
  inset: 0;
  background: transparent;
  outline: none;
  /* The container box is not a hit surface (an empty region selects nothing);
     member proxies above it catch every real click. */
  pointer-events: none;
}

.doc-group-selected > .doc-group-proxy::after {
  content: "";
  position: absolute;
  inset: -3px;
  border: 2px solid #1a73e8;
  transform: rotate(var(--group-rotation, 0deg));
  pointer-events: none;
}
</style>
