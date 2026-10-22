<script setup lang="ts">
// SPEC-056 equation renderer — DOM, not canvas. Math layout is nested boxes
// and horizontal rules, which is what CSS is for, and DOM gives selection,
// text search, screen-reader access, PDF text extraction and e2e assertions
// for free. Fraction and radical bars are CSS borders. This component is
// recursive: a box's children are MathBox instances. The spacing, variant and
// stretch decisions are precomputed by src/lib/math/render.ts's RenderBox.
// Runtime prop declarations, not defineProps<{...}>() generics — espree cannot
// parse TypeScript in a .vue script (see DocTableBlock.vue's note).

defineProps({
  /** A src/lib/math/render.ts RenderBox. */
  box: { type: Object, required: true },
});

// Bracket-piece assembly threshold (AC-44): a delimiter taller than 3× is
// assembled from Unicode pieces rather than a CSS transform.
const MAX_TRANSFORM_STRETCH = 3;

// Unicode bracket pieces (U+239B–U+23AD): [top, middle-extension, bottom] for
// a delimiter that must stretch beyond a CSS transform (AC-44). A brace also
// has a centre joint piece, appended as a 4th entry.
const BRACKET_PIECES = new Map([
  ["(", ["⎛", "⎜", "⎝"]],
  [")", ["⎞", "⎟", "⎠"]],
  ["[", ["⎡", "⎢", "⎣"]],
  ["]", ["⎤", "⎥", "⎦"]],
  ["{", ["⎧", "⎪", "⎩", "⎨"]],
  ["}", ["⎫", "⎪", "⎭", "⎬"]],
]);

function scaleTransform(rows = 1) {
  if (rows <= 1) return undefined;
  return `scaleY(${Math.min(rows, MAX_TRANSFORM_STRETCH)})`;
}

/** Whether a delimiter must be assembled from Unicode pieces (taller than 3×
 * and one this editor has pieces for). */
function assembled(char = "", rows = 1) {
  return rows > MAX_TRANSFORM_STRETCH && BRACKET_PIECES.has(char);
}

/** The stacked Unicode pieces for a tall delimiter: top hook, a run of
 * extension pieces sized to the height, an optional centre joint (braces),
 * then the bottom hook (AC-44). */
function bracketStack(char = "", rows = 1) {
  const pieces = BRACKET_PIECES.get(char);
  if (!pieces) return [];
  const [top, ext, bottom, centre] = pieces;
  const extraRows = Math.max(1, rows - 2);
  if (centre) {
    // A brace: split the extension around the centre joint.
    const half = Math.max(1, Math.floor(extraRows / 2));
    return [top, ...Array(half).fill(ext), centre, ...Array(half).fill(ext), bottom];
  }
  return [top, ...Array(extraRows).fill(ext), bottom];
}
</script>

<template>
  <span
    v-if="box.type === 'row'"
    class="math-row"
  ><span
    v-for="(item, i) in box.items"
    :key="i"
    class="math-item"
    :style="item.spaceBefore === '0' ? undefined : { marginLeft: item.spaceBefore }"
  ><MathBox :box="item.box" /></span></span>

  <span
    v-else-if="box.type === 'glyph'"
    class="math-glyph"
    :style="{ fontStyle: box.variant === 'italic' ? 'italic' : 'normal' }"
  >{{ box.text }}</span>

  <span
    v-else-if="box.type === 'text'"
    class="math-text"
  >{{ box.text }}</span>

  <span
    v-else-if="box.type === 'frac'"
    class="math-frac"
    :style="box.fontScale ? { fontSize: box.fontScale } : undefined"
  ><span class="math-frac-num"><MathBox :box="box.num" /></span><span
    class="math-frac-bar"
    :class="{ 'math-frac-nobar': !box.bar }"
  /><span class="math-frac-den"><MathBox :box="box.den" /></span></span>

  <span
    v-else-if="box.type === 'script'"
    class="math-script"
  ><span class="math-script-base"><MathBox :box="box.base" /></span><span
    class="math-script-stack"
    :style="box.stackScale ? { fontSize: box.stackScale } : undefined"
  ><span
    v-if="box.sup"
    class="math-sup"
  ><MathBox :box="box.sup" /></span><span
    v-if="box.sub"
    class="math-sub"
  ><MathBox :box="box.sub" /></span></span></span>

  <span
    v-else-if="box.type === 'radical'"
    class="math-radical"
  ><span
    v-if="box.index"
    class="math-radical-index"
  ><MathBox :box="box.index" /></span><span
    class="math-radical-sign"
    aria-hidden="true"
  >√</span><span class="math-radical-body"><MathBox :box="box.radicand" /></span></span>

  <span
    v-else-if="box.type === 'nary'"
    class="math-nary"
  ><span class="math-nary-limits"><span
    v-if="box.upper"
    class="math-nary-upper"
  ><MathBox :box="box.upper" /></span><span
    class="math-nary-op"
  >{{ box.op }}</span><span
    v-if="box.lower"
    class="math-nary-lower"
  ><MathBox :box="box.lower" /></span></span><span class="math-nary-body"><MathBox :box="box.body" /></span></span>

  <span
    v-else-if="box.type === 'delim'"
    class="math-delim"
  ><span
    v-if="box.open && assembled(box.open, box.stretch)"
    class="math-delim-pieces"
    aria-hidden="true"
  ><span
    v-for="(piece, i) in bracketStack(box.open, box.stretch)"
    :key="i"
    class="math-delim-piece"
  >{{ piece }}</span></span><span
    v-else-if="box.open"
    class="math-delim-open"
    :style="{ transform: scaleTransform(box.stretch) }"
    aria-hidden="true"
  >{{ box.open }}</span><span class="math-delim-body"><MathBox :box="box.body" /></span><span
    v-if="box.close && assembled(box.close, box.stretch)"
    class="math-delim-pieces"
    aria-hidden="true"
  ><span
    v-for="(piece, i) in bracketStack(box.close, box.stretch)"
    :key="i"
    class="math-delim-piece"
  >{{ piece }}</span></span><span
    v-else-if="box.close"
    class="math-delim-close"
    :style="{ transform: scaleTransform(box.stretch) }"
    aria-hidden="true"
  >{{ box.close }}</span></span>

  <span
    v-else-if="box.type === 'matrix'"
    class="math-matrix"
  ><span
    v-for="(mrow, r) in box.rows"
    :key="r"
    class="math-matrix-row"
  ><span
    v-for="(cell, c) in mrow"
    :key="c"
    class="math-matrix-cell"
  ><MathBox :box="cell" /></span></span></span>

  <span
    v-else-if="box.type === 'accent'"
    class="math-accent"
  ><span
    class="math-accent-mark"
    aria-hidden="true"
  >{{ box.accent }}</span><span class="math-accent-base"><MathBox :box="box.base" /></span></span>

  <span
    v-else-if="box.type === 'func'"
    class="math-func"
  ><span class="math-func-name">{{ box.name }}</span><span
    v-if="box.body.type === 'row' && box.body.items.length"
    class="math-func-body"
  ><MathBox :box="box.body" /></span></span>

  <span
    v-else-if="box.type === 'opaque'"
    class="math-opaque"
  >{{ box.linear }}</span>
</template>

<style scoped>
.math-row {
  display: inline-flex;
  align-items: center;
  white-space: nowrap;
}
.math-item {
  display: inline-flex;
  align-items: center;
}
.math-glyph,
.math-text,
.math-func-name {
  font-family: "Cambria Math", "Latin Modern Math", Caladea, "Liberation Serif", serif;
}
.math-text,
.math-func-name {
  font-style: normal;
}
/* Fraction: a vertical stack with a rule between; the bar's centre sits on the
   surrounding text's maths axis (AC-43) via vertical-align: middle. */
.math-frac {
  display: inline-flex;
  flex-direction: column;
  align-items: center;
  vertical-align: middle;
  text-align: center;
}
.math-frac-num,
.math-frac-den {
  display: block;
  padding: 0 0.2em;
}
/* 1/18 em rule with 1/6 em clearance above/below (AC-42). */
.math-frac-bar {
  display: block;
  align-self: stretch;
  border-top: 0.0556em solid currentColor;
  margin: 0.1667em 0;
}
.math-frac-nobar {
  border-top-color: transparent;
}
/* Scripts are 71% of their base and a script-of-a-script floors at 50% and
   does not shrink further (AC-39); the per-level factor and the 5 pt floor
   (AC-41) are computed by render.ts and applied inline as the stack's
   font-size, so nesting does not compound past the floor. */
.math-script {
  display: inline-flex;
  align-items: center;
}
.math-script-stack {
  display: inline-flex;
  flex-direction: column;
}
.math-sup {
  align-self: flex-start;
  margin-bottom: 0.4em;
}
.math-sub {
  align-self: flex-start;
  margin-top: 0.4em;
}
.math-radical {
  display: inline-flex;
  align-items: flex-start;
}
.math-radical-index {
  font-size: 0.5em;
  align-self: flex-start;
}
.math-radical-sign {
  font-family: serif;
}
/* The radical bar over the radicand: same 1/18 em rule (AC-42). */
.math-radical-body {
  border-top: 0.0556em solid currentColor;
  padding: 0.1667em 0.15em 0;
}
.math-nary {
  display: inline-flex;
  align-items: center;
}
.math-nary-limits {
  display: inline-flex;
  flex-direction: column;
  align-items: center;
}
.math-nary-op {
  font-size: 1.6em;
  line-height: 1;
}
.math-nary-upper,
.math-nary-lower {
  font-size: 0.6em;
}
.math-delim {
  display: inline-flex;
  align-items: center;
}
.math-delim-open,
.math-delim-close {
  font-family: serif;
  transform-origin: center;
}
/* A delimiter too tall for a CSS transform, assembled from Unicode bracket
   pieces stacked vertically (AC-44). */
.math-delim-pieces {
  display: inline-flex;
  flex-direction: column;
  align-items: center;
  font-family: serif;
  line-height: 0.62;
}
.math-delim-piece {
  display: block;
}
.math-matrix {
  display: inline-flex;
  flex-direction: column;
  vertical-align: middle;
}
.math-matrix-row {
  display: inline-flex;
  justify-content: space-around;
}
.math-matrix-cell {
  padding: 0.1em 0.4em;
}
.math-accent {
  display: inline-flex;
  flex-direction: column;
  align-items: center;
}
.math-accent-mark {
  line-height: 0.4;
}
.math-func {
  display: inline-flex;
  align-items: center;
}
.math-func-body {
  margin-left: 0.1667em;
}
.math-nary-body {
  margin-left: 0.1667em;
}
.math-opaque {
  font-family: serif;
}
</style>
