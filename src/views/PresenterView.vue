<script setup lang="ts">
// Presenter view at /slides/:id/presenter/:n (SPEC-100): the tab that drives
// the show. It owns the playback machine and opens /slides/:id/present/:n
// ?follow=<session> as the audience window, then shows the presenter what the
// room cannot see — the current slide with its build state, the next slide,
// counters, the speaker notes, a timer and a grid of every slide. When no
// second window can open or follow, the current-slide pane becomes the show
// (single-window mode). A wiring hub under PresenterView.size.spec.ts: the
// logic lives in usePresenterView (which builds usePresenterSession,
// useAudienceWindow and usePresenterTools) and the UI in the Presenter*
// components. Runtime declarations only — espree lints .vue scripts.
import { useTemplateRef } from "vue";
import { useRoute, useRouter } from "vue-router";
import UiButton from "../components/ui/UiButton.vue";
import PresentationNotFound from "../components/presentation/PresentationNotFound.vue";
import PresenterSlidePane from "../components/presentation/PresenterSlidePane.vue";
import PresenterNotes from "../components/presentation/PresenterNotes.vue";
import PresenterGrid from "../components/presentation/PresenterGrid.vue";
import { usePresenterView } from "./presentation/usePresenterView";

const route = useRoute();
const router = useRouter();
const v = usePresenterView({ route, router });
const currentPane = useTemplateRef("currentPane");
</script>

<template>
  <PresentationNotFound
    v-if="v.loaded.value && !v.presentation.value"
    data-testid="presenter-not-found"
    @back="v.goHome()"
  />
  <div
    v-else
    class="presenter"
    data-testid="presenter-view"
  >
    <header class="presenter-bar">
      <h1 class="presenter-title">
        {{ v.presentation.value?.title ?? "Presenter view" }}
      </h1>
      <span
        v-if="v.blank.value !== 'none'"
        class="presenter-chip"
        data-testid="presenter-blank-badge"
      >
        {{ v.blank.value === "black" ? "Screen is black" : "Screen is white" }}
      </span>
      <p
        v-if="v.solo.value"
        class="presenter-solo-note"
        data-testid="presenter-solo-note"
      >
        Single-window mode: the current slide is the show.
        <UiButton
          v-if="!v.blockedBannerOpen.value"
          size="sm"
          variant="ghost"
          data-testid="presenter-audience-open"
          @click="v.reopenAudience()"
        >
          Open audience window
        </UiButton>
      </p>
      <UiButton
        data-testid="presenter-grid-toggle"
        :aria-pressed="v.gridOpen.value"
        @click="v.toggleGrid()"
      >
        All slides (G)
      </UiButton>
      <UiButton
        variant="primary"
        data-testid="presenter-exit"
        @click="v.exit()"
      >
        End show
      </UiButton>
    </header>

    <div
      v-if="v.solo.value && v.blockedBannerOpen.value"
      class="presenter-banner"
      role="status"
      data-testid="presenter-audience-blocked"
    >
      <span class="banner-text">
        {{ v.transport === "none"
          ? "This browser can't link a second window, so this one presents."
          : "The browser blocked the audience window, so this one presents." }}
      </span>
      <UiButton
        data-testid="presenter-audience-open"
        @click="v.reopenAudience()"
      >
        Open audience window
      </UiButton>
      <UiButton
        variant="primary"
        data-testid="presenter-present-here"
        @click="v.presentHere(currentPane?.$el)"
      >
        Present here
      </UiButton>
    </div>
    <div
      v-else-if="v.audienceStatus.value === 'lost'"
      class="presenter-banner"
      role="status"
      data-testid="presenter-audience-lost"
    >
      <span class="banner-text">The audience window closed. The show is still on this slide.</span>
      <UiButton
        variant="primary"
        data-testid="presenter-audience-open"
        @click="v.reopenAudience()"
      >
        Reopen audience window
      </UiButton>
    </div>

    <main
      v-if="v.presentation.value"
      class="presenter-columns"
    >
      <section
        class="presenter-main"
        aria-label="Current slide"
      >
        <PresenterSlidePane
          ref="currentPane"
          class="presenter-current"
          :class="{ 'is-show': v.solo.value }"
          data-testid="presenter-current-slide"
          :data-building="v.buildPreview.playing.value ? 'true' : 'false'"
          :slide="v.shownSlide.value"
          :width-cm="v.presentation.value.slideWidthCm"
          :height-cm="v.presentation.value.slideHeightCm"
          :media="v.presentation.value.media ?? []"
          :theme="v.presentation.value.theme"
          :fields="v.currentFields.value"
          :frame-styles="v.buildPreview.frameStyles.value"
          :paragraph-styles="v.buildPreview.paragraphStyles.value"
          :blank="v.solo.value ? v.blank.value : 'none'"
          empty-text="End of slide show"
          @click="v.solo.value && v.perform('advance')"
        />
        <div class="presenter-controls">
          <UiButton
            data-testid="presenter-prev-button"
            aria-label="Previous slide"
            @click="v.perform('back')"
          >
            ‹ Previous
          </UiButton>
          <span
            class="presenter-count"
            data-testid="presenter-slide-counter"
          >{{ v.counterLabel.value }}</span>
          <span
            v-if="v.buildTotal.value > 0"
            class="presenter-count"
            data-testid="presenter-build-indicator"
          >{{ v.buildLabel.value }}</span>
          <UiButton
            variant="primary"
            data-testid="presenter-next-button"
            aria-label="Next slide"
            @click="v.perform('advance')"
          >
            Next ›
          </UiButton>
        </div>
      </section>

      <aside class="presenter-side">
        <section
          class="presenter-next"
          aria-label="Next slide"
        >
          <h2 class="side-title">
            Next
          </h2>
          <PresenterSlidePane
            class="next-pane"
            data-testid="presenter-next-slide"
            :slide="v.nextSlide.value"
            :width-cm="v.presentation.value.slideWidthCm"
            :height-cm="v.presentation.value.slideHeightCm"
            :media="v.presentation.value.media ?? []"
            :theme="v.presentation.value.theme"
            :fields="v.nextFields.value"
            empty-text="End of show"
            empty-testid="presenter-next-end"
          />
        </section>
        <div class="presenter-timer-row">
          <span
            class="presenter-timer"
            role="timer"
            aria-label="Elapsed time"
            data-testid="presenter-timer"
          >{{ v.timerText.value }}</span>
          <UiButton
            size="sm"
            data-testid="presenter-timer-toggle"
            :aria-pressed="v.timerPaused.value"
            @click="v.toggleTimer()"
          >
            Pause (T)
          </UiButton>
          <UiButton
            size="sm"
            data-testid="presenter-timer-reset"
            @click="v.resetTimer()"
          >
            Reset (R)
          </UiButton>
        </div>
        <PresenterNotes
          class="presenter-notes-panel"
          :notes="v.notes.value"
          :scale="v.notesScale.value"
          :can-shrink="v.canShrinkNotes.value"
          :can-grow="v.canGrowNotes.value"
          @smaller="v.stepNotesScale(-1)"
          @larger="v.stepNotesScale(1)"
        />
      </aside>
    </main>

    <PresenterGrid
      v-if="v.gridOpen.value && v.presentation.value"
      :slides="v.presentation.value.slides"
      :width-cm="v.presentation.value.slideWidthCm"
      :height-cm="v.presentation.value.slideHeightCm"
      :media="v.presentation.value.media ?? []"
      :theme="v.presentation.value.theme"
      :current-index="v.playback.current.value.slideIndex"
      @choose="v.chooseSlide"
      @close="v.gridOpen.value = false"
    />

    <div
      class="sr-only"
      aria-live="polite"
      data-testid="presenter-live"
    >
      {{ v.liveMessage.value }}
    </div>
  </div>
</template>

<style scoped>
.presenter {
  height: 100vh;
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
  padding: var(--space-4);
  box-sizing: border-box;
  background: var(--surface-canvas);
  color: var(--text-primary);
}

.presenter-bar,
.presenter-banner,
.presenter-controls,
.presenter-timer-row {
  display: flex;
  align-items: center;
  gap: var(--space-3);
}

.presenter-title {
  flex: 1;
  min-width: 0;
  margin: 0;
  overflow: hidden;
  font-size: var(--text-lg);
  text-overflow: ellipsis;
  white-space: nowrap;
}

.presenter-chip {
  padding: var(--space-1) var(--space-3);
  border-radius: var(--radius-pill);
  background: var(--surface-selected);
  font-weight: var(--weight-bold);
}

.presenter-solo-note {
  display: flex;
  align-items: center;
  gap: var(--space-2);
  margin: 0;
  color: var(--text-secondary);
}

.presenter-banner {
  padding: var(--space-2) var(--space-3);
  border-radius: var(--radius-md);
  background: var(--surface-raised);
}

.banner-text {
  flex: 1;
}

.presenter-columns {
  flex: 1;
  min-height: 0;
  display: grid;
  grid-template-columns: minmax(0, 3fr) minmax(0, 2fr);
  gap: var(--space-4);
}

.presenter-main,
.presenter-side,
.presenter-next {
  display: flex;
  flex-direction: column;
  min-height: 0;
  gap: var(--space-3);
}

.presenter-current {
  flex: 1;
}

.presenter-current.is-show {
  cursor: pointer;
}

.presenter-controls {
  justify-content: center;
}

.presenter-count,
.presenter-timer {
  font-variant-numeric: tabular-nums;
}

.presenter-next {
  flex: 0 0 34%;
}

.next-pane {
  flex: 1;
}

.side-title {
  margin: 0;
  color: var(--text-secondary);
  font-size: var(--text-sm);
  font-weight: var(--weight-bold);
}

.presenter-timer {
  flex: 1;
  font-size: var(--text-xl);
  font-weight: var(--weight-bold);
}

.presenter-notes-panel {
  flex: 1;
}

.sr-only {
  position: absolute;
  width: 1px;
  height: 1px;
  padding: 0;
  margin: calc(-1 * var(--border-width));
  overflow: hidden;
  clip: rect(0, 0, 0, 0);
  white-space: nowrap;
  border: 0;
}
</style>
