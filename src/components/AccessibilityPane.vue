<script setup lang="ts">
import { computed } from "vue";
import { accessibilityFixLabel, accessibilityKindLabel, asAccessibilityIssues } from "../views/document/useAccessibilityCheck";

// The accessibility checker's results pane (SPEC-043): issues grouped by
// severity (AC-2), each with a one-sentence explanation (AC-26) and a fix
// button where one exists (AC-27), a summary count (AC-29) and a clean
// state (AC-30). Docked the same way TrackChangesPane.vue is — a plain
// list with per-item actions, since AC-2 only asks for a severity-grouped
// list, not anchored/positioned cards like CommentsPane's.
// Runtime prop/emit declarations, not the generic form — espree parses
// .vue scripts and rejects TS-only syntax (see TrackChangesPane.vue).
const props = defineProps({
  issues: { type: Array, default: () => [] },
  scanning: { type: Boolean, default: false },
});

const emit = defineEmits(["close", "select", "fix"]);

const SEVERITY_GROUPS = [
  { severity: "error", label: "Error" },
  { severity: "warning", label: "Warning" },
  { severity: "tip", label: "Tip" },
];

const issuesTyped = computed(() => asAccessibilityIssues(props.issues));

const groups = computed(() =>
  SEVERITY_GROUPS.map((group) => ({
    ...group,
    items: issuesTyped.value.filter((issue) => issue.severity === group.severity),
  })),
);

const summary = computed(() => {
  const count = issuesTyped.value.length;
  if (count === 0) {
    return props.scanning ? "Checking…" : "No accessibility issues found";
  }
  return `${count} issue${count === 1 ? "" : "s"} found`;
});

</script>

<template>
  <aside
    class="accessibility-pane"
    aria-label="Accessibility checker"
    data-testid="accessibility-pane"
  >
    <div class="accessibility-pane-header">
      <span class="accessibility-pane-title">Accessibility</span>
      <button
        type="button"
        class="accessibility-close"
        data-testid="accessibility-pane-close"
        aria-label="Close accessibility checker"
        @click="emit('close')"
      >
        ×
      </button>
    </div>

    <p
      class="accessibility-summary"
      data-testid="accessibility-summary"
    >
      {{ summary }}
    </p>

    <p
      v-if="!scanning && issuesTyped.length === 0"
      class="accessibility-clean"
      data-testid="accessibility-clean"
    >
      No accessibility issues found
    </p>

    <div
      v-else
      class="accessibility-groups"
    >
      <section
        v-for="group in groups"
        v-show="group.items.length > 0"
        :key="group.severity"
        class="accessibility-group"
        :data-testid="`accessibility-group-${group.severity}`"
      >
        <h3 class="accessibility-group-title">
          {{ group.label }} ({{ group.items.length }})
        </h3>
        <ul class="accessibility-issue-list">
          <li
            v-for="issue in group.items"
            :key="issue.id"
            class="accessibility-issue"
            :data-testid="`accessibility-issue-${issue.id}`"
            tabindex="0"
            role="button"
            @click="emit('select', issue.id)"
            @keydown.enter="emit('select', issue.id)"
          >
            <div class="accessibility-issue-kind">
              {{ accessibilityKindLabel(issue.kind) }}
            </div>
            <p class="accessibility-issue-explanation">
              {{ issue.explanation }}
            </p>
            <button
              v-if="issue.fix"
              type="button"
              class="accessibility-issue-fix"
              data-testid="accessibility-issue-fix"
              @click.stop="emit('fix', issue.id)"
            >
              {{ accessibilityFixLabel(issue.fix) }}
            </button>
          </li>
        </ul>
      </section>
    </div>
  </aside>
</template>

<style scoped>
.accessibility-pane {
  flex: none;
  align-self: flex-start;
  position: sticky;
  top: 0;
  width: 20rem;
  max-height: 100vh;
  display: flex;
  flex-direction: column;
  min-height: 0;
  overflow-y: auto;
  background: #f8f9fa;
  border-left: 1px solid #dadce0;
  box-sizing: border-box;
  font-size: 0.8125rem;
}

.accessibility-pane-header {
  flex: none;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 0.5rem;
  padding: 0.5rem 0.75rem;
}

.accessibility-pane-title {
  font-size: 0.75rem;
  font-weight: 600;
  letter-spacing: 0.04em;
  text-transform: uppercase;
  color: #5f6368;
}

.accessibility-close {
  border: none;
  background: transparent;
  font-size: 1rem;
  line-height: 1;
  cursor: pointer;
  color: #5f6368;
}

.accessibility-summary {
  flex: none;
  margin: 0;
  padding: 0 0.75rem 0.5rem;
  color: #5f6368;
}

.accessibility-clean {
  margin: 0;
  padding: 0.25rem 0.75rem 0.75rem;
  font-style: italic;
  color: #5f6368;
}

.accessibility-groups {
  display: flex;
  flex-direction: column;
  gap: 0.25rem;
}

.accessibility-group-title {
  margin: 0;
  padding: 0.25rem 0.75rem;
  font-size: 0.72rem;
  font-weight: 600;
  text-transform: uppercase;
  letter-spacing: 0.03em;
  color: #5f6368;
}

.accessibility-issue-list {
  list-style: none;
  margin: 0;
  padding: 0 0.5rem 0.5rem;
  display: flex;
  flex-direction: column;
  gap: 0.4rem;
}

.accessibility-issue {
  background: #fff;
  border: 1px solid #dadce0;
  border-radius: 6px;
  padding: 0.5rem;
  cursor: pointer;
}

.accessibility-issue-kind {
  font-weight: 600;
  color: #202124;
  margin-bottom: 0.15rem;
}

.accessibility-issue-explanation {
  margin: 0;
  color: #5f6368;
}

.accessibility-issue-fix {
  margin-top: 0.4rem;
  border: 1px solid #dadce0;
  background: #fff;
  border-radius: 4px;
  padding: 0.15rem 0.5rem;
  font-size: 0.72rem;
  cursor: pointer;
}
</style>
