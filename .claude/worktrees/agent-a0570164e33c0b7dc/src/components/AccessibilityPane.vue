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
  background: var(--surface-raised);
  border-left: var(--border-width) solid var(--border-default);
  box-sizing: border-box;
  font-size: var(--text-md);
}

.accessibility-pane-header {
  flex: none;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--space-3);
  padding: var(--space-3) var(--space-4);
}

.accessibility-pane-title {
  font-size: var(--text-sm);
  font-weight: var(--weight-bold);
  letter-spacing: 0.04em;
  text-transform: uppercase;
  color: var(--text-secondary);
}

.accessibility-close {
  border: none;
  background: transparent;
  font-size: var(--text-lg);
  line-height: 1;
  cursor: pointer;
  color: var(--text-secondary);
}

.accessibility-summary {
  flex: none;
  margin: 0;
  padding: 0 var(--space-4) var(--space-3);
  color: var(--text-secondary);
}

.accessibility-clean {
  margin: 0;
  padding: var(--space-2) var(--space-4) var(--space-4);
  font-style: italic;
  color: var(--text-secondary);
}

.accessibility-groups {
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
}

.accessibility-group-title {
  margin: 0;
  padding: var(--space-2) var(--space-4);
  font-size: var(--text-sm);
  font-weight: var(--weight-bold);
  text-transform: uppercase;
  letter-spacing: 0.03em;
  color: var(--text-secondary);
}

.accessibility-issue-list {
  list-style: none;
  margin: 0;
  padding: 0 var(--space-3) var(--space-3);
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
}

.accessibility-issue {
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
  padding: var(--space-3);
  cursor: pointer;
}

.accessibility-issue-kind {
  font-weight: var(--weight-bold);
  color: var(--text-primary);
  margin-bottom: var(--space-1);
}

.accessibility-issue-explanation {
  margin: 0;
  color: var(--text-secondary);
}

.accessibility-issue-fix {
  margin-top: var(--space-3);
  border: var(--border-width) solid var(--border-default);
  background: var(--surface-base);
  border-radius: var(--radius-sm);
  padding: var(--space-1) var(--space-3);
  font-size: var(--text-sm);
  cursor: pointer;
}
</style>
