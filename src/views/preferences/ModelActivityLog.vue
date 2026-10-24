<script setup lang="ts">
import { onMounted, onUnmounted, shallowRef } from "vue";
import UiButton from "../../components/ui/UiButton.vue";
import { useModelGateway } from "../../composables/useModelGateway";

// SPEC-113: the gateway's session-only activity log, viewable and clearable
// from Labs. One row per request that actually went out: when, where, which
// model, what for, the bytes each way and the outcome — never a word of the
// prompt or the response, which the gateway does not keep. In memory only;
// a reload empties it.
const gateway = useModelGateway();
const entries = shallowRef(gateway.activityLog());

let stop = () => {};
onMounted(() => {
  entries.value = gateway.activityLog();
  stop = gateway.onActivityLogChange(() => {
    entries.value = gateway.activityLog();
  });
});
onUnmounted(() => stop());

function clockTime(time = 0) {
  return new Date(time).toLocaleTimeString();
}
</script>

<template>
  <div
    class="activity"
    data-testid="ai-activity"
  >
    <div class="activity-head">
      <h5 class="subheading">
        Activity this session
      </h5>
      <UiButton
        size="sm"
        data-testid="ai-activity-clear"
        :disabled="entries.length === 0"
        @click="gateway.clearActivityLog()"
      >
        Clear
      </UiButton>
    </div>
    <p class="muted">
      Kept in memory only and cleared on reload. Byte counts, never the text sent or received.
    </p>
    <p
      v-if="entries.length === 0"
      class="muted"
      data-testid="ai-activity-empty"
    >
      No requests yet.
    </p>
    <div
      v-else
      class="activity-scroll"
    >
      <table class="activity-table">
        <thead>
          <tr>
            <th scope="col">
              Time
            </th>
            <th scope="col">
              Destination
            </th>
            <th scope="col">
              Model
            </th>
            <th scope="col">
              For
            </th>
            <th scope="col">
              Sent
            </th>
            <th scope="col">
              Received
            </th>
            <th scope="col">
              Outcome
            </th>
          </tr>
        </thead>
        <tbody>
          <tr
            v-for="(entry, index) in entries"
            :key="index"
            data-testid="ai-activity-entry"
            :data-request-bytes="entry.requestBytes"
            :data-response-bytes="entry.responseBytes"
            :data-outcome="entry.outcome"
          >
            <td>{{ clockTime(entry.time) }}</td>
            <td>{{ entry.origin }}</td>
            <td>{{ entry.model || "—" }}</td>
            <td>{{ entry.purpose }}</td>
            <td>{{ entry.requestBytes }} B</td>
            <td>{{ entry.responseBytes }} B</td>
            <td>{{ entry.outcome }}</td>
          </tr>
        </tbody>
      </table>
    </div>
  </div>
</template>

<style scoped>
.activity {
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
}

.activity-head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--space-3);
}

.subheading {
  margin: 0;
  color: var(--text-primary);
  font-size: var(--text-sm);
  font-weight: var(--weight-bold);
}

.muted {
  margin: 0;
  color: var(--text-secondary);
  font-size: var(--text-sm);
  line-height: var(--line-normal);
}

.activity-scroll {
  overflow-x: auto;
}

.activity-table {
  width: 100%;
  border-collapse: collapse;
  font-size: var(--text-sm);
}

.activity-table th,
.activity-table td {
  padding: var(--space-1) var(--space-2);
  border-bottom: var(--border-width) solid var(--border-subtle);
  text-align: start;
  white-space: nowrap;
}

.activity-table th {
  color: var(--text-secondary);
  font-weight: var(--weight-bold);
}
</style>
