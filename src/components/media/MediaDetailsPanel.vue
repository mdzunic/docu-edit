<script setup lang="ts">
// The media details panel (SPEC-107): for an image — dimensions, format,
// byte size and stored date; for a PDF — page count, per-page size (or
// "mixed"), PDF version, /Info metadata, and the linearised/reconstructed
// flags. Everything shown is built from the stored model (parsePdf ran at
// open), so the panel is there even when the browser's embed is not.
// Runtime prop declarations — see AppShell.vue.
import { computed } from "vue";
import { formatBadgeLabel, formatFileSize, formatTimestamp } from "../../lib/format";
import { pdfPageSizesLabel } from "../../lib/media/thumbnail";
import { asMediaModel, currentImageSize, hasEdits } from "../../lib/media/editModel";
import { asMeasurementUnit } from "../../lib/core/units";
import { usePreferences } from "../../composables/usePreferences";

const props = defineProps({
  /** The MediaModel being viewed. */
  media: { type: Object, required: true },
  /** Epoch ms the library record was stored/last modified. */
  storedAt: { type: Number, required: true },
});

const pdf = computed(() => props.media.pdf ?? null);
// The current image's size (SPEC-112): an edited image reports what its
// stack produces — during an edit session, the live stack the preview shows
// — with the original's size beside it.
const image = computed(() => currentImageSize(asMediaModel(props.media)));
const original = computed(() => (hasEdits(asMediaModel(props.media)) ? props.media.image ?? null : null));
const properties = computed(() => props.media.properties ?? {});

const pageCountLabel = computed(() =>
  pdf.value ? `${pdf.value.pageCount} ${pdf.value.pageCount === 1 ? "page" : "pages"}` : "",
);
// The page size reads in points and, beside them, in the user's own unit.
const { preferences } = usePreferences();
const pageSizeLabel = computed(() =>
  pdf.value ? pdfPageSizesLabel(pdf.value, asMeasurementUnit(preferences.measurementUnit)) : "",
);
const producer = computed(() => {
  const extras = properties.value.extraFields ?? [];
  const entry = extras.find((field = { name: "", value: "" }) => field.name === "Producer");
  return entry ? entry.value : "";
});
const storedLabel = computed(() =>
  props.storedAt > 0 ? formatTimestamp(new Date(props.storedAt)) : "",
);
</script>

<template>
  <aside
    class="media-details"
    data-testid="media-details"
    aria-label="Details"
  >
    <h2 class="details-heading">
      Details
    </h2>
    <dl class="details-list">
      <div class="details-row">
        <dt>Format</dt>
        <dd data-testid="detail-format">
          {{ formatBadgeLabel(props.media.format) }} ({{ props.media.contentType }})
        </dd>
      </div>
      <div class="details-row">
        <dt>Size</dt>
        <dd data-testid="detail-size">
          {{ formatFileSize(props.media.byteLength) }}
        </dd>
      </div>
      <div
        v-if="image"
        class="details-row"
      >
        <dt>Dimensions</dt>
        <dd data-testid="detail-dimensions">
          {{ image.widthPx }} × {{ image.heightPx }} px
        </dd>
      </div>
      <div
        v-if="original"
        class="details-row"
      >
        <dt>Original</dt>
        <dd data-testid="detail-original-dimensions">
          {{ original.widthPx }} × {{ original.heightPx }} px · edited
        </dd>
      </div>
      <template v-if="pdf">
        <div class="details-row">
          <dt>Pages</dt>
          <dd data-testid="detail-pages">
            {{ pageCountLabel }}
          </dd>
        </div>
        <div class="details-row">
          <dt>Page size</dt>
          <dd data-testid="detail-page-size">
            {{ pageSizeLabel }}
          </dd>
        </div>
        <div class="details-row">
          <dt>PDF version</dt>
          <dd data-testid="detail-version">
            {{ pdf.version }}
          </dd>
        </div>
        <div
          v-if="properties.title"
          class="details-row"
        >
          <dt>Title</dt>
          <dd data-testid="detail-info-title">
            {{ properties.title }}
          </dd>
        </div>
        <div
          v-if="properties.author"
          class="details-row"
        >
          <dt>Author</dt>
          <dd data-testid="detail-info-author">
            {{ properties.author }}
          </dd>
        </div>
        <div
          v-if="producer"
          class="details-row"
        >
          <dt>Producer</dt>
          <dd data-testid="detail-producer">
            {{ producer }}
          </dd>
        </div>
        <div
          v-if="properties.created"
          class="details-row"
        >
          <dt>Created</dt>
          <dd data-testid="detail-created">
            {{ properties.created }}
          </dd>
        </div>
        <div
          v-if="properties.modified"
          class="details-row"
        >
          <dt>Modified</dt>
          <dd data-testid="detail-modified">
            {{ properties.modified }}
          </dd>
        </div>
        <div class="details-row">
          <dt>Linearised</dt>
          <dd data-testid="detail-linearised">
            {{ pdf.linearised ? "Yes" : "No" }}
          </dd>
        </div>
        <div class="details-row">
          <dt>Reconstructed</dt>
          <dd data-testid="detail-reconstructed">
            {{ pdf.reconstructed ? "Yes — the file's index was rebuilt to open it" : "No" }}
          </dd>
        </div>
      </template>
      <div
        v-if="storedLabel"
        class="details-row"
      >
        <dt>Stored</dt>
        <dd data-testid="detail-stored">
          {{ storedLabel }}
        </dd>
      </div>
    </dl>
  </aside>
</template>

<style scoped>
.media-details {
  flex: none;
  width: 18rem;
  padding: var(--space-5);
  border-left: var(--border-width) solid var(--border-subtle);
  background: var(--surface-base);
  overflow-y: auto;
}

.details-heading {
  margin: 0 0 var(--space-4);
  font-size: var(--text-lg);
  font-weight: var(--weight-bold);
  color: var(--text-primary);
}

.details-list {
  margin: 0;
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
}

.details-row dt {
  font-size: var(--text-sm);
  color: var(--text-secondary);
}

.details-row dd {
  margin: 0;
  color: var(--text-primary);
  overflow-wrap: anywhere;
}
</style>
