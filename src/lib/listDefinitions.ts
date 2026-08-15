// The document-model half of multilevel list definitions (SPEC-032): the
// defId → definition lookup. Everything else — gallery presets, marker
// composition and the instance-counter scan — moved to
// core/listDefinitions.ts (product-neutral) and is re-exported here so
// existing import paths keep working.
import type { DocumentModel } from "./documentModel";
import type { ListDefinition } from "./core/lists";

/** The definition a defId resolves to, or undefined. */
export function definitionById(
  model: DocumentModel,
  defId: string | undefined,
): ListDefinition | undefined {
  if (defId === undefined) {
    return undefined;
  }
  return model.listDefinitions?.find((def) => def.defId === defId);
}
