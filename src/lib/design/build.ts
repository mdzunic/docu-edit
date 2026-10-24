// Building a reviewed plan into a deck (SPEC-115) — the whole Build step, and
// no model request in it. The plan becomes a DeckDescription
// (planToDescription), `deckFromDescription` validates it and produces the
// PresentationModel through `deckContextFromDescription` — the very path
// `PresentationBuilder.fromDescription` takes — then `applyPlanDesign` sets
// the theme and kinds, and the provenance custom properties are added. Any
// validation failure is SPEC-106's `DocumentGenerationError`, thrown before
// anything exists to store.
import type { CustomProperty, DocumentProperties } from "../core/documentProperties";
import { deckFromDescription } from "../generate/deckDescription";
import type { PresentationModel } from "../presentation/model";
import { applyPlanDesign } from "./applyDesign";
import { planToDescription, type DeckPlan } from "./plan";

/** This application, as the GeneratedBy property names it. */
export const GENERATED_BY = "docu-edit";

/** The custom properties a designed deck carries (SPEC-041's
 * `customProperties`, written into both formats). */
export const PROVENANCE_PROPERTY_NAMES = ["GeneratedBy", "GeneratedAt", "GeneratedModel", "GeneratedPrompt"] as const;

export interface DeckProvenance {
  /** ISO 8601 time of the build. */
  generatedAt: string;
  /** The model id that answered, or "outline-designer". */
  generatedModel: string;
  /** The brief — only when the user ticked "keep the prompt". */
  prompt?: string;
}

/** `properties` with the provenance custom properties set (any earlier
 * Generated* entries replaced). */
export function withProvenance(properties: DocumentProperties | undefined, provenance: DeckProvenance): DocumentProperties {
  const names: readonly string[] = PROVENANCE_PROPERTY_NAMES;
  const kept = (properties?.customProperties ?? []).filter((property) => !names.includes(property.name));
  const added: CustomProperty[] = [
    { name: "GeneratedBy", type: "text", value: GENERATED_BY },
    { name: "GeneratedAt", type: "text", value: provenance.generatedAt },
    { name: "GeneratedModel", type: "text", value: provenance.generatedModel },
  ];
  if (provenance.prompt !== undefined) {
    added.push({ name: "GeneratedPrompt", type: "text", value: provenance.prompt });
  }
  return { ...properties, customProperties: [...kept, ...added] };
}

/** Builds `plan` (already validated — see validatePlan) into a new deck. */
export function buildPlannedDeck(plan: DeckPlan, provenance: DeckProvenance): PresentationModel {
  const model = deckFromDescription(planToDescription(plan));
  applyPlanDesign(model, plan);
  model.properties = withProvenance(model.properties, provenance);
  return model;
}
