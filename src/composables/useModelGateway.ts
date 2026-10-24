// The app layer's one door to the model gateway (SPEC-113). Importing this
// module is what loads lib/ai/modelGateway.ts, so it may only be imported by
// code that is itself behind an experiment's import() — never by anything
// main.ts or a route reaches statically (src/experimentGating.spec.ts walks
// the import graph to hold that). Loading it wires the gateway's consent
// prompt to the shell's ModelConsentDialog, once.
import * as gateway from "../lib/ai/modelGateway";
import { askModelConsent } from "./useModelConsent";

gateway.setConsentPrompt(askModelConsent);

/** The gateway, with its consent prompt wired. */
export function useModelGateway(): typeof gateway {
  return gateway;
}
