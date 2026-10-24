// The credential rules, held whichever way the provider config changes
// (SPEC-113). The gateway's own updateProviderConfig discards the API key
// and stale consent when the Labs form moves the origin; this guard applies
// the same two rules to every other path that can change the stored config —
// "Reset all preferences", an edit in another tab — so a key never outlives
// its origin and a grant never outlives its origin or model. main.ts installs
// it at startup. Network-free and gateway-free: installing it loads nothing
// that can send a request.
import { loadPreferences, onPreferencesChange } from "../preferences";
import { discardApiKeyUnlessFor } from "./apiKey";
import { pruneConsent } from "./consent";
import { originOf } from "./origin";
import type { ModelProviderConfig } from "./providerConfig";

/** Discards a key held for any origin but the configured one, and every
 * grant that no longer matches the configured origin and model. */
export function enforceProviderRules(config: ModelProviderConfig): void {
  const origin = originOf(config.baseUrl);
  discardApiKeyUnlessFor(origin);
  pruneConsent(origin, config.model.trim());
}

/** Applies the rules now and on every preferences change. Returns an
 * unsubscribe. */
export function installProviderGuard(): () => void {
  enforceProviderRules(loadPreferences().modelProvider);
  return onPreferencesChange((prefs) => enforceProviderRules(prefs.modelProvider));
}
