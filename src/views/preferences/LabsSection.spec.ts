// SPEC-113 — the Labs section and the model provider panel it gates: the
// experiment list with its standing note and toggles; the panel absent and
// its code (and the gateway's) never loaded while no model experiment is on;
// and, once it is, the mode-dependent fields, the masked key, the key cleared
// with a line explaining why when the origin moves, "Test connection"
// reporting both outcomes, revocable consent and the activity log.
import { flushPromises, mount, type VueWrapper } from "@vue/test-utils";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import LabsSection from "./LabsSection.vue";
import { EXPERIMENTS, isExperimentEnabled, setExperimentEnabled } from "../../lib/experiments";
import { loadPreferences, resetPreferences, savePreference } from "../../lib/preferences";
import { PROVIDER_CONFIG_DEFAULTS } from "../../lib/ai/providerConfig";
import { API_KEY_STORAGE_KEY } from "../../lib/ai/apiKey";

// Counts loads of the app's one door to the gateway: the factory runs only
// when something actually imports the module.
const loads = vi.hoisted(() => ({ gateway: 0 }));
vi.mock("../../composables/useModelGateway", async (importOriginal) => {
  loads.gateway += 1;
  return importOriginal();
});

let calls: { url: string; init: RequestInit }[] = [];
let respond: () => Promise<Response> | Response = () => new Response("{}");

beforeEach(() => {
  localStorage.clear();
  resetPreferences();
  calls = [];
  vi.stubGlobal(
    "fetch",
    vi.fn(async (input: RequestInfo | URL, init: RequestInit = {}) => {
      calls.push({ url: String(input), init });
      return respond();
    }),
  );
});

afterEach(() => {
  vi.unstubAllGlobals();
});

const find = (wrapper: VueWrapper, testid: string) => wrapper.find(`[data-testid="${testid}"]`);

async function mountWithPanel(): Promise<VueWrapper> {
  setExperimentEnabled("deck-designer", true);
  const wrapper = mount(LabsSection, { attachTo: document.body });
  await vi.waitFor(() => expect(find(wrapper, "ai-provider").exists()).toBe(true));
  return wrapper;
}

/** The gateway the panel loaded, with a consent prompt the test controls. */
async function gatewayAnswering(choice: "once" | "always" | "cancel") {
  const gateway = await import("../../lib/ai/modelGateway");
  gateway.setConsentPrompt(async () => choice);
  return gateway;
}

async function change(wrapper: VueWrapper, testid: string, value: string): Promise<void> {
  await find(wrapper, testid).setValue(value);
  await find(wrapper, testid).trigger("change");
  await flushPromises();
}

describe("the experiment list", () => {
  it("lists every experiment with its title, summary, risk line and an off toggle", () => {
    const wrapper = mount(LabsSection);
    expect(find(wrapper, "prefs-labs").exists()).toBe(true);
    for (const experiment of EXPERIMENTS) {
      const item = find(wrapper, `prefs-labs-item-${experiment.id}`);
      expect(item.text()).toContain(experiment.title);
      expect(item.text()).toContain(experiment.summary);
      expect(find(wrapper, `experiment-risk-${experiment.id}`).text()).toBe(`Risk: ${experiment.risk}`);
      const toggle = find(wrapper, `experiment-${experiment.id}`);
      expect(toggle.attributes("type")).toBe("checkbox");
      expect((toggle.element as HTMLInputElement).checked).toBe(false);
    }
  });

  it("carries the standing note about change, removal and round-trip guarantees", () => {
    const note = find(mount(LabsSection), "prefs-labs-note").text();
    expect(note).toContain("Experimental features may change or be removed");
    expect(note).toContain("not covered by this product's round-trip guarantees");
  });

  it("a toggle switches its experiment live", async () => {
    const wrapper = mount(LabsSection);
    await find(wrapper, "experiment-library-search").setValue(true);
    expect(isExperimentEnabled("library-search")).toBe(true);
    expect(isExperimentEnabled("deck-designer")).toBe(false);
    await find(wrapper, "experiment-library-search").setValue(false);
    expect(loadPreferences().experiments["library-search"]).toBe(false);
  });
});

describe("the model provider panel is gated", () => {
  it("renders nothing and loads nothing — not even the gateway — while no model experiment is on", async () => {
    const wrapper = mount(LabsSection);
    setExperimentEnabled("library-search", true); // not a model experiment
    await flushPromises();
    expect(find(wrapper, "ai-provider").exists()).toBe(false);
    expect(find(wrapper, "experiment-badge").exists()).toBe(false);
    expect(loads.gateway).toBe(0);
  });

  it("appears, with the experiment badge, once deck-designer is on — and goes when it is off", async () => {
    const wrapper = await mountWithPanel();
    expect(loads.gateway).toBe(1);
    expect(find(wrapper, "ai-provider").find('[data-testid="experiment-badge"]').text()).toBe("Experimental");
    await find(wrapper, "experiment-deck-designer").setValue(false);
    await flushPromises();
    expect(find(wrapper, "ai-provider").exists()).toBe(false);
    wrapper.unmount();
  });
});

describe("the provider form", () => {
  it("is off by default, preconfigures nothing, and shows only the mode until a provider is chosen", async () => {
    const wrapper = await mountWithPanel();
    expect((find(wrapper, "ai-mode").element as HTMLSelectElement).value).toBe("off");
    expect(loadPreferences().modelProvider).toEqual(PROVIDER_CONFIG_DEFAULTS);
    expect(find(wrapper, "ai-base-url").exists()).toBe(false);
    expect(find(wrapper, "ai-test").exists()).toBe(false);
    wrapper.unmount();
  });

  it("local mode shows endpoint, model and bounds but no key field; cloud mode adds the masked key", async () => {
    const wrapper = await mountWithPanel();
    await find(wrapper, "ai-mode").setValue("local");
    await flushPromises();
    for (const testid of ["ai-base-url", "ai-model", "ai-timeout", "ai-max-tokens", "ai-temperature", "ai-test"]) {
      expect(find(wrapper, testid).exists(), testid).toBe(true);
    }
    expect(find(wrapper, "ai-api-key").exists()).toBe(false);
    await find(wrapper, "ai-mode").setValue("cloud");
    await flushPromises();
    expect(find(wrapper, "ai-api-key").attributes("type")).toBe("password");
    wrapper.unmount();
  });

  it("writes edits through to the stored config and says where requests would go", async () => {
    const wrapper = await mountWithPanel();
    await find(wrapper, "ai-mode").setValue("local");
    await flushPromises();
    await change(wrapper, "ai-base-url", "http://localhost:11434/v1");
    await change(wrapper, "ai-model", "llama3.1:8b");
    await change(wrapper, "ai-timeout", "30");
    expect(loadPreferences().modelProvider).toMatchObject({
      mode: "local",
      baseUrl: "http://localhost:11434/v1",
      model: "llama3.1:8b",
      timeoutMs: 30_000,
    });
    expect(find(wrapper, "ai-origin").text()).toBe("Requests go to http://localhost:11434, on this computer.");
    await change(wrapper, "ai-timeout", "0");
    expect(loadPreferences().modelProvider.timeoutMs).toBe(30_000);
    await change(wrapper, "ai-base-url", "http://models.example.test/v1");
    expect(wrapper.text()).toContain("plain http: is only allowed for this computer");
    wrapper.unmount();
  });

  it("stores the key outside preferences, and clears the field with a line saying why when the origin moves", async () => {
    const wrapper = await mountWithPanel();
    await find(wrapper, "ai-mode").setValue("cloud");
    await flushPromises();
    await change(wrapper, "ai-base-url", "https://a.example.test/v1");
    await change(wrapper, "ai-api-key", "sk-test-key");
    expect(localStorage.getItem(API_KEY_STORAGE_KEY)).toContain("sk-test-key");
    expect(JSON.stringify(loadPreferences())).not.toContain("sk-test-key");
    expect(find(wrapper, "ai-api-key-cleared").exists()).toBe(false);

    await change(wrapper, "ai-base-url", "https://a.example.test/v2"); // same origin: kept
    expect((find(wrapper, "ai-api-key").element as HTMLInputElement).value).toBe("sk-test-key");

    await change(wrapper, "ai-base-url", "https://b.example.test/v1");
    expect((find(wrapper, "ai-api-key").element as HTMLInputElement).value).toBe("");
    expect(localStorage.getItem(API_KEY_STORAGE_KEY)).toBeNull();
    const notice = find(wrapper, "ai-api-key-cleared").text();
    expect(notice).toContain("The API key was removed");
    expect(notice).toContain("https://a.example.test");
    expect(notice).toContain("https://b.example.test");
    wrapper.unmount();
  });
});

describe("standing consent and external changes", () => {
  it("stays listed and revocable after its experiment is switched off", async () => {
    const { grantConsent, consentGrants } = await import("../../lib/ai/consent");
    grantConsent("http://localhost:11434", "m", "Test connection");
    const wrapper = mount(LabsSection);
    expect(find(wrapper, "ai-provider").exists()).toBe(false);
    expect(wrapper.findAll('[data-testid="ai-consent-grant"]')).toHaveLength(1);
    await find(wrapper, "ai-consent-revoke").trigger("click");
    await flushPromises();
    expect(consentGrants()).toEqual([]);
    expect(find(wrapper, "ai-consents").exists()).toBe(false);
  });

  it("says why the key went when another tab moved the endpoint's origin", async () => {
    const wrapper = await mountWithPanel();
    await find(wrapper, "ai-mode").setValue("cloud");
    await flushPromises();
    await change(wrapper, "ai-base-url", "https://a.example.test/v1");
    await change(wrapper, "ai-api-key", "sk-test-key");
    // Another tab's edit, then the guard main.ts installs discarding the key.
    const { installProviderGuard } = await import("../../lib/ai/providerGuard");
    const stop = installProviderGuard();
    savePreference("modelProvider", { ...loadPreferences().modelProvider, baseUrl: "https://c.example.test/v1" });
    await flushPromises();
    stop();
    expect((find(wrapper, "ai-api-key").element as HTMLInputElement).value).toBe("");
    expect(find(wrapper, "ai-api-key-cleared").text()).toContain(
      "moved from https://a.example.test to https://c.example.test",
    );
    wrapper.unmount();
  });
});

describe("Test connection", () => {
  async function configured(): Promise<VueWrapper> {
    const wrapper = await mountWithPanel();
    savePreference("modelProvider", {
      ...PROVIDER_CONFIG_DEFAULTS,
      mode: "local",
      baseUrl: "http://localhost:11434/v1",
      model: "test-model",
    });
    await flushPromises();
    return wrapper;
  }

  it("reports reachable, the model id and the round trip in milliseconds", async () => {
    const wrapper = await configured();
    await gatewayAnswering("once");
    respond = () => new Response(JSON.stringify({ model: "served-model", choices: [{ message: { content: "OK" } }] }));
    await find(wrapper, "ai-test").trigger("click");
    await vi.waitFor(() => expect(find(wrapper, "ai-test-result").attributes("data-state")).toBe("ok"));
    expect(find(wrapper, "ai-test-result").text()).toMatch(/^Reachable — model served-model, \d+ ms round trip\.$/);
    expect(calls).toHaveLength(1);
    wrapper.unmount();
  });

  it("reports the typed error when it fails — the loopback cause spelled out", async () => {
    const wrapper = await configured();
    await gatewayAnswering("once");
    respond = () => Promise.reject(new TypeError("Failed to fetch"));
    await find(wrapper, "ai-test").trigger("click");
    await vi.waitFor(() => expect(find(wrapper, "ai-test-result").attributes("data-state")).toBe("error"));
    expect(find(wrapper, "ai-test-result").text()).toContain(
      "Nothing is listening on http://localhost:11434; is your local model running?",
    );
    wrapper.unmount();
  });

  it("is subject to consent: declined, nothing is sent and it says so", async () => {
    const wrapper = await configured();
    await gatewayAnswering("cancel");
    await find(wrapper, "ai-test").trigger("click");
    await vi.waitFor(() => expect(find(wrapper, "ai-test-result").attributes("data-state")).toBe("error"));
    expect(find(wrapper, "ai-test-result").text()).toContain("nothing was sent");
    expect(calls).toHaveLength(0);
    wrapper.unmount();
  });

  it("logs the request in the activity log (counts only), which Labs can clear; Always is revocable", async () => {
    const wrapper = await configured();
    const gateway = await gatewayAnswering("always");
    gateway.clearActivityLog();
    respond = () => new Response(JSON.stringify({ model: "served-model", choices: [{ message: { content: "OK" } }] }));
    await find(wrapper, "ai-test").trigger("click");
    await vi.waitFor(() => expect(wrapper.findAll('[data-testid="ai-activity-entry"]')).toHaveLength(1));
    const entry = wrapper.get('[data-testid="ai-activity-entry"]');
    expect(Number(entry.attributes("data-request-bytes"))).toBe(String(calls[0].init.body).length);
    expect(Number(entry.attributes("data-response-bytes"))).toBeGreaterThan(0);
    expect(find(wrapper, "ai-activity").text()).not.toContain("Reply with the single word OK.");

    expect(wrapper.findAll('[data-testid="ai-consent-grant"]')).toHaveLength(1);
    await find(wrapper, "ai-consent-revoke").trigger("click");
    await flushPromises();
    expect(wrapper.findAll('[data-testid="ai-consent-grant"]')).toHaveLength(0);

    await find(wrapper, "ai-activity-clear").trigger("click");
    await flushPromises();
    expect(find(wrapper, "ai-activity-empty").exists()).toBe(true);
    wrapper.unmount();
  });
});
