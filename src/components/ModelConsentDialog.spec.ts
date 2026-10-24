// SPEC-113 — the consent dialog and the bridge that shows it: the exact
// destination and whether it is this machine, the exact payload with its
// byte count, the three answers, and a prompt that fails closed.
import { flushPromises, mount } from "@vue/test-utils";
import { defineComponent, h } from "vue";
import { describe, expect, it } from "vitest";
import ModelConsentDialog from "./ModelConsentDialog.vue";
import { askModelConsent, useModelConsentHost } from "../composables/useModelConsent";
import type { ConsentRequest } from "../lib/ai/modelGateway";

const body = JSON.stringify({ model: "m", messages: [{ role: "user", content: "Zürich notes" }] }, null, 2);

function consentRequest(patch: Partial<ConsentRequest> = {}): ConsentRequest {
  return {
    origin: "http://localhost:11434",
    local: true,
    purpose: "Test connection",
    model: "m",
    method: "POST",
    url: "http://localhost:11434/v1/chat/completions",
    body,
    bytes: new TextEncoder().encode(body).byteLength,
    sendsKey: false,
    ...patch,
  };
}

describe("ModelConsentDialog", () => {
  it("says plainly that a loopback destination stays on this computer", () => {
    const wrapper = mount(ModelConsentDialog, { props: { request: consentRequest() } });
    const destination = wrapper.get('[data-testid="ai-consent-destination"]').text();
    expect(destination).toContain("This stays on your computer");
    expect(destination).toContain("http://localhost:11434");
  });

  it("names a third-party origin as not part of this product", () => {
    const wrapper = mount(ModelConsentDialog, {
      props: { request: consentRequest({ origin: "https://models.example.test", local: false, sendsKey: true }) },
    });
    expect(wrapper.get('[data-testid="ai-consent-destination"]').text()).toBe(
      "Your text will be sent to https://models.example.test, which is not part of this product.",
    );
    expect(wrapper.text()).toContain("Your API key goes with it");
  });

  it("shows the exact body in a scrollable, focusable region with its byte count", () => {
    const wrapper = mount(ModelConsentDialog, { props: { request: consentRequest() } });
    const payload = wrapper.get('[data-testid="ai-consent-payload"]');
    expect(payload.element.tagName).toBe("PRE");
    expect(payload.element.textContent).toBe(body);
    expect(payload.attributes("tabindex")).toBe("0");
    const bytes = new TextEncoder().encode(body).byteLength;
    expect(bytes).toBeGreaterThan(body.length); // "ü" is two bytes: bytes, not characters
    expect(wrapper.get('[data-testid="ai-consent-bytes"]').text()).toBe(`Exactly this will be sent — ${bytes} bytes:`);
    expect(wrapper.get('[data-testid="ai-consent-purpose"]').text()).toBe("Test connection");
  });

  it("says so when there is no body (a GET)", () => {
    const wrapper = mount(ModelConsentDialog, { props: { request: consentRequest({ method: "GET", body: "", bytes: 0 }) } });
    expect(wrapper.get('[data-testid="ai-consent-payload"]').text()).toBe("(no request body)");
  });

  it("offers Send once, Always allow and Cancel — and Escape is Cancel", async () => {
    const wrapper = mount(ModelConsentDialog, { props: { request: consentRequest() } });
    await wrapper.get('[data-testid="ai-consent-once"]').trigger("click");
    await wrapper.get('[data-testid="ai-consent-always"]').trigger("click");
    await wrapper.get('[data-testid="ai-consent-cancel"]').trigger("click");
    await wrapper.get('[role="dialog"]').trigger("keydown", { key: "Escape" });
    expect(wrapper.emitted("answer")).toEqual([["once"], ["always"], ["cancel"], ["cancel"]]);
  });
});

describe("useModelConsent", () => {
  const Host = defineComponent({
    setup() {
      const consent = useModelConsentHost();
      return () =>
        consent.request.value
          ? h(ModelConsentDialog, { request: consent.request.value, onAnswer: consent.answer })
          : h("p", "idle");
    },
  });

  it("fails closed with no shell mounted to ask", async () => {
    expect(await askModelConsent(consentRequest())).toBe("cancel");
  });

  it("shows the request in the hosted dialog and resolves with the answer", async () => {
    const host = mount(Host);
    const answer = askModelConsent(consentRequest());
    await flushPromises();
    expect(host.find('[data-testid="ai-consent"]').exists()).toBe(true);
    await host.get('[data-testid="ai-consent-always"]').trigger("click");
    expect(await answer).toBe("always");
    await flushPromises();
    expect(host.find('[data-testid="ai-consent"]').exists()).toBe(false);
    host.unmount();
  });

  it("refuses a second request while one is on screen, and cancels on the caller's abort", async () => {
    const host = mount(Host);
    const controller = new AbortController();
    const first = askModelConsent(consentRequest(), controller.signal);
    expect(await askModelConsent(consentRequest({ purpose: "other" }))).toBe("cancel");
    controller.abort();
    expect(await first).toBe("cancel");
    await flushPromises();
    expect(host.find('[data-testid="ai-consent"]').exists()).toBe(false);
    host.unmount();
  });

  it("answers cancel when the shell unmounts with a request open", async () => {
    const host = mount(Host);
    const pending = askModelConsent(consentRequest());
    host.unmount();
    expect(await pending).toBe("cancel");
  });
});
