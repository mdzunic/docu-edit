// SPEC-146 — ShowTimer, the show timer's surface in both windows: testids from
// the prefix, the readout's role/data/label contract, the toggle naming its
// next action (no aria-pressed), Stop disabled while stopped, the polite live
// region, the editable mode group and length field, a HUD pill keeping focus
// off itself on a pointer press — and the overtime colours' contrast,
// computed from tokens.css for every palette in both schemes.
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { mount } from "@vue/test-utils";
import { PALETTE_IDS } from "../../lib/preferences";
import ShowTimer from "./ShowTimer.vue";

const get = (wrapper: ReturnType<typeof mount>, testid: string) => wrapper.get(`[data-testid="${testid}"]`);

describe("ShowTimer", () => {
  it("names everything from the prefix and renders the readout contract", () => {
    const wrapper = mount(ShowTimer, { props: { testidPrefix: "play", phase: "running", mode: "stopwatch", text: "1:15" } });
    const readout = get(wrapper, "play-timer");
    expect(readout.text()).toBe("1:15");
    expect(readout.attributes()).toMatchObject({
      role: "timer",
      "data-state": "running",
      "data-mode": "stopwatch",
      "data-overtime": "false",
      "aria-label": "Elapsed time",
    });
    for (const id of ["play-timer-toggle", "play-timer-stop", "play-timer-status"]) expect(wrapper.find(`[data-testid="${id}"]`).exists()).toBe(true);
    const status = get(wrapper, "play-timer-status");
    expect(status.attributes("aria-live")).toBe("polite");
    expect(status.classes()).toContain("sr-only");
  });

  it("labels the readout Time remaining in Countdown and Time over in overtime", async () => {
    const wrapper = mount(ShowTimer, { props: { phase: "running", mode: "countdown", text: "0:45" } });
    expect(get(wrapper, "presenter-timer").attributes("aria-label")).toBe("Time remaining");
    await wrapper.setProps({ text: "−0:42", overtime: true });
    const readout = get(wrapper, "presenter-timer");
    expect(readout.attributes("aria-label")).toBe("Time over");
    expect(readout.attributes("data-overtime")).toBe("true");
    expect(wrapper.classes()).toContain("is-overtime");
  });

  it("the toggle names its next action, with no aria-pressed; Stop is disabled while stopped", async () => {
    const wrapper = mount(ShowTimer, { props: { phase: "stopped" } });
    const toggle = () => get(wrapper, "presenter-timer-toggle");
    const stop = () => get(wrapper, "presenter-timer-stop");
    const rows: Array<[string, string, string, boolean]> = [
      ["stopped", "Start (T)", "Start timer", true],
      ["running", "Pause (T)", "Pause timer", false],
      ["paused", "Resume (T)", "Resume timer", false],
    ];
    for (const [phase, label, name, disabled] of rows) {
      await wrapper.setProps({ phase });
      expect(toggle().text()).toBe(label);
      expect(toggle().attributes("aria-label")).toBe(name);
      expect(toggle().attributes("aria-pressed")).toBeUndefined();
      expect(stop().text()).toBe("Stop (S)");
      expect(stop().attributes("aria-label")).toBe("Stop timer");
      expect(stop().attributes("disabled") !== undefined, phase).toBe(disabled);
    }
  });

  it("emits toggle and stop, holding no state of its own", async () => {
    const wrapper = mount(ShowTimer, { props: { phase: "running" } });
    await get(wrapper, "presenter-timer-toggle").trigger("click");
    await get(wrapper, "presenter-timer-stop").trigger("click");
    expect(Object.keys(wrapper.emitted())).toEqual(expect.arrayContaining(["toggle", "stop"]));
    expect(get(wrapper, "presenter-timer-toggle").text()).toBe("Pause (T)"); // unchanged until the props change
  });

  it("a click on a HUD button never bubbles to the stage, and a pointer press keeps focus off it", async () => {
    let stageClicks = 0;
    const host = document.createElement("div");
    host.addEventListener("click", () => (stageClicks += 1));
    document.body.append(host);
    const wrapper = mount(ShowTimer, { props: { testidPrefix: "play", phase: "running" }, attachTo: host });
    for (const id of ["play-timer-toggle", "play-timer-stop"]) {
      const button = get(wrapper, id);
      expect(button.element.tagName).toBe("BUTTON");
      expect(button.attributes("type")).toBe("button");
      const press = new MouseEvent("mousedown", { bubbles: true, cancelable: true });
      button.element.dispatchEvent(press);
      expect(press.defaultPrevented, id).toBe(true);
      await button.trigger("click");
    }
    expect(stageClicks).toBe(0);
    // The presenter's buttons focus normally (its keys leave a focused button's Enter/Space to it).
    const presenter = mount(ShowTimer, { props: { phase: "running" }, attachTo: host });
    const press = new MouseEvent("mousedown", { bubbles: true, cancelable: true });
    get(presenter, "presenter-timer-toggle").element.dispatchEvent(press);
    expect(press.defaultPrevented).toBe(false);
    wrapper.unmount();
    presenter.unmount();
    host.remove();
  });

  it("renders no mode or length controls unless editable", () => {
    const wrapper = mount(ShowTimer, { props: { testidPrefix: "play", mode: "countdown" } });
    expect(wrapper.find('[role="group"]').exists()).toBe(false);
    expect(wrapper.find("input").exists()).toBe(false);
  });

  it("editable: the Timer mode group with aria-pressed buttons, and the length field only in Countdown", async () => {
    const wrapper = mount(ShowTimer, { props: { editable: true, mode: "stopwatch", lengthMinutes: 20 } });
    const group = wrapper.get('[role="group"]');
    expect(group.attributes("aria-label")).toBe("Timer mode");
    expect(get(wrapper, "presenter-timer-mode-stopwatch").attributes("aria-pressed")).toBe("true");
    expect(get(wrapper, "presenter-timer-mode-countdown").attributes("aria-pressed")).toBe("false");
    expect(wrapper.find('[data-testid="presenter-timer-length"]').exists()).toBe(false);
    await get(wrapper, "presenter-timer-mode-countdown").trigger("click");
    expect(wrapper.emitted("mode")).toEqual([["countdown"]]);
    await wrapper.setProps({ mode: "countdown" });
    expect(get(wrapper, "presenter-timer-mode-countdown").attributes("aria-pressed")).toBe("true");
    const length = get(wrapper, "presenter-timer-length");
    expect(length.attributes()).toMatchObject({ type: "text", inputmode: "numeric", id: "presenter-timer-length" });
    expect(wrapper.get('label[for="presenter-timer-length"]').text()).toBe("Talk length (minutes)");
    expect((length.element as HTMLInputElement).value).toBe("20");
  });

  it("commits the length on change through parseCountdownMinutes, restoring anything that is not a number", async () => {
    const wrapper = mount(ShowTimer, { props: { editable: true, mode: "countdown", lengthMinutes: 20 } });
    const length = get(wrapper, "presenter-timer-length");
    const input = length.element as HTMLInputElement;
    for (const [typed, minutes] of [["25", 25], [" 7 ", 7], ["2.6", 3], ["0", 1], ["999", 240]] as const) {
      input.value = typed;
      await length.trigger("change");
      expect(input.value, typed).toBe(String(minutes));
    }
    expect(wrapper.emitted("length")).toEqual([[25], [7], [3], [1], [240]]);
    for (const typed of ["", "abc"]) {
      input.value = typed;
      await length.trigger("change");
      expect(input.value).toBe("20");
    }
    expect(wrapper.emitted("length")).toHaveLength(5);
  });
});

// --- overtime contrast (WCAG 1.4.3), computed from tokens.css ----------------
// The presenter reads overtime in --text-danger on its --surface-canvas; the
// HUD in white on the --danger family mixed 60% toward the HUD's black (the
// dark schemes' --danger is a light pink white text cannot sit on).
const here = dirname(fileURLToPath(import.meta.url));
const css = readFileSync(join(here, "../../styles/tokens.css"), "utf8");

function block(header: string): Record<string, string> {
  const at = css.indexOf(header);
  if (at < 0) return {};
  const body = css.slice(css.indexOf("{", at) + 1, css.indexOf("}", at));
  return Object.fromEntries([...body.matchAll(/(--[\w-]+):\s*([^;]+);/g)].map((m) => [m[1], m[2].trim()]));
}

function rgb(value: string): number[] {
  const hex = value.match(/#([0-9a-fA-F]{6})\b/);
  if (!hex) throw new Error(`not a #rrggbb colour: ${value}`);
  const n = parseInt(hex[1], 16);
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
}

function luminance(c: number[]): number {
  const [r, g, b] = c.map((v) => {
    const s = v / 255;
    return s <= 0.03928 ? s / 12.92 : ((s + 0.055) / 1.055) ** 2.4;
  });
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

function contrast(a: number[], b: number[]): number {
  const [la, lb] = [luminance(a), luminance(b)];
  return (Math.max(la, lb) + 0.05) / (Math.min(la, lb) + 0.05);
}

const WHITE = [255, 255, 255];
const mixToBlack = (c: number[], share: number) => c.map((v) => Math.round(v * share));

describe("ShowTimer overtime contrast", () => {
  const baseLight = block(":root {");
  const baseDark = { ...baseLight, ...block(':root[data-theme="dark"] {') };
  const schemes: Array<[string, Record<string, string>]> = [
    ["modernist light", baseLight],
    ["modernist dark", baseDark],
  ];
  for (const id of PALETTE_IDS.filter((palette) => palette !== "modernist")) {
    schemes.push([`${id} light`, { ...baseLight, ...block(`:root[data-palette="${id}"] {`) }]);
    schemes.push([`${id} dark`, { ...baseDark, ...block(`:root[data-palette="${id}"][data-theme="dark"] {`) }]);
  }
  for (const [name, tokens] of schemes) {
    it(`${name}: --text-danger on --surface-canvas, and white on the HUD's danger pill, ≥ 4.5:1`, () => {
      expect(contrast(rgb(tokens["--text-danger"]), rgb(tokens["--surface-canvas"]))).toBeGreaterThanOrEqual(4.5);
      expect(contrast(WHITE, mixToBlack(rgb(tokens["--danger"]), 0.6))).toBeGreaterThanOrEqual(4.5);
    });
  }
});
