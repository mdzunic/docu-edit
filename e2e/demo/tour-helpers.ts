import type { Page } from "@playwright/test";

// Helpers for the headed demo tour (demo.tour.ts). Everything here is
// presentation-only: an on-screen caption bar so a viewer can follow along,
// and pacing utilities that keep the tour watchable. Nothing in this file
// asserts product behaviour.

const CAPTION_ID = "__demo_caption";

/**
 * Show (or update) a floating caption pill at the bottom of the viewport.
 * Survives SPA route changes; recreated automatically after a full page load.
 */
export async function caption(page: Page, text: string): Promise<void> {
  await page.evaluate(
    ({ id, text }) => {
      let el = document.getElementById(id);
      if (!el) {
        el = document.createElement("div");
        el.id = id;
        el.style.cssText = [
          "position:fixed",
          "left:50%",
          "bottom:28px",
          "transform:translateX(-50%)",
          "max-width:80vw",
          "padding:14px 26px",
          "border-radius:999px",
          "background:rgba(17,24,39,0.92)",
          "color:#fff",
          "font:600 20px/1.3 system-ui,sans-serif",
          "letter-spacing:0.01em",
          "box-shadow:0 8px 30px rgba(0,0,0,0.35)",
          "z-index:2147483647",
          "pointer-events:none",
          "transition:opacity 200ms ease",
          "text-align:center",
        ].join(";");
        document.body.appendChild(el);
      }
      el.textContent = text;
      el.style.opacity = "1";
    },
    { id: CAPTION_ID, text },
  );
}

/** Full-screen title card, shown briefly at the start and end of the tour. */
export async function titleCard(page: Page, title: string, subtitle: string, ms: number): Promise<void> {
  await page.evaluate(
    ({ title, subtitle }) => {
      const el = document.createElement("div");
      el.id = "__demo_title_card";
      el.style.cssText = [
        "position:fixed",
        "inset:0",
        "display:flex",
        "flex-direction:column",
        "align-items:center",
        "justify-content:center",
        "gap:14px",
        "background:rgba(17,24,39,0.96)",
        "color:#fff",
        "z-index:2147483647",
        "pointer-events:none",
        "transition:opacity 350ms ease",
      ].join(";");
      const h = document.createElement("div");
      h.textContent = title;
      h.style.cssText = "font:700 52px/1.1 system-ui,sans-serif;letter-spacing:-0.02em";
      const s = document.createElement("div");
      s.textContent = subtitle;
      s.style.cssText = "font:400 22px/1.4 system-ui,sans-serif;opacity:0.75";
      el.append(h, s);
      document.body.appendChild(el);
    },
    { title, subtitle },
  );
  await page.waitForTimeout(ms);
  await page.evaluate(() => {
    const el = document.getElementById("__demo_title_card");
    if (el) {
      el.style.opacity = "0";
      setTimeout(() => el.remove(), 400);
    }
  });
}

/** A short viewer-paced pause between actions. */
export async function beat(page: Page, ms = 800): Promise<void> {
  await page.waitForTimeout(ms);
}

/**
 * Debug-only still frames: with DEMO_SHOTS=1 in the environment, saves a
 * full-page screenshot to demo-results/shots/<nn>-<name>.png. A no-op in a
 * normal demo run, so the tour's pacing is unaffected.
 */
let shotCounter = 0;
export async function shot(page: Page, name: string): Promise<void> {
  if (!process.env.DEMO_SHOTS) {
    return;
  }
  shotCounter += 1;
  const nn = String(shotCounter).padStart(2, "0");
  await page.screenshot({ path: `demo-results/shots/${nn}-${name}.png` });
}
