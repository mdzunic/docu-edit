// SPEC-045 AC-21: a letter with a letterhead image. Run via
// `node examples/letter.js`.
//
// The letterhead bytes below are a tiny embedded PNG so this example has no
// external asset file to ship or find at runtime — a real integration would
// read actual letterhead artwork instead, e.g.
// `await readFile(new URL("./letterhead.png", import.meta.url))`, and pass
// the resulting Buffer straight to addImage's `data` (a Buffer/Uint8Array
// works exactly like the base64 string used here).
import { writeFile, mkdir } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { DocumentBuilder } from "../../src/lib/generate/builder";

// A small solid-colour PNG (120x24) standing in for a company letterhead
// banner graphic.
const LETTERHEAD_PNG_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAHgAAAAYCAIAAAC+8q7fAAAASUlEQVR4nO3QAQkAIADAMPOYyU5WtoXCHTzA2Zhr60Lj+cEngQbdCjToVqBBtwINuhVo0K1Ag24FGnQr0KBbgQbdCjToVqBBtzqO1Cz13PUFaAAAAABJRU5ErkJggg==";

const builder = new DocumentBuilder({ title: "Letter to Acme Robotics" });

builder
  .addImage({ data: LETTERHEAD_PNG_BASE64, contentType: "image/png", altText: "Company letterhead" })
  .addParagraph("August 11, 2026")
  .addParagraph("Acme Robotics\n42 Industrial Way\nSpringfield")
  .addParagraph("Dear Acme Robotics,")
  .addParagraph(
    "Thank you for your continued partnership. This letter confirms the renewal of your annual " +
      "support contract, effective September 1st.",
  )
  .addParagraph("Sincerely,")
  .addParagraph("The docu-edit Team");

const outDir = fileURLToPath(new URL("../output/", import.meta.url));
await mkdir(outDir, { recursive: true });
const outPath = fileURLToPath(new URL("../output/letter.docx", import.meta.url));
await writeFile(outPath, await builder.toBuffer("docx"));
console.log(`Wrote ${outPath}`);
