// Duplicating an .ods sheet's retained source (SPEC-121 "Duplicating a
// sheet"), at command time and copy-on-write: the copy's `sourceXml` is the
// origin's with every `xml:id`/`draw:id` value suffixed `-c<k>` consistently
// (the attributes pointing at them included; k the smallest that collides
// with no id in the workbook), and each embedded object it shows
// (`draw:object` → `Object N/`, its `ObjectReplacements/Object N` image) is
// copied in `opaque` to the next free `Object N` with the manifest spliced
// — a new manifest part, never one edited in place. Pictures stay shared.
// SPEC-131: the copy's drawings' source XML (`drawingTexts`, each a draw:*
// element also inside the table text) takes the same suffix and retargets.
// The copy's `table:name` stays the origin's source name until the writer
// splices it on save. Deleting a copy drops the objects it alone showed
// (removeCopyObjects).
import { base64ToBytes, bytesToBase64, type OpaquePart } from "../../core/packageParts";
import type { SpreadsheetModel } from "../model";
import {
  declaredIds,
  isObjectPart,
  manifestEntryLike,
  nextFreeObjectName,
  objectsReferenced,
  readManifest,
  retargetObject,
  spliceManifest,
  suffixIds,
} from "./package";

const MANIFEST = "META-INF/manifest.xml";
const utf8 = new TextDecoder("utf-8");
const encoder = new TextEncoder();

/** Every id the workbook's retained text declares. */
function workbookIds(model: SpreadsheetModel): Set<string> {
  const ids = new Set<string>();
  const texts = [model.odsContentEnvelope?.prefix ?? "", model.odsContentEnvelope?.suffix ?? "", ...model.sheets.map((sheet) => sheet.sourceXml ?? "")];
  for (const text of texts) {
    for (const id of declaredIds(text)) {
      ids.add(id);
    }
  }
  return ids;
}

/** The copy's retained source and `opaque` (see the module header). */
export function copyOdsSheetSource(
  model: SpreadsheetModel,
  originXml: string,
  drawingTexts: readonly string[] = [],
): { sourceXml: string; opaque: OpaquePart[]; drawingTexts: string[] } {
  const own = new Set([...declaredIds(originXml), ...drawingTexts.flatMap((text) => [...declaredIds(text)])]);
  const taken = workbookIds(model);
  let k = 1;
  while ([...own].some((id) => taken.has(`${id}-c${k}`))) {
    k += 1;
  }
  let sourceXml = suffixIds(originXml, `-c${k}`);
  let texts = drawingTexts.map((text) => (text === "" ? text : suffixIds(text, `-c${k}`)));

  const objects = [...new Set([...objectsReferenced(sourceXml), ...texts.flatMap((text) => objectsReferenced(text))])];
  if (objects.length === 0) {
    return { sourceXml, opaque: model.opaque, drawingTexts: texts };
  }
  const names = model.opaque.map((part) => part.name);
  const allocated = new Set<string>();
  const added: OpaquePart[] = [];
  const manifestPart = model.opaque.find((part) => part.name === MANIFEST);
  const manifestXml = manifestPart ? utf8.decode(base64ToBytes(manifestPart.data)) : undefined;
  const entries = manifestXml === undefined ? [] : readManifest(manifestXml);
  const newEntries: string[] = [];
  for (const object of objects) {
    if (!names.some((name) => isObjectPart(name, object))) {
      continue;
    }
    const copy = nextFreeObjectName(names, allocated);
    allocated.add(copy);
    for (const part of model.opaque) {
      if (isObjectPart(part.name, object)) {
        added.push({ name: part.name.replace(object, copy), data: part.data });
      }
    }
    for (const entry of entries) {
      if (entry.fullPath === `${object}/` || entry.fullPath.startsWith(`${object}/`) || entry.fullPath === `ObjectReplacements/${object}`) {
        newEntries.push(manifestEntryLike(manifestXml!, entry, entry.fullPath.replace(object, copy)));
      }
    }
    sourceXml = retargetObject(sourceXml, object, copy);
    texts = texts.map((text) => (text === "" ? text : retargetObject(text, object, copy)));
  }
  let opaque: OpaquePart[] = [...model.opaque, ...added];
  if (manifestXml !== undefined && newEntries.length > 0) {
    const next = spliceManifest(manifestXml, new Set(), newEntries);
    opaque = opaque.map((part) => (part.name === MANIFEST ? { name: MANIFEST, data: bytesToBase64(encoder.encode(next)) } : part));
  }
  return { sourceXml, opaque, drawingTexts: texts };
}

/** `opaque` without the embedded objects only a deleted copy's text showed
 * (`remaining` is every other retained text), their manifest entries
 * spliced out. */
export function removeCopyObjects(opaque: readonly OpaquePart[], copyXml: string, remaining: readonly string[]): readonly OpaquePart[] {
  const stillShown = new Set(remaining.flatMap((text) => objectsReferenced(text)));
  const orphaned = objectsReferenced(copyXml).filter((object) => !stillShown.has(object));
  if (orphaned.length === 0) {
    return opaque;
  }
  const dropped = (name: string): boolean => orphaned.some((object) => isObjectPart(name, object));
  const manifestPart = opaque.find((part) => part.name === MANIFEST);
  const kept = opaque.filter((part) => !dropped(part.name));
  if (!manifestPart) {
    return kept;
  }
  const manifestXml = utf8.decode(base64ToBytes(manifestPart.data));
  const remove = new Set(
    readManifest(manifestXml)
      .map((entry) => entry.fullPath)
      .filter((path) => orphaned.some((object) => path === `${object}/` || path.startsWith(`${object}/`) || path === `ObjectReplacements/${object}`)),
  );
  const next = spliceManifest(manifestXml, remove, []);
  return kept.map((part) => (part.name === MANIFEST ? { name: MANIFEST, data: bytesToBase64(encoder.encode(next)) } : part));
}
