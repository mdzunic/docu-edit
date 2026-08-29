// LibreOffice's group shape (`draw:g`), flattened into the model's flat
// `groupOrdinal` grouping — the .odt twin of docx/group.spec.ts. See
// group.ts's header for why flattening rather than a new container block,
// and what that costs.
import { describe, it, expect } from "vitest";
import { parseOdt, type ParsedOdtDocument } from "./reader";
import { writeOdt } from "./writer";
import { zip } from "../core/zipWriter";
import type { DocumentModel, Image, Shape } from "../documentModel";

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const NAMESPACES =
  ' xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"' +
  ' xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"' +
  ' xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"' +
  ' xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"' +
  ' xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"' +
  ' xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"' +
  ' xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"' +
  ' xmlns:xlink="http://www.w3.org/1999/xlink"' +
  ' xmlns:de="urn:docu-edit:shape"' +
  ' office:version="1.2"';

/** One automatic graphic style, the way LibreOffice writes a member's. */
function graphicStyle(name: string, fill: string): string {
  return (
    `<style:style style:name="${name}" style:family="graphic">` +
    `<style:graphic-properties draw:fill="solid" draw:fill-color="${fill}"` +
    ' draw:stroke="solid" svg:stroke-color="#000000" svg:stroke-width="0.05cm"/>' +
    "</style:style>"
  );
}

interface MemberOptions {
  style: string;
  xCm: number;
  yCm: number;
  widthCm?: number;
  heightCm?: number;
  desc: string;
  drawType?: string;
}

/** A `draw:custom-shape` group member: its `svg:x`/`svg:y` are already in the
 * group's own frame — ODF has no `chOff`/`chExt` twin. */
function member(options: MemberOptions): string {
  const { style, xCm, yCm, widthCm = 2, heightCm = 2, desc, drawType = "rectangle" } = options;
  return (
    `<draw:custom-shape draw:style-name="${style}" svg:x="${xCm}cm" svg:y="${yCm}cm"` +
    ` svg:width="${widthCm}cm" svg:height="${heightCm}cm">` +
    `<svg:desc>${desc}</svg:desc><draw:enhanced-geometry draw:type="${drawType}"/>` +
    "</draw:custom-shape>"
  );
}

function group(members: string, attrs = ' text:anchor-type="paragraph"'): string {
  return `<draw:g draw:name="Group"${attrs}>${members}</draw:g>`;
}

const TWO_MEMBERS =
  member({ style: "gA", xCm: 1, yCm: 2, widthCm: 3, desc: "member A" }) +
  member({ style: "gB", xCm: 5, yCm: 3, desc: "member B", drawType: "ellipse" });

const PNG_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

function base64ToBytes(base64: string): Uint8Array {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

/** A `draw:frame` picture group member, referencing the package's one image. */
function picMember(xCm: number, yCm: number, desc: string): string {
  return (
    `<draw:frame svg:x="${xCm}cm" svg:y="${yCm}cm" svg:width="2cm" svg:height="2cm">` +
    '<draw:image xlink:href="Pictures/image1.png"/>' +
    `<svg:desc>${desc}</svg:desc></draw:frame>`
  );
}

async function odtBytes(bodyContent: string, extraStyles = ""): Promise<Uint8Array> {
  const encoder = new TextEncoder();
  const content =
    `${XML}<office:document-content${NAMESPACES}><office:automatic-styles>` +
    graphicStyle("gA", "#4472c4") +
    graphicStyle("gB", "#70ad47") +
    extraStyles +
    "</office:automatic-styles><office:body><office:text>" +
    bodyContent +
    "<text:p>end</text:p></office:text></office:body></office:document-content>";
  return zip([
    { name: "mimetype", data: encoder.encode("application/vnd.oasis.opendocument.text"), store: true },
    { name: "content.xml", data: encoder.encode(content) },
    { name: "styles.xml", data: encoder.encode(`${XML}<office:document-styles${NAMESPACES}/>`) },
    { name: "Pictures/image1.png", data: base64ToBytes(PNG_BASE64) },
    {
      name: "META-INF/manifest.xml",
      data: encoder.encode(
        `${XML}<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"` +
          ' manifest:version="1.2"><manifest:file-entry manifest:full-path="/"' +
          ' manifest:media-type="application/vnd.oasis.opendocument.text"/></manifest:manifest>',
      ),
    },
  ]);
}

async function parse(bodyContent: string, extraStyles = ""): Promise<ParsedOdtDocument> {
  return parseOdt(await odtBytes(bodyContent, extraStyles));
}

async function blocksOf(bodyContent: string, extraStyles = "") {
  return (await parse(bodyContent, extraStyles)).sections[0].blocks;
}

function preservedXml(block: { type: string }): string {
  const paragraph = block as { type: string; opaqueInlines?: { xml: string }[] };
  return paragraph.type === "paragraph" ? (paragraph.opaqueInlines ?? []).map((i) => i.xml).join("") : "";
}

describe("parseOdt — a draw:g group flattens into its members", () => {
  it("yields one block per member at the member's own coordinates", async () => {
    const blocks = await blocksOf(`<text:p>${group(TWO_MEMBERS)}</text:p>`);
    expect(blocks.map((block) => block.type)).toEqual(["shape", "shape", "paragraph"]);
    const [first, second] = blocks as [Shape, Shape];
    // A `draw:g` has no child coordinate space, so `svg:x`/`svg:y` need no
    // mapping — unlike OOXML's `chOff`/`chExt`.
    expect(first.offsetXCm).toBeCloseTo(1, 6);
    expect(first.offsetYCm).toBeCloseTo(2, 6);
    expect(second.offsetXCm).toBeCloseTo(5, 6);
    expect(second.offsetYCm).toBeCloseTo(3, 6);
    expect(first.widthCm).toBeCloseTo(3, 6);
    expect(second.widthCm).toBeCloseTo(2, 6);
  });

  it("gives every member the same ordinal and reads each one's own appearance", async () => {
    const blocks = await blocksOf(`<text:p>${group(TWO_MEMBERS)}</text:p>`);
    const [first, second] = blocks as [Shape, Shape];
    expect(first.groupOrdinal).toBe(1);
    expect(second.groupOrdinal).toBe(1);
    expect(first.geometry).toBe("rect");
    expect(second.geometry).toBe("ellipse");
    expect(first.fill).toBe("#4472c4");
    expect(second.fill).toBe("#70ad47");
    expect(first.altText).toBe("member A");
    expect(second.altText).toBe("member B");
  });

  it("floats the members — a member's own style carries no anchor of its own", async () => {
    const blocks = await blocksOf(`<text:p>${group(TWO_MEMBERS)}</text:p>`);
    const [first] = blocks as [Shape];
    expect(first.wrap).toBe("infront");
    expect(first.relativeToH).toBe("column");
    expect(first.relativeToV).toBe("paragraph");
  });

  it("floats an as-char group's members rather than putting each in the text flow", async () => {
    // The model's grouping is floating-only (shapes.ts's groupShapes converts
    // an inline member the same way).
    const blocks = await blocksOf(`<text:p>${group(TWO_MEMBERS, ' text:anchor-type="as-char"')}</text:p>`);
    const [first] = blocks as [Shape];
    expect(first.wrap).toBe("square");
    expect(first.groupOrdinal).toBe(1);
  });

  it("flattens a picture member beside a shape one", async () => {
    // A `draw:frame` member took the whole group down with it until Image
    // gained real placement: an image had no position anywhere in the model,
    // so a flattened one would have painted at the group's origin.
    const mixed = group(member({ style: "gA", xCm: 1, yCm: 2, desc: "member A" }) + picMember(5, 3, "picture B"));
    const parsed = await parse(`<text:p>${mixed}</text:p>`);
    const blocks = parsed.sections[0].blocks;
    expect(blocks.map((block) => block.type)).toEqual(["shape", "image", "paragraph"]);
    const shape = blocks[0] as Shape;
    const image = blocks[1] as Image;
    expect(image.groupOrdinal).toBe(shape.groupOrdinal);
    expect(image.offsetXCm).toBeCloseTo(5, 6);
    expect(image.offsetYCm).toBeCloseTo(3, 6);
    expect(image.altText).toBe("picture B");
    expect(parsed.media.map((part) => part.mediaId)).toContain("Pictures/image1.png");
  });

  it("skips the group's own svg:desc rather than treating it as a member", async () => {
    const withDesc = `<draw:g draw:name="Group" text:anchor-type="paragraph"><svg:desc>The group</svg:desc>${TWO_MEMBERS}</draw:g>`;
    const blocks = await blocksOf(`<text:p>${withDesc}</text:p>`);
    expect(blocks.map((block) => block.type)).toEqual(["shape", "shape", "paragraph"]);
  });
});

describe("parseOdt — draw:g transforms", () => {
  it("adds a translate() transform to every member's offset", async () => {
    const translated = group(TWO_MEMBERS, ' text:anchor-type="paragraph" draw:transform="translate(2cm 1cm)"');
    const [first, second] = (await blocksOf(`<text:p>${translated}</text:p>`)) as [Shape, Shape];
    expect(first.offsetXCm).toBeCloseTo(3, 6);
    expect(first.offsetYCm).toBeCloseTo(3, 6);
    expect(second.offsetXCm).toBeCloseTo(7, 6);
    expect(second.offsetYCm).toBeCloseTo(4, 6);
  });

  it("composes a nested group's translation with its parent's, into one ordinal", async () => {
    const nested =
      '<draw:g draw:name="Outer" text:anchor-type="paragraph" draw:transform="translate(1cm 1cm)">' +
      member({ style: "gA", xCm: 0, yCm: 0, desc: "outer member" }) +
      '<draw:g draw:name="Inner" draw:transform="translate(2cm 0cm)">' +
      member({ style: "gB", xCm: 1, yCm: 1, desc: "inner member" }) +
      "</draw:g></draw:g>";
    const blocks = await blocksOf(`<text:p>${nested}</text:p>`);
    expect(blocks.map((block) => block.type)).toEqual(["shape", "shape", "paragraph"]);
    const [outer, inner] = blocks as [Shape, Shape];
    expect(outer.offsetXCm).toBeCloseTo(1, 6);
    expect(inner.offsetXCm).toBeCloseTo(1 + 2 + 1, 6);
    expect(inner.offsetYCm).toBeCloseTo(1 + 1, 6);
    // The flat model cannot express a group inside a group.
    expect(inner.groupOrdinal).toBe(outer.groupOrdinal);
  });

  it("preserves a group whose transform is not a plain translation", async () => {
    // A rotate/skew/matrix would move its members somewhere this flat model
    // cannot express, so the whole group keeps its verbatim preservation
    // rather than being drawn at the untransformed coordinates.
    const rotated = group(TWO_MEMBERS, ' text:anchor-type="paragraph" draw:transform="rotate(0.5) translate(2cm 1cm)"');
    const blocks = await blocksOf(`<text:p>${rotated}</text:p>`);
    expect(blocks.map((block) => block.type)).toEqual(["paragraph", "paragraph"]);
    expect(preservedXml(blocks[0])).toContain("draw:g");
  });
});

describe("parseOdt — draw:g ordinals", () => {
  it("gives two groups in one document different ordinals", async () => {
    const body = `<text:p>${group(TWO_MEMBERS)}</text:p><text:p>${group(TWO_MEMBERS)}</text:p>`;
    const ordinals = (await blocksOf(body))
      .filter((block) => block.type === "shape")
      .map((block) => (block as Shape).groupOrdinal);
    expect(ordinals).toEqual([1, 1, 2, 2]);
  });

  it("mints above every de:group the document already carries", async () => {
    // Both group flavours coexist once a user groups shapes in LibreOffice
    // inside a file this editor saved; minting from 1 would merge them.
    const existing =
      '<draw:custom-shape draw:style-name="gA" text:anchor-type="paragraph" de:group="4"' +
      ' svg:x="0cm" svg:y="0cm" svg:width="2cm" svg:height="2cm">' +
      "<svg:desc>existing</svg:desc><draw:enhanced-geometry draw:type=\"rectangle\"/></draw:custom-shape>";
    const body = `<text:p>${existing}</text:p><text:p>${group(TWO_MEMBERS)}</text:p>`;
    const ordinals = (await blocksOf(body))
      .filter((block) => block.type === "shape")
      .map((block) => (block as Shape).groupOrdinal);
    expect(ordinals).toEqual([4, 5, 5]);
  });
});

describe("parseOdt — a draw:g nothing can fully read stays preserved", () => {
  // All-or-nothing, like the .docx side: reading the members we understand
  // and dropping the rest would discard document content (SPEC-008).
  it("preserves a group whose members carry no coordinates to place them by", async () => {
    const noCoords =
      '<draw:g draw:name="Group" text:anchor-type="paragraph">' +
      '<draw:custom-shape draw:style-name="gA" svg:width="2cm" svg:height="2cm">' +
      '<draw:enhanced-geometry draw:type="rectangle"/></draw:custom-shape></draw:g>';
    expect((await blocksOf(`<text:p>${noCoords}</text:p>`)).map((block) => block.type)).toEqual([
      "paragraph",
      "paragraph",
    ]);
  });

  it("preserves an empty group rather than erasing it", async () => {
    const blocks = await blocksOf(`<text:p>${group("")}</text:p>`);
    expect(blocks.map((block) => block.type)).toEqual(["paragraph", "paragraph"]);
  });
});

describe("writeOdt — a flattened draw:g survives save and reopen", () => {
  it("keeps both members grouped and positioned through a round trip", async () => {
    const parsed = await parse(`<text:p>${group(TWO_MEMBERS)}</text:p>`);
    const model: DocumentModel = {
      id: "test-doc",
      title: "Groups",
      sections: parsed.sections,
      defaultStyles: parsed.defaultStyles,
      format: "odt",
      opaque: parsed.opaque,
      sourceXml: parsed.sourceXml,
    };
    const { bytes } = await writeOdt(model);
    const reread = await parseOdt(bytes);
    const blocks = reread.sections[0].blocks;
    expect(blocks.map((block) => block.type)).toEqual(["shape", "shape", "paragraph"]);
    const [first, second] = blocks as [Shape, Shape];
    expect(first.groupOrdinal).toBe(1);
    expect(second.groupOrdinal).toBe(1);
    expect(first.offsetXCm).toBeCloseTo(1, 4);
    expect(second.offsetXCm).toBeCloseTo(5, 4);
    expect(second.offsetYCm).toBeCloseTo(3, 4);
  });
});
