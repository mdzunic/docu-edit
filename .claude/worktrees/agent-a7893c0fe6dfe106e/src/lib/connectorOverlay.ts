// SPEC-070: the connector overlay's plain-data shapes and typed `no*` seed
// helpers for ConnectorOverlay.vue's prop defaults (the `.vue` script's espree
// parser bars a type annotation on a default, the same idiom the other `no*`
// seeds in lib/ follow). Framework-free — the view composable (useConnectors)
// fills these in from DOM measurement and the overlay component paints them.
import type { ArrowheadKind, ShapeOutline } from "./core/shapes";
import type { ConnectorKind } from "./core/connectors";
import type { Point } from "./core/shapeGeometry";

/** Where a connector lives, in `{section, block}` form. */
export interface ConnectorRenderLocation {
  section: number;
  block: number;
}

/** One connector's measured geometry, ready for the overlay to paint. All
 * coordinates are unscaled page-relative px (the `.shape-handles` frame). */
export interface ConnectorRender {
  location: ConnectorRenderLocation;
  key: string;
  kind: ConnectorKind;
  /** SVG `d` for the visible line and the wide transparent hit path. */
  d: string;
  selected: boolean;
  outline?: ShapeOutline;
  startArrow?: ArrowheadKind;
  endArrow?: ArrowheadKind;
  a: Point;
  b: Point;
  /** Midpoint adjustment handle position (elbow/curved only). */
  mid?: Point;
}

/** A rendered set of connection-site dots on a hovered shape. */
export interface SiteHint {
  location: ConnectorRenderLocation;
  sites: Point[];
}

export function noConnectorRenders(): ConnectorRender[] {
  return [];
}

export function noSiteHint(): SiteHint | null {
  return null;
}

export function noDrawFrom(): { point: Point } | null {
  return null;
}

export function noDrawPoint(): Point | null {
  return null;
}

/** Typed pass-throughs for ConnectorOverlay.vue's runtime-declared props (the
 * `.vue` script cannot annotate them — same idiom as pageRender's
 * asRenderEntries). */
export function asConnectorRenders(value: unknown): ConnectorRender[] {
  return value as ConnectorRender[];
}

export function asSiteHint(value: unknown): SiteHint | null {
  return value as SiteHint | null;
}

export function asDrawFrom(value: unknown): { point: Point } | null {
  return value as { point: Point } | null;
}

export function asDrawPoint(value: unknown): Point | null {
  return value as Point | null;
}
