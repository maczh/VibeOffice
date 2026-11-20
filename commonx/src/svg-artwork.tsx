import { createElement, useMemo, type CSSProperties, type ReactNode } from 'react';

const svgAttributes: Record<string, string> = {
  'clip-path': 'clipPath', 'fill-opacity': 'fillOpacity', 'fill-rule': 'fillRule', 'font-family': 'fontFamily', 'font-size': 'fontSize', 'font-style': 'fontStyle', 'font-weight': 'fontWeight',
  'gradienttransform': 'gradientTransform', 'gradientunits': 'gradientUnits', 'marker-end': 'markerEnd', 'marker-mid': 'markerMid', 'marker-start': 'markerStart', 'markerheight': 'markerHeight', 'markerunits': 'markerUnits', 'markerwidth': 'markerWidth',
  'paint-order': 'paintOrder', 'pointer-events': 'pointerEvents', 'preserveaspectratio': 'preserveAspectRatio', 'refx': 'refX', 'refy': 'refY', 'shape-rendering': 'shapeRendering',
  'stop-color': 'stopColor', 'stop-opacity': 'stopOpacity', 'stroke-dasharray': 'strokeDasharray', 'stroke-dashoffset': 'strokeDashoffset', 'stroke-linecap': 'strokeLinecap', 'stroke-linejoin': 'strokeLinejoin', 'stroke-miterlimit': 'strokeMiterlimit', 'stroke-opacity': 'strokeOpacity', 'stroke-width': 'strokeWidth',
  'text-anchor': 'textAnchor', 'text-decoration': 'textDecoration', 'textlength': 'textLength', 'viewbox': 'viewBox', 'writing-mode': 'writingMode', 'xlink:href': 'xlinkHref', 'xml:space': 'xmlSpace',
};
const rejectedElements = new Set(['a', 'animate', 'animatecolor', 'animatemotion', 'animatetransform', 'audio', 'discard', 'embed', 'foreignobject', 'iframe', 'image', 'object', 'script', 'set', 'style', 'video']);

function attributeName(name: string): string {
  if (name === 'class') return 'className';
  if (name === 'tabindex') return 'tabIndex';
  return svgAttributes[name.toLowerCase()] || name;
}

function safeLocalReference(name: string, value: string): boolean {
  if (name !== 'href' && name !== 'xlink:href') return true;
  return value.startsWith('#');
}

function safeSvgValue(value: string): boolean {
  if (/(?:javascript|vbscript|data)\s*:/i.test(value)) return false;
  for (const match of value.matchAll(/url\s*\(([^)]*)\)/gi)) {
    const reference = (match[1] || '').trim().replace(/^(['"])(.*)\1$/, '$2').trim();
    if (!reference.startsWith('#')) return false;
  }
  return true;
}

function reactStyle(element: SVGElement): CSSProperties {
  const style: Record<string, string> = {};
  for (let index = 0; index < element.style.length; index++) {
    const property = element.style.item(index);
    const value = element.style.getPropertyValue(property);
    if (/url\s*\(|expression\s*\(/i.test(value)) continue;
    const name = property.startsWith('--') ? property : property.replace(/^-ms-/, 'ms-').replace(/-([a-z])/g, (_match, letter: string) => letter.toUpperCase());
    style[name] = value;
  }
  return style as CSSProperties;
}

function svgNode(node: Node, key: string, owner: string, root = false): ReactNode {
  if (node.nodeType === Node.TEXT_NODE) return node.textContent;
  if (!(node instanceof Element) || rejectedElements.has(node.localName.toLowerCase())) return null;
  const props: Record<string, unknown> = { key };
  for (const attribute of Array.from(node.attributes)) {
    if (/^on/i.test(attribute.name) || !safeLocalReference(attribute.name, attribute.value) || !safeSvgValue(attribute.value)) continue;
    if (attribute.name === 'style' && node instanceof SVGElement) {
      props.style = reactStyle(node);
      continue;
    }
    props[attributeName(attribute.name)] = attribute.value;
  }
  if (root) props['data-react-owned'] = owner;
  const children = Array.from(node.childNodes, (child, index) => svgNode(child, `${key}.${index}`, owner));
  return createElement(node.localName, props, ...children);
}

/** Reconcile a single engine-generated SVG as React elements; reject HTML roots and active/external content. */
export function SvgArtwork({ markup, owner = 'shared-svg-artwork' }: { markup: string; owner?: string }) {
  const artwork = useMemo(() => {
    const parsed = new DOMParser().parseFromString(markup, 'text/html');
    const root = parsed.body.firstElementChild;
    if (parsed.body.children.length !== 1 || !root || root.namespaceURI !== 'http://www.w3.org/2000/svg' || root.localName.toLowerCase() !== 'svg') return null;
    return svgNode(root, 'svg', owner, true);
  }, [markup, owner]);
  return artwork;
}
