import { createElement, type ReactNode } from 'react';

function accessLabelNode(node: Node, key: string): ReactNode {
  if (node.nodeType === Node.TEXT_NODE) return node.textContent;
  if (!(node instanceof Element)) return null;
  const children = Array.from(node.childNodes, (child, index) => accessLabelNode(child, `${key}.${index}`));
  return node.localName === 'u' ? createElement('u', { key }, ...children) : children;
}

/** Render the escaped `ui.ampHTML()` access-key label as React text and underline nodes. */
export function AccessLabel({ html }: { html: string }) {
  const parsed = new DOMParser().parseFromString(html, 'text/html');
  return <>{Array.from(parsed.body.childNodes, (node, index) => accessLabelNode(node, `label.${index}`))}</>;
}
