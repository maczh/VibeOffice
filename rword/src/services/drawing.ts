import type { ImageRef, Run } from '../engine/model'
export function imageRun(image: ImageRef, alt = ''): Run { return { image: { ...image }, rPr: {}, text: alt || undefined } }
export function imageSize(width: number, height: number, maxWidth = 640): { width: number; height: number } { if (width <= maxWidth) return { width, height }; const ratio = maxWidth / width; return { width: maxWidth, height: Math.max(1, Math.round(height * ratio)) } }
export function mediaType(name: string): string { const ext = name.split('.').pop()?.toLowerCase(); return ({ png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', webp: 'image/webp', svg: 'image/svg+xml' } as Record<string, string>)[ext ?? ''] ?? 'application/octet-stream' }
