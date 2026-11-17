/* UI strings are written in English in the code; t() fills {0}, {1}… from args.
 * "Line|chart" is "Line" in a context of its own: it shows as "Line".
 * A translation catalog (English text → translation) can be registered with setCatalog;
 * without one every string shows in English. */

let catalog: Record<string, string> | null = null

export function setCatalog(c: Record<string, string> | null): void { catalog = c }

export function t(s: string, ...args: (string | number)[]): string {
  const key = s.includes('|') ? s.split('|')[0] : s
  let out = catalog?.[s] ?? catalog?.[key] ?? key
  if (args.length) out = out.replace(/\{(\d+)\}/g, (m, i: string) => (args[+i] != null ? String(args[+i]) : m))
  return out
}
