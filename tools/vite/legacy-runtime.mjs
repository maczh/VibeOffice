import { cp, readFile, stat, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
const publicRoot = path.join(repoRoot, 'public');
const MIME = new Map([
  ['.html', 'text/html; charset=utf-8'], ['.js', 'text/javascript; charset=utf-8'], ['.css', 'text/css; charset=utf-8'],
  ['.json', 'application/json; charset=utf-8'], ['.svg', 'image/svg+xml'], ['.webmanifest', 'application/manifest+json'],
  ['.woff2', 'font/woff2'], ['.woff', 'font/woff'], ['.ttf', 'font/ttf'], ['.otf', 'font/otf'], ['.png', 'image/png'],
]);

function reactManifest(source) {
  const manifest = JSON.parse(source);
  const products = [
    { legacy: 'quire', id: 'wordx', name: 'Wordx' },
    { legacy: 'ledger', id: 'excelx', name: 'Excelx' },
    { legacy: 'lectern', id: 'pptx', name: 'PPTX' },
  ];
  for (const product of products) {
    const shortcut = manifest.shortcuts?.find((entry) => entry.url.startsWith(product.legacy + '/'));
    if (shortcut) {
      shortcut.url = shortcut.url.replace(product.legacy + '/', product.id + '/');
      shortcut.description = shortcut.description?.replace(new RegExp(product.legacy, 'i'), product.name);
    }
    const screenshot = manifest.screenshots?.find((entry) => entry.src.startsWith(product.legacy + '/'));
    if (screenshot) {
      screenshot.src = screenshot.src.replace(product.legacy + '/', product.id + '/');
      screenshot.label = screenshot.label?.replace(new RegExp(product.legacy, 'i'), product.name);
    }
  }
  return JSON.stringify(manifest, null, 2);
}

/** Rebrand generated template metadata and preview URLs for the React package routes. */
export function reactTemplates(source) {
  const templates = JSON.parse(source);
  const products = [
    { legacy: 'quire', id: 'wordx' },
    { legacy: 'ledger', id: 'excelx' },
    { legacy: 'lectern', id: 'pptx' },
  ];
  const output = {};
  for (const product of products) {
    const entries = templates[product.legacy];
    if (!Array.isArray(entries)) continue;
    output[product.id] = entries.map((entry) => ({
      ...entry,
      ...(typeof entry.img === 'string' ? { img: entry.img.replace(new RegExp(`^${product.legacy}/`), `${product.id}/`) } : {}),
    }));
  }
  return JSON.stringify(output, null, 2);
}

export function legacyRuntime(app) {
  const apps = Array.isArray(app) ? app : [app];
  const legacyFolders = { wordx: 'quire', excelx: 'ledger', pptx: 'lectern' };
  if (!apps.length || apps.some((name) => !Object.hasOwn(legacyFolders, name))) throw new TypeError('Unknown editor runtime');
  const roots = new Map([...apps.map((name) => [`/${name}`, path.join(publicRoot, legacyFolders[name])]), ['/common', path.join(publicRoot, 'common')],
    ['/fonts', path.join(publicRoot, 'fonts')], ['/icons', path.join(publicRoot, 'icons')]]);
  roots.set('/manifest.webmanifest', path.join(publicRoot, 'manifest.webmanifest'));
  roots.set('/sw.js', path.join(publicRoot, 'sw.js'));

  return {
    name: `vibeoffice-${apps.join('-')}-legacy-runtime`,
    configureServer(server) {
      server.middlewares.use((request, response, next) => {
        let pathname;
        try { pathname = decodeURIComponent(new URL(request.url || '/', 'http://localhost').pathname); }
        catch { return next(); }
        const route = [...roots].find(([prefix]) => prefix.startsWith('/') && (pathname === prefix || pathname.startsWith(prefix + '/')));
        if (!route) return next();
        const [prefix, root] = route;
        const file = prefix === '/manifest.webmanifest' || prefix === '/sw.js' ? root : path.resolve(root, '.' + pathname.slice(prefix.length));
        if (file !== root && !file.startsWith(root + path.sep)) return next();
        stat(file).then((info) => info.isDirectory() ? path.join(file, 'index.html') : file)
          .then((target) => readFile(target).then((contents) => ({ target, contents })))
          .then(({ target, contents }) => {
            response.statusCode = 200;
            response.setHeader('Content-Type', MIME.get(path.extname(target).toLowerCase()) || 'application/octet-stream');
            response.setHeader('Cache-Control', 'no-store');
            const text = contents.toString('utf8');
            response.end(prefix === '/manifest.webmanifest' ? reactManifest(text) : pathname === '/common/templates.json' ? reactTemplates(text) : contents);
          }).catch(next);
      });
    },
    async writeBundle(options) {
      if (!options.dir) return;
      const output = path.resolve(options.dir);
      for (const folder of [...apps.map((name) => [legacyFolders[name], name]), ['common', 'common'], ['fonts', 'fonts'], ['icons', 'icons']]) {
        await cp(path.join(publicRoot, folder[0]), path.join(output, folder[1]), { recursive: true });
      }
      await writeFile(path.join(output, 'common', 'templates.json'), reactTemplates(await readFile(path.join(publicRoot, 'common', 'templates.json'), 'utf8')));
      const manifest = reactManifest(await readFile(path.join(publicRoot, 'manifest.webmanifest'), 'utf8'));
      await writeFile(path.join(output, 'manifest.webmanifest'), manifest);
      await cp(path.join(publicRoot, 'sw.js'), path.join(output, 'sw.js'));
    },
  };
}
