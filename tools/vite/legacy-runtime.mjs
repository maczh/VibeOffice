import { cp, readFile, stat } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
const publicRoot = path.join(repoRoot, 'public');
const MIME = new Map([
  ['.html', 'text/html; charset=utf-8'], ['.js', 'text/javascript; charset=utf-8'], ['.css', 'text/css; charset=utf-8'],
  ['.json', 'application/json; charset=utf-8'], ['.svg', 'image/svg+xml'], ['.webmanifest', 'application/manifest+json'],
  ['.woff2', 'font/woff2'], ['.woff', 'font/woff'], ['.ttf', 'font/ttf'], ['.otf', 'font/otf'], ['.png', 'image/png'],
]);

export function legacyRuntime(app) {
  const apps = Array.isArray(app) ? app : [app];
  if (!apps.length || apps.some((name) => !['quire', 'ledger', 'lectern'].includes(name))) throw new TypeError('Unknown legacy editor runtime');
  const roots = new Map([...apps.map((name) => [`/${name}`, path.join(publicRoot, name)]), ['/common', path.join(publicRoot, 'common')],
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
            response.end(contents);
          }).catch(next);
      });
    },
    async writeBundle(options) {
      if (!options.dir) return;
      const output = path.resolve(options.dir);
      for (const folder of [...apps, 'common', 'fonts', 'icons']) {
        await cp(path.join(publicRoot, folder), path.join(output, folder), { recursive: true });
      }
      await cp(path.join(publicRoot, 'manifest.webmanifest'), path.join(output, 'manifest.webmanifest'));
      await cp(path.join(publicRoot, 'sw.js'), path.join(output, 'sw.js'));
    },
  };
}
