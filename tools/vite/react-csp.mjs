export function reactCsp(origins = process.env.VITE_REMOTE_ORIGINS || '') {
  const allowed = origins.split(',').map((value) => value.trim()).filter(Boolean).map((value) => {
    const url = new URL(value);
    if (url.origin !== value.replace(/\/$/, '') || !['https:', ...(url.hostname === '127.0.0.1' || url.hostname === 'localhost' ? ['http:'] : [])].includes(url.protocol)) {
      throw new TypeError(`Invalid remote origin in VITE_REMOTE_ORIGINS: ${value}`);
    }
    return url.origin;
  });
  const policy = `default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data: blob:; font-src 'self' data:; connect-src 'self' ${allowed.join(' ')}; worker-src 'self'; manifest-src 'self'; frame-src 'self'; object-src 'none'; base-uri 'none'; form-action 'none'`;
  return {
    name: 'vibeoffice-react-csp',
    transformIndexHtml(html) {
      return html.replace('</head>', `<meta name="vibeoffice-remote-origins" content="${allowed.join(' ')}"><meta http-equiv="Content-Security-Policy" content="${policy}"></head>`);
    },
  };
}
