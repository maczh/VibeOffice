import { createWebDavProvider } from '@vibeoffice/core';

const multistatus = `<?xml version="1.0" encoding="utf-8"?>
<d:multistatus xmlns:d="DAV:">
  <d:response><d:href>https://dav.example/docs/2026/</d:href><d:propstat><d:prop><d:displayname>2026</d:displayname><d:resourcetype><d:collection/></d:resourcetype></d:prop></d:propstat></d:response>
  <d:response><d:href>reports/annual%20summary.xlsx</d:href><d:propstat><d:prop><d:displayname>Annual Summary</d:displayname><d:getetag>"rev-1"</d:getetag><d:getcontenttype>application/vnd.openxmlformats-officedocument.spreadsheetml.sheet</d:getcontenttype></d:prop></d:propstat></d:response>
  <d:response><d:href>/docs/2026/notes.docx</d:href><d:propstat><d:prop><d:displayname>Notes</d:displayname></d:prop></d:propstat></d:response>
</d:multistatus>`;

const provider = createWebDavProvider({
  baseUrl: 'https://dav.example/docs/',
  allowedOrigins: ['https://dav.example'],
  fetcher: async () => new Response(multistatus, { status: 207 }),
});

try {
  const files = await provider.list('/docs/2026/');
  document.querySelector('#result')!.textContent = JSON.stringify(files);
} catch (error) {
  document.querySelector('#result')!.textContent = JSON.stringify({ error: error instanceof Error ? error.message : String(error) });
}
