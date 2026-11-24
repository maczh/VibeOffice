import test from 'node:test';
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import net from 'node:net';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { openPage } from '../../tools/ooxml/cdp.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');

async function unusedPort() {
  const probe = net.createServer();
  await new Promise((resolve) => probe.listen(0, '127.0.0.1', resolve));
  const port = probe.address().port;
  await new Promise((resolve) => probe.close(resolve));
  return port;
}

test('Wordx package OMML service converts Office Math and reports explicit run size', { timeout: 45000 }, async () => {
  const previousBase = process.env.BASE;
  const port = await unusedPort();
  const base = `http://127.0.0.1:${port}`;
  const server = spawn('python3', [path.join(root, 'tools/serve.py'), String(port), path.join(root, 'suite/dist')], { stdio: 'ignore' });
  process.env.BASE = base;
  let page;
  try {
    for (let attempt = 0; attempt < 40 && !await fetch(base).then((response) => response.ok).catch(() => false); attempt++) await new Promise((resolve) => setTimeout(resolve, 100));
    page = await openPage(undefined, { url: `${base}/?new=wordx`, ready: "typeof window.L?.omml?.toMathML === 'function'", timeout: 15000, bypassServiceWorker: true });
    const result = await page.evaluate(`(() => {
      const prefix='http://schemas.openxmlformats.org/officeDocument/2006/math';
      const convert=(body, para=false)=>L.omml.toMathML(para
        ? '<m:oMathPara xmlns:m="'+prefix+'"><m:oMath>'+body+'</m:oMath></m:oMathPara>'
        : '<m:oMath xmlns:m="'+prefix+'">'+body+'</m:oMath>');
      const fraction=convert('<m:f><m:num><m:r><m:t>1</m:t></m:r></m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f>');
      const matrix=convert('<m:m><m:mr><m:e><m:r><m:t>a</m:t></m:r></m:e><m:e><m:r><m:t>b</m:t></m:r></m:e></m:mr></m:m>',true);
      const parser=new DOMParser().parseFromString('<m:oMath xmlns:m="'+prefix+'"><m:r><w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:sz w:val="32"/></w:rPr><m:t>x</m:t></m:r></m:oMath>','application/xml');
      const bad=L.omml.toMathML('<m:oMath><m:r>');
      return {
        fraction:{tag:fraction?.math.firstElementChild?.localName,text:fraction?.math.textContent},
        matrix:{display:matrix?.display,jc:matrix?.jc,table:matrix?.math.getElementsByTagNameNS('http://www.w3.org/1998/Math/MathML','mtable').item(0)?.localName,cells:matrix?.math.getElementsByTagNameNS('http://www.w3.org/1998/Math/MathML','mtd').length},
        fontSize:L.omml.fontSize(parser.documentElement),
        documentInput:L.omml.toMathML(parser)?.math.textContent,
        malformedRejected:bad===null,
        implicitNamespace:L.omml.toMathML('<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>')?.math.textContent,
      };
    })()`);
    assert.deepEqual(result, {
      fraction: { tag: 'mfrac', text: '12' },
      matrix: { display: true, jc: 'centerGroup', table: 'mtable', cells: 2 },
      fontSize: 16,
      documentInput: 'x',
      malformedRejected: true,
      implicitNamespace: 'x',
    });
    assert.deepEqual(page.errors, []);
    assert.deepEqual(page.requests, []);
  } finally {
    if (page) await page.close();
    server.kill();
    if (previousBase === undefined) delete process.env.BASE; else process.env.BASE = previousBase;
  }
});
