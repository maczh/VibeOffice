import { mkdir, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const repo = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
const target = process.argv[2];
if (!['wordx', 'excelx', 'pptx'].includes(target)) throw new Error('Expected wordx, excelx, or pptx');
const output = path.join(repo, target, 'dist');
await mkdir(output, { recursive: true });
const shared = await readFile(path.join(repo, 'commonx', 'dist', 'index.css'), 'utf8');
let app = '';
try { app = await readFile(path.join(output, 'index.css'), 'utf8'); }
catch (error) { if (error.code !== 'ENOENT') throw error; }
await writeFile(path.join(output, 'style.css'), [shared, app].filter(Boolean).join('\n'));
