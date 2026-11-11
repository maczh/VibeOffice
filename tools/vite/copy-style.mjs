import { copyFile, mkdir } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const repo = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
const target = process.argv[2];
if (!['wordx', 'excelx', 'pptx'].includes(target)) throw new Error('Expected wordx, excelx, or pptx');
const output = path.join(repo, target, 'dist');
await mkdir(output, { recursive: true });
await copyFile(path.join(repo, 'commonx', 'dist', 'index.css'), path.join(output, 'style.css'));
