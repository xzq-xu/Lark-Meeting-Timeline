import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { basename, dirname, resolve } from 'node:path';

const entry = resolve(new URL('../packages/meeting-timeline-sdk/producer.mjs', import.meta.url).pathname);
const visited = new Set();
const imports = [];

async function visit(file) {
  if (visited.has(file)) return;
  visited.add(file);
  const source = await readFile(file, 'utf8');
  const pattern = /(?:import|export)\s+(?:[^'";]+?\s+from\s+)?['"]([^'"]+)['"]/g;
  for (const match of source.matchAll(pattern)) {
    const specifier = match[1];
    imports.push({ file, specifier });
    assert.equal(specifier.startsWith('node:'), false, `${basename(file)} imports ${specifier}`);
    if (specifier.startsWith('.')) await visit(resolve(dirname(file), specifier));
  }
}

await visit(entry);

const files = [...visited].map((file) => basename(file)).sort();
assert.deepEqual(files, ['errors.mjs', 'producer.mjs']);
assert.equal(imports.some((row) => row.specifier === './index.mjs'), false);

console.log('ok sdk producer browser entry');
