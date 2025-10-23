#!/usr/bin/env node
/*
  Strip sourceMappingURL comments from docx-preview dist files to silence
  noisy source-map-loader warnings in CRA dev/prod.
*/
const fs = require('fs');
const path = require('path');

const root = process.cwd();
const targets = [
  'node_modules/docx-preview/dist/docx-preview.mjs',
  'node_modules/docx-preview/dist/docx-preview.min.mjs',
  'node_modules/docx-preview/dist/docx-preview.js',
  'node_modules/docx-preview/dist/docx-preview.min.js',
];

let changed = 0;
for (const rel of targets) {
  const p = path.join(root, rel);
  if (!fs.existsSync(p)) continue;
  try {
    const src = fs.readFileSync(p, 'utf8');
    // Remove any line that contains sourceMappingURL (handles // and /* */ forms)
    const out = src
      .replace(/\n\/\/\s*#?\s*sourceMappingURL=.*$/gm, '')
      .replace(/\/*#\s*sourceMappingURL=.*?\*\//gs, '');
    if (out !== src) {
      fs.writeFileSync(p, out, 'utf8');
      changed++;
      console.log(`[patch] stripped sourceMappingURL in ${rel}`);
    }
  } catch (e) {
    console.warn(`[patch] failed to process ${rel}:`, e && e.message);
  }
}

if (changed === 0) {
  console.log('[patch] no docx-preview files modified (already clean or not found).');
} else {
  console.log(`[patch] docx-preview sourcemap comments removed in ${changed} file(s).`);
}
