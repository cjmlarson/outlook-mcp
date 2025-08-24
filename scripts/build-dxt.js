#!/usr/bin/env node

import { createWriteStream } from 'fs';
import { mkdir, readFile } from 'fs/promises';
import { dirname, join } from 'path';
import { fileURLToPath } from 'url';
import archiver from 'archiver';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const rootDir = join(__dirname, '..');

async function buildDxt() {
  const packageJsonContent = await readFile(join(rootDir, 'package.json'), 'utf-8');
  const packageJson = JSON.parse(packageJsonContent);
  const version = packageJson.version;
  const outputFile = join(rootDir, `outlook-mcp-${version}.dxt`);

  console.log(`Building Desktop Extension: ${outputFile}`);

  // Ensure output directory exists
  await mkdir(dirname(outputFile), { recursive: true });

  // Create archive
  const output = createWriteStream(outputFile);
  const archive = archiver('zip', {
    zlib: { level: 9 } // Maximum compression
  });

  output.on('close', () => {
    console.log(`✅ Desktop Extension created: ${outputFile}`);
    console.log(`   Size: ${(archive.pointer() / 1024).toFixed(2)} KB`);
  });

  archive.on('error', (err) => {
    throw err;
  });

  archive.pipe(output);

  // Add desktop-extension folder contents
  archive.directory(join(rootDir, 'desktop-extension'), false);

  await archive.finalize();
}

buildDxt().catch(console.error);