import fs from 'node:fs';
import process from 'node:process';

const scriptPath = new URL('../chaoxing-quiz-tool.user.js', import.meta.url);
const packagePath = new URL('../package.json', import.meta.url);
const source = fs.readFileSync(scriptPath, 'utf8');
const packageJson = JSON.parse(fs.readFileSync(packagePath, 'utf8'));

const blockMatch = source.match(/^\/\/ ==UserScript==\n([\s\S]*?)^\/\/ ==\/UserScript==/m);
if (!blockMatch) {
    console.error('Missing userscript metadata block.');
    process.exit(1);
}

const metadata = new Map();
for (const line of blockMatch[1].split('\n')) {
    const match = line.match(/^\/\/\s+@(\S+)\s+(.*)$/);
    if (!match) continue;
    const [, key, value] = match;
    const current = metadata.get(key) ?? [];
    current.push(value.trim());
    metadata.set(key, current);
}

const required = [
    'name',
    'namespace',
    'version',
    'description',
    'author',
    'homepageURL',
    'supportURL',
    'updateURL',
    'downloadURL',
    'match',
    'grant',
    'run-at'
];

const missing = required.filter((key) => !metadata.has(key));
if (missing.length > 0) {
    console.error(`Missing metadata keys: ${missing.join(', ')}`);
    process.exit(1);
}

const userscriptVersion = metadata.get('version')[0];
if (userscriptVersion !== packageJson.version) {
    console.error(`Version mismatch: userscript=${userscriptVersion}, package=${packageJson.version}`);
    process.exit(1);
}

const matches = metadata.get('match') ?? [];
if (matches.some((value) => value.includes('*://*/*exam*') || value.includes('*://*/*test*'))) {
    console.error('Overly broad exam/test match rule detected.');
    process.exit(1);
}

const connects = metadata.get('connect') ?? [];
if (connects.includes('*')) {
    console.error('Overly broad @connect * detected.');
    process.exit(1);
}

console.log(`Userscript metadata OK (${userscriptVersion}).`);
