import { cpSync, existsSync, mkdirSync, readdirSync } from 'node:fs';
if (!existsSync('public/index.html')) throw new Error('public/index.html not found');
mkdirSync('dist', {recursive:true});
for (const name of readdirSync('public')) cpSync(`public/${name}`, `dist/${name}`, {recursive:true, force:true});
console.log('Static label studio created at dist/');
