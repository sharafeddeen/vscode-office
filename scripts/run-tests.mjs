import {build} from 'esbuild';
import {readdir,mkdir} from 'node:fs/promises';
import {spawnSync} from 'node:child_process';
await mkdir('.cache',{recursive:true});
const files=(await readdir('test/spreadsheet')).filter(f=>f.endsWith('.test.ts'));
await build({entryPoints:files.map(f=>'test/spreadsheet/'+f),outdir:'.cache/tests',bundle:true,platform:'node',format:'cjs',outExtension:{'.js':'.cjs'},logLevel:'warning'});
const result=spawnSync(process.execPath,['--test',...files.map(f=>'.cache/tests/'+f.replace(/\.ts$/,'.cjs'))],{stdio:'inherit',env:process.env});
process.exitCode=result.status??1;
