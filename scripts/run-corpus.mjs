import {build} from 'esbuild';
import {spawnSync} from 'node:child_process';
await build({entryPoints:['scripts/verify-corpus.ts'],outfile:'.cache/verify-corpus.cjs',bundle:true,platform:'node',format:'cjs',logLevel:'warning'});
const result=spawnSync(process.execPath,['.cache/verify-corpus.cjs',...process.argv.slice(2)],{stdio:'inherit',env:process.env});
process.exitCode=result.status??1;
