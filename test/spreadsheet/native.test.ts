import {test} from 'node:test';
import assert from 'node:assert/strict';
import {promises as fs} from 'node:fs';
import {join} from 'node:path';
import {tmpdir} from 'node:os';
import {createHash} from 'node:crypto';
import {fixture} from './fixture';
import {findBackend,recalculate} from '../../src/spreadsheet/native';

test('native backend calculates missing ABS/IF and cross-sheet results; source is unchanged and jobs are cleaned',{skip:!process.env.RUN_NATIVE},async()=>{
    const backend=await findBackend({backend:process.env.NATIVE_BACKEND ?? 'auto'});
    assert.ok(backend,'Requested native backend must be installed');
    const testRoot=await fs.mkdtemp(join(tmpdir(),'ahmed-spreadsheet-test-'));
    const cacheRoot=join(testRoot,'long-profile-'.repeat(8),'nested-profile-'.repeat(6));
    try {
        const bytes=await fixture({missing:true,unusualComments:true});
        const hash=createHash('sha256').update(bytes).digest('hex');
        const results=await recalculate(bytes,backend,{cacheRoot,excelScript:join(process.cwd(),'src/spreadsheet/recalculate-excel.ps1'),timeoutMs:90000});
        assert.equal(results.QUICK_SPLITS.H5.value,'OK');
        assert.equal(results.QUICK_SPLITS.A13.value,2);
        assert.equal(results.QUICK_SPLITS.A14.value,1);
        assert.equal(createHash('sha256').update(bytes).digest('hex'),hash);
        assert.equal((await fs.readdir(cacheRoot)).filter(n=>n.startsWith('job-')).length,0);
        const cached=await recalculate(bytes,{...backend,executable:'not-used-when-cached'},{cacheRoot,excelScript:'unused',timeoutMs:1000}).catch(()=>null);
        // Cache identity includes the executable; the invalid backend must never share another backend's entry.
        assert.equal(cached,null);
    } finally {await fs.rm(testRoot,{recursive:true,force:true});}
});
