import { spawn } from 'node:child_process';
import { promises as fs } from 'node:fs';
import { join, delimiter, dirname, isAbsolute } from 'node:path';
import { pathToFileURL } from 'node:url';
import { createHash } from 'node:crypto';
import { tmpdir } from 'node:os';
import {readResults, type Results} from './results';

export interface Backend { kind: 'excel' | 'libreoffice'; executable: string; version: string }
export interface BackendOptions { backend?: string; libreOfficePath?: string; timeoutMs?: number }
export function runProcess(command: string, args: string[], options: {timeoutMs?: number; signal?: AbortSignal; env?: NodeJS.ProcessEnv} = {}): Promise<string> {
    return new Promise((resolve, reject) => {
        if (options.signal?.aborted) return reject(new Error('Calculation cancelled'));
        const environment={...process.env,...options.env};
        // Native spreadsheet applications must not inherit Electron/extension-host runtime hooks.
        for (const key of Object.keys(environment)) if (/^(VSCODE_|ELECTRON_|NODE_OPTIONS$|NODE_REPL_EXTERNAL_MODULE$)/i.test(key)) delete environment[key];
        const child = spawn(command, args, {windowsHide:true,cwd:isAbsolute(command)?dirname(command):undefined,env:environment,stdio:['ignore','pipe','pipe']});
        let output = '', error = '', aborted = false;
        const stop = () => {
            if (aborted) return;
            aborted = true;
            if (process.platform === 'win32' && child.pid) spawn('taskkill.exe',['/PID',String(child.pid),'/T','/F'],{windowsHide:true,stdio:'ignore'});
            else child.kill('SIGKILL');
        };
        const timer = setTimeout(stop, options.timeoutMs ?? 90000);
        options.signal?.addEventListener('abort', stop, {once:true});
        child.stdout.on('data', b => {if(output.length < 100000) output += b;});
        child.stderr.on('data', b => {if(error.length < 100000) error += b;});
        const cleanup = () => {clearTimeout(timer); options.signal?.removeEventListener('abort', stop);};
        child.on('error', e => {cleanup(); reject(e);});
        child.on('close', code => {
            cleanup();
            if (aborted) reject(new Error(options.signal?.aborted ? 'Calculation cancelled' : 'Calculation timed out'));
            else if (code !== 0) reject(new Error(error.trim() || output.trim() || command + ' exited with ' + code));
            else resolve(output.trim());
        });
    });
}
async function findExecutable(candidates: string[]): Promise<string | undefined> {
    for (const path of candidates) try {if((await fs.stat(path)).isFile()) return path;} catch {}
}
export async function findBackend(options: BackendOptions): Promise<Backend | undefined> {
    if (options.backend === 'none') return undefined;
    if (process.platform === 'win32' && (!options.backend || ['auto','excel'].includes(options.backend))) {
        try {
            const server = await runProcess('reg.exe',['query','HKCR\\Excel.Application\\CLSID'],{timeoutMs:5000});
            const clsid = server.match(/\{[0-9A-F-]+\}/i)?.[0];
            if (clsid) {
                const registration = await runProcess('reg.exe',['query','HKCR\\CLSID\\'+clsid+'\\LocalServer32'],{timeoutMs:5000});
                const match = registration.match(/REG_SZ\s+(.+)/);
                const executable = match?.[1].match(/(?:"([^"]+\.exe)"|(.+?\.exe))/i);
                const path = executable?.[1] ?? executable?.[2];
                const stat = path ? await fs.stat(path) : undefined;
                return {kind:'excel',executable:'powershell.exe',version:registration+':'+(stat?.mtimeMs ?? '')};
            }
        } catch {}
        if (options.backend === 'excel') throw new Error('Excel automation is not installed or registered.');
    }
    if (!options.backend || ['auto','libreoffice'].includes(options.backend)) {
        const candidates = [
            options.libreOfficePath,
            ...process.platform === 'win32' ? ['C:/Program Files/LibreOffice/program/soffice.exe','C:/Program Files/LibreOffice/program/soffice.com'] : ['/usr/bin/libreoffice','/usr/bin/soffice','/Applications/LibreOffice.app/Contents/MacOS/soffice'],
            ...(process.env.PATH ?? '').split(delimiter).map(p => join(p,process.platform === 'win32'?'soffice.exe':'soffice'))
        ].filter(Boolean) as string[];
        const executable = await findExecutable(candidates);
        if (executable) {
            const stat=await fs.stat(executable);
            // soffice.exe --version opens a GUI dialog on Windows; fingerprint the binary instead.
            const version=process.platform==='win32' ? stat.size+':'+stat.mtimeMs : await runProcess(executable,['--version'],{timeoutMs:10000});
            return {kind:'libreoffice',executable,version};
        }
        if (options.backend === 'libreoffice') throw new Error('LibreOffice was not found. Set ahmedSpreadsheet.libreOfficePath.');
    }
    return undefined;
}
export async function recalculate(buffer: Uint8Array, backend: Backend, options: {
    cacheRoot: string; excelScript: string; timeoutMs?: number; signal?: AbortSignal; force?: boolean
}): Promise<Results> {
    if(options.signal?.aborted)throw new Error('Calculation cancelled');
    await fs.mkdir(options.cacheRoot,{recursive:true});
    const key = createHash('sha256').update('results-v1').update(buffer).update(JSON.stringify(backend)).digest('hex');
    const cacheFile = join(options.cacheRoot,key+'.json');
    if (!options.force) try {return JSON.parse(await fs.readFile(cacheFile,'utf8'));} catch {}
    // Keep native profile paths short: LibreOffice still encounters Windows MAX_PATH limits.
    const job = await fs.mkdtemp(join(tmpdir(),'ahmed-xlsx-'));
    const input = join(job,'input.xlsx'), outputDir = join(job,'output'), output = join(outputDir,'input.xlsx');
    const pidFile = join(job,'excel.pid');
    await fs.mkdir(outputDir); await fs.writeFile(input,buffer);
    try {
        if (backend.kind === 'libreoffice') {
            const profile = join(job,'profile'); await fs.mkdir(join(profile,'user'),{recursive:true});
            await fs.writeFile(join(profile,'user','registrymodifications.xcu'), `<?xml version="1.0"?><oor:items xmlns:oor="http://openoffice.org/2001/registry"><item oor:path="/org.openoffice.Office.Common/Security/Scripting"><prop oor:name="MacroSecurityLevel" oor:op="fuse"><value>3</value></prop></item><item oor:path="/org.openoffice.Office.Calc/Content/Update"><prop oor:name="Link" oor:op="fuse"><value>1</value></prop></item></oor:items>`);
            await runProcess(backend.executable,['-env:UserInstallation='+pathToFileURL(profile).href,'--headless','--nologo','--nodefault','--norestore','--convert-to','xlsx:Calc MS Excel 2007 XML','--outdir',outputDir,input],options);
        } else {
            await runProcess(backend.executable,['-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',options.excelScript],{...options,env:{AHMED_XLSX_INPUT:input,AHMED_XLSX_OUTPUT:output,AHMED_XLSX_PID:pidFile}});
        }
        const recalculated = new Uint8Array(await fs.readFile(output));
        const results = await readResults(recalculated);
        const temporary = join(job,'results.json');
        await fs.writeFile(temporary,JSON.stringify(results));
        await fs.rename(temporary,cacheFile).catch(async () => {await fs.copyFile(temporary,cacheFile);});
        return results;
    } finally {
        // Only an automation instance created by our worker is eligible for cleanup.
        if (backend.kind === 'excel') {
            try {
                const pid = (await fs.readFile(pidFile,'utf8')).trim();
                if (/^\d+$/.test(pid)) await runProcess('taskkill.exe',['/PID',pid,'/T','/F'],{timeoutMs:5000}).catch(()=>{});
            } catch {}
        }
        await fs.rm(job,{recursive:true,force:true});
    }
}
