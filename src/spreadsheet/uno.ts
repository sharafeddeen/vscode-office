import {spawn,type ChildProcessWithoutNullStreams} from 'node:child_process';
import {createInterface} from 'node:readline';
import {promises as fs} from 'node:fs';
import {join,dirname,resolve,isAbsolute} from 'node:path';
import {tmpdir} from 'node:os';
import {findBackend,runProcess} from './native';
import {readResults} from './results';
import {inspectWorkbook} from './package';

export interface UnoOptions {officePath?:string;pythonPath?:string;workerPath:string;timeoutMs?:number;signal?:AbortSignal}
export async function discoverUno(options:Partial<UnoOptions>={}):Promise<{office:string;python:string}> {
    const backend=await findBackend({backend:'libreoffice',libreOfficePath:options.officePath});
    if(!backend)throw new Error('LibreOffice is required for native editing.');
    const program=dirname(backend.executable);
    const candidates=[options.pythonPath,join(program,process.platform==='win32'?'python.exe':'python'),
        '/usr/bin/python3','/usr/local/bin/python3',
        '/Applications/LibreOffice.app/Contents/Resources/python'].filter(Boolean) as string[];
    for(const python of [...new Set(candidates)]) {
        try {
            if(!(await fs.stat(python)).isFile())continue;
            await runProcess(python,['-s','-c','import uno; print("UNO_READY")'],{timeoutMs:5000});
            return {office:backend.executable,python};
        }catch{}
    }
    throw new Error('Python UNO was not found. Install the LibreOffice Python/UNO component (python3-uno on Linux), or set ahmedSpreadsheet.pythonPath.');
}
export class UnoDocument {
    readonly root:string;
    private process:ChildProcessWithoutNullStreams;
    private pending=new Map<number,{resolve:(v:any)=>void;reject:(e:Error)=>void;timer:NodeJS.Timeout}>();
    private next=0;
    private exited=false;
    private disposed=false;
    private stopped=false;
    private model:any[]|undefined;
    get alive(){return !this.exited&&!this.disposed&&!this.stopped;}
    private stderr='';
    private sequence:Promise<any>=Promise.resolve();
    private constructor(root:string,process:ChildProcessWithoutNullStreams,private options:UnoOptions){this.root=root;this.process=process;}
    static async open(source:Uint8Array,ext:string,options:UnoOptions):Promise<UnoDocument> {
        const runtime=await discoverUno(options);
        if(options.signal?.aborted)throw new Error('Native document opening cancelled');
        const root=await fs.mkdtemp(join(tmpdir(),'ahmed-uno-'));
        const input=join(root,'input.'+ext.replace(/^\./,''));
        await fs.writeFile(input,source);
        const env={...process.env,PYTHONIOENCODING:'utf-8',PYTHONUNBUFFERED:'1'};
        for(const key of Object.keys(env))if(/^(VSCODE_|ELECTRON_|NODE_OPTIONS$|PYTHONPATH$|PYTHONHOME$)/i.test(key))delete env[key];
        const child=spawn(runtime.python,['-s',options.workerPath,'--office',runtime.office,'--root',root,'--input',input],{
            cwd:dirname(runtime.office),windowsHide:true,env,stdio:'pipe',detached:process.platform!=='win32'
        });
        const document=new UnoDocument(root,child,options);
        try {
            await new Promise<void>((ready,reject)=>{
                const timer=setTimeout(()=>reject(new Error('LibreOffice document opening timed out')),options.timeoutMs ?? 90000);
                const abort=()=>reject(new Error('Native document opening cancelled'));
                options.signal?.addEventListener('abort',abort,{once:true});
                const finish=(error?:Error)=>{clearTimeout(timer);options.signal?.removeEventListener('abort',abort);error?reject(error):ready();};
                const lines=createInterface({input:child.stdout});
                lines.on('line',line=>{
                    try {
                        const message=JSON.parse(line);
                        if(message.ready){finish();return;}
                        if(message.fatal){finish(new Error(message.fatal));document.fail(new Error(message.fatal));return;}
                        const waiting=document.pending.get(message.id);
                        if(waiting){clearTimeout(waiting.timer);document.pending.delete(message.id);message.error?waiting.reject(new Error(message.error)):waiting.resolve(message.result);}
                    }catch(e){document.fail(new Error('Invalid response from LibreOffice worker'));}
                });
                child.stderr.on('data',bytes=>{document.stderr=(document.stderr+bytes.toString()).slice(-8000);});
                child.on('error',error=>{finish(error);document.fail(error);});
                child.on('exit',code=>{
                    document.exited=true;
                    const error=new Error('LibreOffice worker exited ('+code+'). '+document.stderr);
                    finish(error);document.fail(error);
                });
            });
            return document;
        }catch(error){await document.dispose();throw error;}
    }
    private fail(error:Error){for(const request of this.pending.values()){clearTimeout(request.timer);request.reject(error);}this.pending.clear();}
    private send(method:string,data:any={}):Promise<any>{
        if(this.exited||this.disposed)return Promise.reject(new Error('LibreOffice document session is closed. Reopen the workbook to reconnect.'));
        const id=++this.next;
        return new Promise((resolve,reject)=>{
            const timer=setTimeout(()=>{this.pending.delete(id);reject(new Error('LibreOffice request timed out; the editor has kept its unsaved data.'));void this.stop();},this.options.timeoutMs ?? 90000);
            this.pending.set(id,{resolve,reject,timer});
            this.process.stdin.write(JSON.stringify({id,method,...data})+'\n',error=>{if(error){clearTimeout(timer);this.pending.delete(id);reject(error);}});
        });
    }
    request(method:string,data:any={}):Promise<any>{
        const task=this.sequence.then(()=>this.send(method,data));
        this.sequence=task.catch(()=>{});
        return task;
    }
    async bind(model:any[]){const result=await this.request('bind',{model});this.model=structuredClone(model);return result;}
    async sync(model:any[]){const result=await this.request('sync',{model});this.model=structuredClone(model);return result;}
    operation(command:any){this.model=undefined;return this.request('operation',{command});}
    async snapshot(){const response=await this.request('calculate');const bytes=await this.readOwned(response.path);return {bytes,results:await readResults(bytes),info:await inspectWorkbook(bytes),history:await this.request('inspect')};}
    async export(ext='xlsx'){const response=await this.request('export',{ext});return this.readOwned(response.path);}
    private readOwned(path:string) {
        const target=resolve(path);
        if(!isAbsolute(path)||!target.startsWith(resolve(this.root)+require('node:path').sep))throw new Error('Worker returned a path outside its temporary directory');
        return fs.readFile(target).then(bytes=>new Uint8Array(bytes));
    }
    async recovery(){return {bytes:await this.readOwned(join(this.root,'recovery.ods')),model:this.model};}
    async cancel(){this.stopped=true;await this.stop();}
    private async stop(){
        this.stopped=true;
        if(this.process.pid&&!this.exited) {
            if(process.platform==='win32')await runProcess('taskkill.exe',['/PID',String(this.process.pid),'/T','/F'],{timeoutMs:5000}).catch(()=>{});
            else {try{process.kill(-this.process.pid,'SIGKILL');}catch{}}
        }
        this.fail(new Error('Native document stopped'));
    }
    async dispose(){
        if(this.disposed)return;
        if(!this.exited&&!this.stopped) {
            let timer:NodeJS.Timeout;
            try{await Promise.race([this.send('close'),new Promise((_,reject)=>{timer=setTimeout(()=>reject(new Error('Closing native document timed out')),5000);})]);}catch{}
            finally{clearTimeout(timer!);}
            // The close response precedes the worker's cleanup/exit.
            await new Promise<void>(resolve=>{if(this.exited)return resolve();const timeout=setTimeout(resolve,6000);this.process.once('exit',()=>{clearTimeout(timeout);resolve();});});
        }
        this.disposed=true;await this.stop();
        // root was created by mkdtemp above, never derived from a source/workspace path.
        await fs.rm(this.root,{recursive:true,force:true}).catch(()=>{});
    }
}
