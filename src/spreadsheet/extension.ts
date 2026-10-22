import * as vscode from 'vscode';
import {randomBytes,createHash} from 'node:crypto';
import {promises as fs} from 'node:fs';
import {join,basename,extname} from 'node:path';
import {inspectWorkbook} from './package';
import {findBackend,recalculate} from './native';

const VIEW='sharafeddeen.spreadsheetViewer';
interface Document extends vscode.CustomDocument {
    source:Uint8Array;dirty:boolean;readonly:boolean;generation:number;
    panel?:vscode.WebviewPanel;load?:()=>Promise<void>;
    request?:(type:string,content:any)=>Promise<any>;
    pendingSave?:Uint8Array;lastWrite?:number;
}
const same=(a:Uint8Array,b:Uint8Array)=>createHash('sha256').update(a).digest('hex')===createHash('sha256').update(b).digest('hex');
export function activate(context:vscode.ExtensionContext) {
    const sessions=new Map<string,any>();
    const cacheRoot=join(context.globalStorageUri.fsPath,'calculation-cache');
    const output=vscode.window.createOutputChannel('Ahmed Spreadsheet Viewer');
    const changes=new vscode.EventEmitter<vscode.CustomDocumentContentChangeEvent<Document>>();
    context.subscriptions.push(output,changes);
    const serialize=async(document:Document,ext=extname(document.uri.path).slice(1))=>{
        if(!document.request)throw new Error('Spreadsheet editor is not ready');
        const result=await document.request('exportRequest',{ext});
        return new Uint8Array(result.buffer);
    };
    const write=async(document:Document,target:vscode.Uri,bytes:Uint8Array,overwrite=false)=>{
        if(overwrite) {
            const disk=await vscode.workspace.fs.readFile(document.uri);
            if(!same(disk,document.source))throw new Error('The file changed on disk. Save As or revert before saving.');
        }
        document.lastWrite=Date.now();
        await vscode.workspace.fs.writeFile(target,bytes);
        if(target.toString()===document.uri.toString()) {
            document.source=bytes;document.dirty=false;
            await document.panel?.webview.postMessage({type:'saveDone'});
        }
    };
    const provider:vscode.CustomEditorProvider<Document>={
        onDidChangeCustomDocument:changes.event,
        async openCustomDocument(uri,openContext) {
            const source=await vscode.workspace.fs.readFile(uri);
            const stat=await vscode.workspace.fs.stat(uri);
            const document:Document={uri,source,dirty:!!openContext.backupId,readonly:!!(stat.permissions! & vscode.FilePermission.Readonly)||extname(uri.path).toLowerCase()==='.xlsm',
                generation:0,dispose(){}};
            if(openContext.backupId)document.pendingSave=await vscode.workspace.fs.readFile(vscode.Uri.parse(openContext.backupId));
            return document;
        },
        async saveCustomDocument(document) {
            if(document.readonly)throw new Error('File is read only. Use Save As.');
            const bytes=document.pendingSave ?? await serialize(document);
            document.pendingSave=undefined;
            await write(document,document.uri,bytes,true);
        },
        async saveCustomDocumentAs(document,target) {await write(document,target,await serialize(document),target.toString()===document.uri.toString());},
        async revertCustomDocument(document) {document.dirty=false;document.pendingSave=undefined;await document.load?.();},
        async backupCustomDocument(document,backupContext) {
            await vscode.workspace.fs.writeFile(backupContext.destination,await serialize(document));
            return {id:backupContext.destination.toString(),delete:()=>vscode.workspace.fs.delete(backupContext.destination)};
        },
        async resolveCustomEditor(document,panel) {
            document.panel=panel;
            const key=document.uri.toString(),testMode=context.extensionMode===vscode.ExtensionMode.Test;
            let info:Awaited<ReturnType<typeof inspectWorkbook>>|undefined;
            let autoGeneration=-1,controller:AbortController|undefined,disposed=false;
            let resolveReady:(value:any)=>void;
            const ready=new Promise(resolve=>{resolveReady=resolve;});
            let nextRequest=0;
            const requests=new Map<number,{resolve:(value:any)=>void;reject:(e:Error)=>void;timer:NodeJS.Timeout}>();
            document.request=(type,content)=>new Promise((resolve,reject)=>{
                const id=++nextRequest;
                const timer=setTimeout(()=>{requests.delete(id);reject(new Error('Spreadsheet editor request timed out'));},60000);
                requests.set(id,{resolve,reject,timer});
                void panel.webview.postMessage({type,content:{...content,id}});
            });
            const status=(value:string,warnings:string[]=[])=>{
                const session=sessions.get(key);if(session)session.lastStatus=value;
                return panel.webview.postMessage({type:'status',content:{value,warnings}});
            };
            const calculate=async(force=false,revision=0,edited=false)=>{
                if(disposed)return;
                if(!vscode.workspace.isTrusted){await status('Saved results · trust this workspace to recalculate',info?.warnings);return;}
                const version=document.generation;
                controller?.abort();controller=new AbortController();
                const localController=controller;
                try {
                    const bytes=edited ? await serialize(document,'xlsx') : document.source;
                    const currentInfo=edited ? await inspectWorkbook(bytes) : info;
                    if(!currentInfo?.formulaCount){await status('No formulas to recalculate');return;}
                    const config=vscode.workspace.getConfiguration('ahmedSpreadsheet');
                    const backend=await findBackend({backend:config.get('backend'),libreOfficePath:config.get('libreOfficePath')});
                    if(!backend){await status('Saved results · no calculation backend available',info?.warnings);return;}
                    if(version!==document.generation||disposed||localController.signal.aborted)return;
                    await status('Calculating with '+backend.kind+'…',info?.warnings);
                    const results=await vscode.window.withProgress({location:vscode.ProgressLocation.Notification,title:'Calculating spreadsheet',cancellable:true},async(_,token)=>{
                        const subscription=token.onCancellationRequested(()=>localController.abort());
                        try{return await recalculate(bytes,backend,{cacheRoot,excelScript:join(context.extensionPath,'out','backend','recalculate-excel.ps1'),timeoutMs:config.get<number>('timeoutSeconds',90)*1000,signal:localController.signal,force});}
                        finally{subscription.dispose();}
                    });
                    if(version!==document.generation||disposed)return;
                    await panel.webview.postMessage({type:'calculated',content:{results,backend:backend.kind,revision}});
                }catch(e){
                    if(version!==document.generation||disposed)return;
                    const message=e instanceof Error?e.message:String(e);
                    output.appendLine('Calculation: '+message);
                    await status('Saved results · calculation unavailable: '+message,info?.warnings);
                }
            };
            const load=async()=>{
                const version=++document.generation;controller?.abort();
                try {
                    document.source=await vscode.workspace.fs.readFile(document.uri);
                    const bytes=document.pendingSave ?? document.source;
                    document.pendingSave=undefined;
                    const ext=extname(document.uri.path).toLowerCase();
                    info=['.xlsx','.xlsm'].includes(ext)?await inspectWorkbook(bytes):undefined;
                    if(version!==document.generation||disposed)return;
                    await panel.webview.postMessage({type:'open',content:{buffer:Array.from(bytes),ext,readOnly:document.readonly,documentCacheId:key,info}});
                }catch(e){await panel.webview.postMessage({type:'loadError',content:String(e instanceof Error?e.message:e)});}
            };
            document.load=load;
            const session:any={ready,recalculate:()=>panel.webview.postMessage({type:'recalculateRequest'})};
            if(testMode){session.action=(content:any)=>document.request!('testAction',content);session.save=()=>provider.saveCustomDocument(document,new vscode.CancellationTokenSource().token);session.saveAs=(target:vscode.Uri)=>provider.saveCustomDocumentAs(document,target,new vscode.CancellationTokenSource().token);session.isDirty=()=>document.dirty;}
            sessions.set(key,session);
            panel.webview.options={enableScripts:true,localResourceRoots:[vscode.Uri.joinPath(context.extensionUri,'out','webview')]};
            const listeners:vscode.Disposable[]=[];
            listeners.push(panel.webview.onDidReceiveMessage(async message=>{
                try {
                    if(message.type==='init')await load();
                    else if(message.type==='exportResponse'||message.type==='testResponse'){
                        const p=message.content,pending=requests.get(p.id);
                        if(pending){clearTimeout(pending.timer);requests.delete(p.id);if(p.error)pending.reject(new Error(p.error));else pending.resolve(p);}
                    }else if(message.type==='previewReady'){
                        resolveReady!(message.content);
                        if(autoGeneration!==document.generation&&info?.missingResults&&vscode.workspace.getConfiguration('ahmedSpreadsheet').get('autoCalculateMissing',true)){autoGeneration=document.generation;void calculate();}
                        if(document.dirty)changes.fire({document});
                    }else if(message.type==='recalculate')await calculate(true,message.content?.revision ?? 0,message.content?.edited ?? false);
                    else if(message.type==='cancel')controller?.abort();
                    else if(message.type==='change'){
                        if(!document.readonly){document.dirty=true;document.generation++;controller?.abort();changes.fire({document});}
                    }else if(message.type==='save'){
                        if(document.readonly)throw new Error('File is read only. Use Save As.');
                        document.pendingSave=new Uint8Array(message.content);
                        if(!document.dirty){document.dirty=true;changes.fire({document});}
                        await vscode.commands.executeCommand('workbench.action.files.save');
                    }else if(message.type==='saveAs'){
                        const ext=String(message.content.ext ?? 'xlsx').replace(/[^a-z0-9]/g,'');
                        const name=basename(document.uri.path,extname(document.uri.path))+'.'+ext;
                        const target=await vscode.window.showSaveDialog({defaultUri:vscode.Uri.joinPath(document.uri,'..',name),filters:{Spreadsheet:[ext]}});
                        if(target){
                            if(target.toString()===key){document.pendingSave=new Uint8Array(message.content.content);if(!document.dirty){document.dirty=true;changes.fire({document});}await vscode.commands.executeCommand('workbench.action.files.save');}
                            else {await write(document,target,new Uint8Array(message.content.content));await vscode.commands.executeCommand('vscode.openWith',target,VIEW);}
                        }
                    }else if(message.type==='openDesktop')await vscode.env.openExternal(document.uri);
                    else if(message.type==='openExternal'){
                        const uri=vscode.Uri.parse(String(message.content));
                        if(['http','https','mailto'].includes(uri.scheme))await vscode.env.openExternal(uri);
                    }else if(message.type==='editInVSCode')await vscode.commands.executeCommand('vscode.openWith',document.uri,'default');
                    else if(message.type==='developerTool')await vscode.commands.executeCommand('workbench.action.toggleDevTools');
                    else if(message.type==='renderError'){output.appendLine('Preview: '+String(message.content));resolveReady!({error:String(message.content)});}
                    else if(message.type==='calculationRendered')session.calculation=message.content;
                }catch(e){const message=String(e instanceof Error?e.message:e);output.appendLine(message);await vscode.window.showErrorMessage(message);}
            }));
            if(document.uri.scheme==='file'){
                const watcher=vscode.workspace.createFileSystemWatcher(new vscode.RelativePattern(vscode.Uri.joinPath(document.uri,'..'),basename(document.uri.path)));
                listeners.push(watcher,watcher.onDidChange(async()=>{
                    if(Date.now()-(document.lastWrite ?? 0)<1500)return;
                    if(document.dirty)await status('File changed on disk · Save As to keep your edits');
                    else void load();
                }),watcher.onDidDelete(()=>void panel.webview.postMessage({type:'loadError',content:'Source workbook was removed.'})));
            }
            listeners.push(panel.onDidDispose(()=>{
                disposed=true;controller?.abort();sessions.delete(key);listeners.forEach(l=>l.dispose());
                for(const pending of requests.values()){clearTimeout(pending.timer);pending.reject(new Error('Editor closed'));}requests.clear();
                resolveReady!({error:'Preview closed'});
            }));
            const htmlPath=vscode.Uri.joinPath(context.extensionUri,'out','webview','spreadsheet.html');
            let html=new TextDecoder().decode(await vscode.workspace.fs.readFile(htmlPath));
            const base=panel.webview.asWebviewUri(vscode.Uri.joinPath(context.extensionUri,'out','webview')).toString(),nonce=randomBytes(18).toString('hex');
            const csp=`default-src 'none'; script-src 'nonce-${nonce}' ${panel.webview.cspSource}; style-src ${panel.webview.cspSource} 'unsafe-inline'; font-src ${panel.webview.cspSource} data:; img-src ${panel.webview.cspSource} data: blob:; connect-src ${panel.webview.cspSource};`;
            html=html.replace('{{configs}}',JSON.stringify({language:vscode.env.language,route:'excel',testMode}))
                .replace('<base href="/">','<base href="'+base+'/">').replace('<head>','<head><meta http-equiv="Content-Security-Policy" content="'+csp+'">').replace(/<script\b/g,'<script nonce="'+nonce+'"');
            panel.webview.html=html;
        }
    };
    context.subscriptions.push(vscode.window.registerCustomEditorProvider(VIEW,provider,{webviewOptions:{retainContextWhenHidden:true},supportsMultipleEditorsPerDocument:false}));
    context.subscriptions.push(vscode.commands.registerCommand('ahmedSpreadsheet.clearCache',async()=>{
        await fs.rm(cacheRoot,{recursive:true,force:true});vscode.window.showInformationMessage('Spreadsheet calculation cache cleared.');
    }));
    return {sessions,viewType:VIEW};
}
