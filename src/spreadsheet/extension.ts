import * as vscode from 'vscode';
import {randomBytes,createHash} from 'node:crypto';
import {promises as fs} from 'node:fs';
import {join,basename,extname} from 'node:path';
import {inspectWorkbook} from './package';
import {findBackend,recalculate} from './native';
import {UnoDocument} from './uno';

const VIEW='sharafeddeen.spreadsheetViewer';
const WORKBOOKS=new Set(['xlsx','xlsm','xls','ods']);
interface Document extends vscode.CustomDocument {
    source:Uint8Array;dirty:boolean;readonly:boolean;generation:number;revision:number;
    native?:UnoDocument;queue:Promise<any>;model?:any[];nativeSource?:Uint8Array;nativeAllowed?:boolean;nativeOpening?:AbortController;closed?:boolean;
    panel?:vscode.WebviewPanel;load?:()=>Promise<void>;
    request?:(type:string,content:any)=>Promise<any>;lastWrite?:number;restored?:Uint8Array;
}
const same=(a:Uint8Array,b:Uint8Array)=>createHash('sha256').update(a).digest('hex')===createHash('sha256').update(b).digest('hex');
export function activate(context:vscode.ExtensionContext) {
    const sessions=new Map<string,any>(),documents=new Set<Document>();
    const cacheRoot=join(context.globalStorageUri.fsPath,'calculation-cache');
    const output=vscode.window.createOutputChannel('Ahmed Spreadsheet Viewer');
    const changes=new vscode.EventEmitter<vscode.CustomDocumentContentChangeEvent<Document>>();
    context.subscriptions.push(output,changes,{dispose(){for(const document of documents){document.closed=true;document.nativeOpening?.abort();void document.native?.dispose();}}});
    const enqueue=<T>(document:Document,action:()=>Promise<T>):Promise<T>=>{
        const task=document.queue.then(action);document.queue=task.catch(()=>{});return task;
    };
    const nativeOptions=()=>{
        const config=vscode.workspace.getConfiguration('ahmedSpreadsheet');
        return {officePath:config.get<string>('libreOfficePath'),pythonPath:config.get<string>('pythonPath'),
            workerPath:join(context.extensionPath,'out','backend','uno-worker.py'),timeoutMs:config.get<number>('timeoutSeconds',90)*1000};
    };
    const ensureNative=async(document:Document)=>{
        if(!document.native){
            if(!document.nativeAllowed||!vscode.workspace.isTrusted)throw new Error('Native workbook editing requires a trusted workspace and the LibreOffice backend.');
            if(!document.model)throw new Error('Spreadsheet preview is not ready.');
            await document.panel?.webview.postMessage({type:'status',content:{value:'Starting LibreOffice for editing/calculation…'}});
            const opening=new AbortController();document.nativeOpening=opening;
            let native:UnoDocument|undefined;
            try{
                native=await UnoDocument.open(document.nativeSource ?? document.source,extname(document.uri.path).slice(1),{...nativeOptions(),signal:opening.signal});
                await native.bind(document.model);
                if(document.closed||opening.signal.aborted)throw new Error('Native document opening cancelled');
                document.native=native;
            }catch(e){await native?.dispose();throw e;}
            finally{if(document.nativeOpening===opening)document.nativeOpening=undefined;}
        }
        if(!document.native.alive){
            const recovery=await document.native.recovery();
            const replacement=await UnoDocument.open(recovery.bytes,'ods',nativeOptions());
            try{if(recovery.model)await replacement.bind(recovery.model);}
            catch(e){await replacement.dispose();throw e;}
            await document.native.dispose();document.native=replacement;
        }
        return document.native;
    };
    const syncModel=async(document:Document,model:any[])=>{
        const native=await ensureNative(document);
        const result=await native.sync(model);document.model=model;
        return result;
    };
    const currentModel=async(document:Document)=>{
        if(!document.request)throw new Error('Spreadsheet editor is not ready');
        const result=await document.request('modelRequest',{});
        document.revision=result.revision;return result;
    };
    const serialize=async(document:Document,ext=extname(document.uri.path).slice(1))=>{
        if(WORKBOOKS.has(extname(document.uri.path).slice(1).toLowerCase())){
            if(!document.dirty&&!document.native&&ext===extname(document.uri.path).slice(1).toLowerCase())return document.nativeSource ?? document.source;
            if(document.nativeAllowed||document.native){
                const current=await currentModel(document);
                await syncModel(document,current.model);
                return (await ensureNative(document)).export(ext);
            }
            if(!document.dirty&&ext===extname(document.uri.path).slice(1).toLowerCase())return document.source;
            throw new Error('This workbook needs the LibreOffice UNO backend to save without rebuilding it from the grid.');
        }
        const result=await document.request!('exportRequest',{ext});
        return new Uint8Array(result.buffer);
    };
    const write=async(document:Document,target:vscode.Uri,bytes:Uint8Array,overwrite=false)=>{
        if(overwrite&&!same(await vscode.workspace.fs.readFile(document.uri),document.source))
            throw new Error('The file changed on disk. Save As or revert before saving.');
        document.lastWrite=Date.now();
        await vscode.workspace.fs.writeFile(target,bytes);
        if(target.toString()===document.uri.toString()) {
            document.source=bytes;document.dirty=false;
            await document.panel?.webview.postMessage({type:'saveDone'});
        }
    };
    const markChanged=(document:Document)=>{
        document.dirty=true;document.generation++;changes.fire({document});
    };
    const provider:vscode.CustomEditorProvider<Document>={
        onDidChangeCustomDocument:changes.event,
        async openCustomDocument(uri,openContext) {
            const source=await vscode.workspace.fs.readFile(uri),stat=await vscode.workspace.fs.stat(uri);
            const document:Document={uri,source,dirty:!!openContext.backupId,
                readonly:!!(stat.permissions! & vscode.FilePermission.Readonly)||extname(uri.path).toLowerCase()==='.xlsm',
                generation:0,revision:0,queue:Promise.resolve(),dispose(){document.closed=true;document.nativeOpening?.abort();documents.delete(document);void document.native?.dispose();}};
            if(openContext.backupId)document.restored=await vscode.workspace.fs.readFile(vscode.Uri.parse(openContext.backupId));
            documents.add(document);return document;
        },
        saveCustomDocument(document) {return enqueue(document,async()=>{
            if(document.readonly)throw new Error('File is read only. Use Save As.');
            const generation=document.generation;
            const bytes=await serialize(document);
            await write(document,document.uri,bytes,true);
            if(generation!==document.generation){document.dirty=true;changes.fire({document});}
        });},
        saveCustomDocumentAs(document,target) {return enqueue(document,async()=>{
            const ext=extname(target.path).slice(1).toLowerCase();
            await write(document,target,await serialize(document,ext),target.toString()===document.uri.toString());
        });},
        revertCustomDocument(document) {return enqueue(document,async()=>{document.dirty=false;document.restored=undefined;await document.load?.();});},
        backupCustomDocument(document,backupContext) {return enqueue(document,async()=>{
            await vscode.workspace.fs.writeFile(backupContext.destination,await serialize(document));
            return {id:backupContext.destination.toString(),delete:()=>vscode.workspace.fs.delete(backupContext.destination)};
        });},
        async resolveCustomEditor(document,panel) {
            document.panel=panel;
            const key=document.uri.toString(),testMode=context.extensionMode===vscode.ExtensionMode.Test;
            let info:Awaited<ReturnType<typeof inspectWorkbook>>|undefined;
            let loadWarning='',autoGeneration=-1,controller:AbortController|undefined,disposed=false;
            let resolveReady:(value:any)=>void;
            const ready=new Promise(resolve=>{resolveReady=resolve;});
            let nextRequest=0;
            const requests=new Map<number,{resolve:(value:any)=>void;reject:(e:Error)=>void;timer:NodeJS.Timeout}>();
            document.request=(type,content)=>new Promise((resolve,reject)=>{
                const id=++nextRequest;
                const timer=setTimeout(()=>{requests.delete(id);reject(new Error('Spreadsheet editor request timed out'));},60000);
                requests.set(id,{resolve,reject,timer});void panel.webview.postMessage({type,content:{...content,id}});
            });
            const status=(value:string)=>{
                const session=sessions.get(key);if(session)session.lastStatus=value;
                return panel.webview.postMessage({type:'status',content:{value}});
            };
            const nativeResults=async(revision:number)=>{
                const result=await (await ensureNative(document)).snapshot();
                await panel.webview.postMessage({type:'calculated',content:{results:result.results,buffer:Array.from(result.bytes),arrays:result.info.arrays,backend:'LibreOffice UNO',revision,native:true,history:result.history}});
            };
            const calculate=async(revision=0)=>{
                if(disposed)return;
                if(!vscode.workspace.isTrusted){await status('Saved results · trust this workspace to recalculate');return;}
                if(document.nativeAllowed||document.native){
                    await status('Calculating with LibreOffice UNO…');
                    const model=await currentModel(document);
                    await syncModel(document,model.model);
                    await nativeResults(model.revision);
                    return;
                }
                // Legacy calculation remains available for read-only viewing with Excel.
                if(!info?.formulaCount)return;
                const version=document.generation;controller?.abort();controller=new AbortController();
                const config=vscode.workspace.getConfiguration('ahmedSpreadsheet'),backend=await findBackend({backend:config.get('backend'),libreOfficePath:config.get('libreOfficePath')});
                if(!backend){await status('Saved results · no calculation backend available');return;}
                await status('Calculating with '+backend.kind+'…');
                const results=await recalculate(document.source,backend,{cacheRoot,excelScript:join(context.extensionPath,'out','backend','recalculate-excel.ps1'),timeoutMs:config.get<number>('timeoutSeconds',90)*1000,signal:controller.signal,force:true});
                if(version===document.generation&&!disposed)await panel.webview.postMessage({type:'calculated',content:{results,backend:backend.kind,revision}});
            };
            const load=async()=>{
                const version=++document.generation;controller?.abort();document.revision=0;
                await document.native?.dispose();document.native=undefined;document.model=undefined;loadWarning='';
                try {
                    document.source=await vscode.workspace.fs.readFile(document.uri);
                    const bytes=document.restored ?? document.source;document.restored=undefined;
                    const ext=extname(document.uri.path).slice(1).toLowerCase(),workbook=WORKBOOKS.has(ext);
                    info=['xlsx','xlsm'].includes(ext)?await inspectWorkbook(bytes):undefined;
                    const backend=vscode.workspace.getConfiguration('ahmedSpreadsheet').get('backend','auto');
                    // Viewing never starts or probes the native runtime. Keep the original bytes
                    // and an immutable preview baseline; open UNO only on edit/recalculation.
                    document.nativeSource=bytes;
                    document.nativeAllowed=workbook&&vscode.workspace.isTrusted&&['auto','libreoffice'].includes(backend);
                    if(workbook&&!document.nativeAllowed)loadWarning=!vscode.workspace.isTrusted?'Trust this workspace to enable native editing.':'Select the LibreOffice backend to enable native workbook editing.';
                    if(version!==document.generation||disposed){await document.native?.dispose();return;}
                    const readOnly=document.readonly||(workbook&&!document.nativeAllowed);
                    await panel.webview.postMessage({type:'open',content:{buffer:Array.from(bytes),ext:'.'+ext,readOnly,documentCacheId:key,info,nativeEditing:!!document.nativeAllowed&&!document.readonly,nativeAvailable:!!document.nativeAllowed,nativeWarning:loadWarning}});
                }catch(e){await panel.webview.postMessage({type:'loadError',content:String(e instanceof Error?e.message:e)});}
            };
            document.load=load;
            const session:any={ready,recalculate:()=>panel.webview.postMessage({type:'recalculateRequest'})};
            if(testMode){
                session.action=(content:any)=>document.request!('testAction',content);
                session.save=()=>provider.saveCustomDocument(document,new vscode.CancellationTokenSource().token);
                session.saveAs=(target:vscode.Uri)=>provider.saveCustomDocumentAs(document,target,new vscode.CancellationTokenSource().token);
                session.isDirty=()=>document.dirty;
                session.nativeStarted=()=>!!document.native;
                session.nativeInspect=()=>enqueue(document,async()=> (await ensureNative(document)).request('inspect'));
                session.idle=()=>document.queue;
                session.cancel=()=>document.native?.cancel();
            }
            sessions.set(key,session);
            panel.webview.options={enableScripts:true,localResourceRoots:[vscode.Uri.joinPath(context.extensionUri,'out','webview')]};
            const listeners:vscode.Disposable[]=[];
            const report=async(e:unknown)=>{
                const message=String(e instanceof Error?e.message:e);output.appendLine(message);
                await status('Native document: '+message);
                await panel.webview.postMessage({type:'nativeFailure',content:message});
            };
            const background=(action:()=>Promise<any>)=>{void enqueue(document,action).catch(report);};
            listeners.push(panel.webview.onDidReceiveMessage(async message=>{
                const p=message.content;
                try {
                    if(['exportResponse','modelResponse','testResponse'].includes(message.type)){
                        const pending=requests.get(p.id);
                        if(pending){clearTimeout(pending.timer);requests.delete(p.id);p.error?pending.reject(new Error(p.error)):pending.resolve(p);}return;
                    }
                    if(message.type==='init')await enqueue(document,load);
                    else if(message.type==='previewReady'){
                        await enqueue(document,async()=>{
                            document.model=structuredClone(p.model);
                            await status(loadWarning||('Saved results'+(info?.missingResults?' · '+info.missingResults+' results need calculation':'')));
                            resolveReady!(p);if(document.dirty)changes.fire({document});
                        });
                        if(autoGeneration!==document.generation&&info?.missingResults&&vscode.workspace.getConfiguration('ahmedSpreadsheet').get('autoCalculateMissing',false)){autoGeneration=document.generation;background(()=>calculate());}
                    }else if(message.type==='change'){
                        if(!document.readonly){document.revision=p?.revision ?? document.revision;markChanged(document);}
                    }else if(message.type==='nativeSync'){
                        if(document.readonly)throw new Error('File is read only');
                        background(async()=>{
                            if(p.revision!==document.revision)return;
                            const result=await syncModel(document,p.model);
                            if(result.changed&&vscode.workspace.getConfiguration('ahmedSpreadsheet').get('autoRecalculate',true)){await status('Calculating with LibreOffice UNO…');await nativeResults(p.revision);}
                            await panel.webview.postMessage({type:'nativeSynced',content:{revision:p.revision,history:await (await ensureNative(document)).request('inspect')}});
                        });
                    }else if(message.type==='nativeCommand'){
                        if(document.readonly)throw new Error('File is read only');
                        document.revision=p.revision;
                        background(async()=>{
                            let changed=false;
                            if(p.command.action!=='restore'){
                                await syncModel(document,p.model);
                                changed=!!(await (await ensureNative(document)).operation(p.command)).changed;
                                if(changed)markChanged(document);
                            }
                            const snapshot=await (await ensureNative(document)).snapshot();
                            const projected=await document.request!('nativeSnapshot',{buffer:Array.from(snapshot.bytes),info:snapshot.info,revision:p.revision,selection:p.selection,history:snapshot.history});
                            await (await ensureNative(document)).bind(projected.model);document.model=projected.model;
                            await panel.webview.postMessage({type:'nativeReady',content:{revision:p.revision,changed,restored:p.command.action==='restore',history:await (await ensureNative(document)).request('inspect')}});
                        });
                    }else if(message.type==='nativeProjectionAccepted'){
                        background(async()=>{if(p.revision===document.revision&&document.native?.alive){await document.native.bind(p.model);document.model=p.model;}});
                    }else if(message.type==='recalculate')background(()=>calculate(p?.revision ?? 0));
                    else if(message.type==='cancel'){controller?.abort();document.nativeOpening?.abort();await document.native?.cancel();await status('Calculation cancelled · unsaved native document can reconnect from its checkpoint');}
                    else if(message.type==='save'||message.type==='nativeSave'){
                        if(document.readonly)throw new Error('File is read only. Use Save As.');
                        if(!document.dirty)markChanged(document);
                        await vscode.commands.executeCommand('workbench.action.files.save');
                    }else if(message.type==='saveAs'||message.type==='nativeSaveAs'){
                        const ext=String(p?.ext ?? 'xlsx').replace(/[^a-z0-9]/g,'');
                        const name=basename(document.uri.path,extname(document.uri.path))+'.'+ext;
                        const target=await vscode.window.showSaveDialog({defaultUri:vscode.Uri.joinPath(document.uri,'..',name),filters:{Spreadsheet:[ext]}});
                        if(target){
                            if(target.toString()===key){if(!document.dirty)markChanged(document);await vscode.commands.executeCommand('workbench.action.files.save');}
                            else {await enqueue(document,async()=>write(document,target,await serialize(document,ext)));await vscode.commands.executeCommand('vscode.openWith',target,VIEW);}
                        }
                    }else if(message.type==='openDesktop')await vscode.env.openExternal(document.uri);
                    else if(message.type==='openExternal'){
                        const uri=vscode.Uri.parse(String(p));if(['http','https','mailto'].includes(uri.scheme))await vscode.env.openExternal(uri);
                    }else if(message.type==='editInVSCode')await vscode.commands.executeCommand('vscode.openWith',document.uri,'default');
                    else if(message.type==='developerTool')await vscode.commands.executeCommand('workbench.action.toggleDevTools');
                    else if(message.type==='renderError'){output.appendLine('Preview: '+String(p));resolveReady!({error:String(p)});}
                    else if(message.type==='calculationRendered')session.calculation=p;
                    else if(message.type==='nativeRendered')session.nativeRender=p;
                }catch(e){await report(e);}
            }));
            if(document.uri.scheme==='file'){
                const watcher=vscode.workspace.createFileSystemWatcher(new vscode.RelativePattern(vscode.Uri.joinPath(document.uri,'..'),basename(document.uri.path)));
                listeners.push(watcher,watcher.onDidChange(()=>{
                    if(Date.now()-(document.lastWrite ?? 0)<1500)return;
                    if(document.dirty)void status('File changed on disk · Save As to keep your edits');
                    else background(load);
                }),watcher.onDidDelete(()=>void panel.webview.postMessage({type:'loadError',content:'Source workbook was removed.'})));
            }
            listeners.push(panel.onDidDispose(()=>{
                disposed=true;document.closed=true;document.nativeOpening?.abort();controller?.abort();sessions.delete(key);listeners.forEach(l=>l.dispose());void document.native?.dispose();
                for(const pending of requests.values()){clearTimeout(pending.timer);pending.reject(new Error('Editor closed'));}requests.clear();resolveReady!({error:'Preview closed'});
            }));
            let html=new TextDecoder().decode(await vscode.workspace.fs.readFile(vscode.Uri.joinPath(context.extensionUri,'out','webview','spreadsheet.html')));
            const base=panel.webview.asWebviewUri(vscode.Uri.joinPath(context.extensionUri,'out','webview')).toString(),nonce=randomBytes(18).toString('hex');
            const csp=`default-src 'none'; script-src 'nonce-${nonce}' ${panel.webview.cspSource}; style-src ${panel.webview.cspSource} 'unsafe-inline'; font-src ${panel.webview.cspSource} data:; img-src ${panel.webview.cspSource} data: blob:; connect-src ${panel.webview.cspSource};`;
            panel.webview.html=html.replace('{{configs}}',JSON.stringify({language:vscode.env.language,route:'excel',testMode}))
                .replace('<base href="/">','<base href="'+base+'/">').replace('<head>','<head><meta http-equiv="Content-Security-Policy" content="'+csp+'">').replace(/<script\b/g,'<script nonce="'+nonce+'"');
        }
    };
    context.subscriptions.push(vscode.window.registerCustomEditorProvider(VIEW,provider,{webviewOptions:{retainContextWhenHidden:true},supportsMultipleEditorsPerDocument:false}));
    context.subscriptions.push(vscode.commands.registerCommand('ahmedSpreadsheet.clearCache',async()=>{
        await fs.rm(cacheRoot,{recursive:true,force:true});vscode.window.showInformationMessage('Spreadsheet calculation cache cleared.');
    }));
    return {sessions,viewType:VIEW};
}
