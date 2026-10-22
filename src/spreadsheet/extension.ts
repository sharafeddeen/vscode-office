import * as vscode from 'vscode';
import { randomBytes } from 'node:crypto';
import { promises as fs } from 'node:fs';
import { join } from 'node:path';
import { inspectWorkbook } from './package';
import { findBackend, recalculate } from './native';

const VIEW = 'sharafeddeen.spreadsheetViewer';
export function activate(context: vscode.ExtensionContext) {
    const sessions = new Map<string, {recalculate:()=>Promise<void>; ready:Promise<any>}>();
    const cacheRoot = join(context.globalStorageUri.fsPath,'calculation-cache');
    const output = vscode.window.createOutputChannel('Ahmed Spreadsheet Viewer');
    context.subscriptions.push(output);
    const provider: vscode.CustomReadonlyEditorProvider = {
        openCustomDocument(uri) {return {uri,dispose(){}};},
        async resolveCustomEditor(document, panel) {
            const key = document.uri.toString();
            let source: Uint8Array, info: Awaited<ReturnType<typeof inspectWorkbook>> | undefined;
            let autoScheduledGeneration = -1;
            let generation = 0, controller: AbortController | undefined, disposed = false;
            let resolveReady: (value:any)=>void;
            const ready = new Promise(resolve=>{resolveReady=resolve;});
            const status = (value:string, warnings:string[] = []) => {const session=sessions.get(key) as any;if(session)session.lastStatus=value;return panel.webview.postMessage({type:'status',content:{value,warnings}});};
            const calculate = async (force=false) => {
                if (!source || !info?.formulaCount || disposed) return;
                if (!vscode.workspace.isTrusted) {await status('Saved results · trust this workspace to recalculate',info.warnings);return;}
                const version = generation;
                controller?.abort(); controller = new AbortController();
                const localController = controller;
                try {
                    const config = vscode.workspace.getConfiguration('ahmedSpreadsheet');
                    const backend = await findBackend({backend:config.get('backend'),libreOfficePath:config.get('libreOfficePath')});
                    if (!backend) {await status('Saved results · no calculation backend available',info.warnings); return;}
                    if (version !== generation || disposed || localController.signal.aborted) return;
                    await status('Calculating with '+backend.kind+'…',info.warnings);
                    const results = await vscode.window.withProgress({location:vscode.ProgressLocation.Notification,title:'Calculating spreadsheet preview',cancellable:true},async (_,token)=>{
                        const subscription = token.onCancellationRequested(()=>localController.abort());
                        try {return await recalculate(source,backend,{cacheRoot,excelScript:join(context.extensionPath,'out','backend','recalculate-excel.ps1'),timeoutMs:config.get<number>('timeoutSeconds',90)*1000,signal:localController.signal,force});}
                        finally {subscription.dispose();}
                    });
                    if (version !== generation || disposed) return;
                    await panel.webview.postMessage({type:'calculated',content:{results,backend:backend.kind}});
                } catch(e) {
                    if (version !== generation || disposed) return;
                    const message = e instanceof Error ? e.message : String(e);
                    output.appendLine('Calculation: '+message);
                    await status('Saved results · calculation unavailable: '+message,info?.warnings);
                }
            };
            const load = async () => {
                const version = ++generation; controller?.abort();
                try {
                    source = await vscode.workspace.fs.readFile(document.uri);
                    const ext = document.uri.path.slice(document.uri.path.lastIndexOf('.')).toLowerCase();
                    info = ['.xlsx','.xlsm'].includes(ext) ? await inspectWorkbook(source) : undefined;
                    if (version !== generation || disposed) return;
                    await panel.webview.postMessage({type:'open',content:{buffer:Array.from(source),ext,readOnly:true,documentCacheId:key,info}});
                } catch(e) {await panel.webview.postMessage({type:'loadError',content:String(e instanceof Error?e.message:e)});}
            };
            sessions.set(key,{recalculate:()=>calculate(true),ready});
            panel.webview.options={enableScripts:true,localResourceRoots:[vscode.Uri.joinPath(context.extensionUri,'out','webview')]};
            const listeners:vscode.Disposable[] = [];
            listeners.push(panel.webview.onDidReceiveMessage(async message=>{
                if (message.type==='init') await load();
                else if (message.type==='previewReady') {
                    resolveReady!(message.content);
                    if (autoScheduledGeneration !== generation && info?.missingResults && vscode.workspace.getConfiguration('ahmedSpreadsheet').get('autoCalculateMissing',true)) {autoScheduledGeneration=generation;void calculate();}
                } else if (message.type==='recalculate') await calculate(true);
                else if (message.type==='cancel') controller?.abort();
                else if (message.type==='openDesktop') await vscode.env.openExternal(document.uri);
                else if (message.type==='renderError') {output.appendLine('Preview: '+String(message.content));resolveReady!({error:String(message.content)});}
                else if (message.type==='calculationRendered') {
                    // Kept only in this session for integration verification; never sent to a service.
                    const session = sessions.get(key) as any; if(session) session.calculation = message.content;
                }
            }));
            if (document.uri.scheme === 'file') {
                const watcher = vscode.workspace.createFileSystemWatcher(new vscode.RelativePattern(vscode.Uri.joinPath(document.uri,'..'),document.uri.path.split('/').pop()!));
                listeners.push(watcher,watcher.onDidChange(()=>void load()),watcher.onDidDelete(()=>void panel.webview.postMessage({type:'loadError',content:'Source workbook was removed.'})));
            }
            listeners.push(panel.onDidDispose(()=>{disposed=true;controller?.abort();sessions.delete(key);listeners.forEach(l=>l.dispose());resolveReady!({error:'Preview closed'});}));
            const htmlPath = vscode.Uri.joinPath(context.extensionUri,'out','webview','spreadsheet.html');
            let html = new TextDecoder().decode(await vscode.workspace.fs.readFile(htmlPath));
            const base = panel.webview.asWebviewUri(vscode.Uri.joinPath(context.extensionUri,'out','webview')).toString();
            const nonce = randomBytes(18).toString('hex');
            const csp = `default-src 'none'; script-src 'nonce-${nonce}' ${panel.webview.cspSource}; style-src ${panel.webview.cspSource} 'unsafe-inline'; font-src ${panel.webview.cspSource} data:; img-src ${panel.webview.cspSource} data: blob:; connect-src ${panel.webview.cspSource};`;
            html = html.replace('{{configs}}',JSON.stringify({language:vscode.env.language,route:'excel'}))
                .replace('<base href="/">','<base href="'+base+'/">')
                .replace('<head>','<head><meta http-equiv="Content-Security-Policy" content="'+csp+'">')
                .replace(/<script\b/g,'<script nonce="'+nonce+'"');
            panel.webview.html = html;
        }
    };
    context.subscriptions.push(vscode.window.registerCustomEditorProvider(VIEW,provider,{webviewOptions:{retainContextWhenHidden:true},supportsMultipleEditorsPerDocument:false}));
    context.subscriptions.push(vscode.commands.registerCommand('ahmedSpreadsheet.clearCache',async()=>{
        // A fixed extension-owned cache directory, never a source/workspace path.
        await fs.rm(cacheRoot,{recursive:true,force:true});
        vscode.window.showInformationMessage('Spreadsheet calculation cache cleared.');
    }));
    return {sessions,viewType:VIEW};
}
