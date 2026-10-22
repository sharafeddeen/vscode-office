import * as vscode from 'vscode';
import assert from 'node:assert/strict';
import {promises as fs} from 'node:fs';
import {join} from 'node:path';
import {createHash} from 'node:crypto';
import JSZip from 'jszip';
import {readResults} from '../../src/spreadsheet/results';
const hash=(b:Uint8Array)=>createHash('sha256').update(b).digest('hex');

async function wait<T>(read:()=>T|undefined,label:string,timeout=60000):Promise<T>{
    const start=Date.now();
    while(Date.now()-start<timeout){const result=read();if(result)return result;await new Promise(r=>setTimeout(r,100));}
    throw new Error('Timed out waiting for '+label);
}
async function bounded<T>(promise:Promise<T>,label:string):Promise<T>{
    let timer:NodeJS.Timeout;
    try{return await Promise.race([promise,new Promise<T>((_,reject)=>{timer=setTimeout(()=>reject(new Error('Timed out waiting for '+label)),60000);})]);}
    finally{clearTimeout(timer!);}
}
export async function run(){
    const extension=vscode.extensions.getExtension('sharafeddeen.ahmed-spreadsheet-viewer');
    assert.ok(extension,'Development extension registered');
    const api=await extension.activate();
    const folder=process.env.AHMED_INTEGRATION_DIR!;
    const saved=vscode.Uri.file(join(folder,'saved.xlsx'));
    await vscode.workspace.getConfiguration('ahmedSpreadsheet').update('autoCalculateMissing',false,vscode.ConfigurationTarget.Global);
    const previewStart=Date.now();
    await vscode.workspace.getConfiguration('ahmedSpreadsheet').update('pythonPath','C:/nonexistent-viewing-test/python.exe',vscode.ConfigurationTarget.Global);
    await vscode.commands.executeCommand('vscode.openWith',saved,api.viewType);
    const session=await wait(()=>api.sessions.get(saved.toString()),'saved webview session');
    const rendered:any=await bounded(session.ready,'saved preview render');
    await fs.writeFile(join(folder,'rendered-debug.json'),JSON.stringify(rendered,null,2));
    assert.ok(!rendered.error,rendered.error);
    assert.equal(rendered.warnings?.length ?? 0,0,'Primary reader must preserve formatting in the browser');
    assert.ok(rendered.canvasCount>0,'Real canvas created in webview');
    assert.equal(rendered.sheets.length,3);
    assert.equal(rendered.toolbarCount,1);assert.ok(rendered.toolbarButtons>20);assert.equal(rendered.formulaBarCount,1);assert.ok(rendered.editable);assert.ok(rendered.themeToggle);
    assert.equal(rendered.sheets[0].cells['4:7'],'OK');
    assert.equal(rendered.sheets[0].cells['4:6'],'100%');
    assert.equal(rendered.sheets[0].cells['9:0'],'');
    assert.equal(rendered.sheets[0].cells['10:0'],'#N/A');


    const previewMs=Date.now()-previewStart;
    assert.equal(session.nativeStarted(),false,'Viewing must not start LibreOffice');
    const before=hash(await fs.readFile(saved.fsPath));
    for(const jitter of [0,7]){
        const browsed=await session.action({action:'browse',sheets:[1,2,0],jitter});
        assert.ok(!browsed.error,browsed.error);
        assert.equal(browsed.activeSheet,0,'Sheet clicks including pointer jitter select the requested sheet');
        assert.deepEqual(browsed.sheets.map((s:any)=>s.name),rendered.sheets.map((s:any)=>s.name),'Browsing must never reorder sheets');
        assert.equal(browsed.revision,0,'Browsing must not emit an edit');
        assert.ok(!browsed.status.includes('edits applied'));
    }
    await session.action({action:'noopChange'});
    await new Promise(r=>setTimeout(r,400));
    assert.equal(session.isDirty(),false,'Sheet switching and unchanged change events keep the editor clean');
    assert.equal(session.nativeStarted(),false,'Browsing must not launch the native runtime');
    assert.equal(hash(await fs.readFile(saved.fsPath)),before);
    console.log('Saved preview rendered in '+previewMs+'ms; browsing kept native runtime stopped.');
    const matches=await session.action({action:'find',text:'Heat allocation'});assert.equal(matches.matches.length,1);
    await vscode.workspace.getConfiguration('ahmedSpreadsheet').update('pythonPath','',vscode.ConfigurationTarget.Global);
    assert.equal(session.nativeStarted(),false);
    const edited=await session.action({action:'edit',ri:4,ci:1,text:'0.5'});
    assert.equal(edited.sheets[0].cells['4:1'],'0.5');assert.equal(edited.sheets[0].cells['4:7'],'Not calculated');
    await wait(()=>session.isDirty(),'VS Code dirty state');
    assert.equal(hash(await fs.readFile(saved.fsPath)),before,'Editing cannot save implicitly');
    const undo=await session.action({action:'undo'});assert.ok(!undo.error,undo.error);assert.equal(Number(undo.sheets[0].cells['4:1']),0.4);
    assert.equal(session.nativeStarted(),true,'First edit starts UNO using the unedited preview baseline');
    const redo=await session.action({action:'redo'});assert.ok(!redo.error,redo.error);assert.equal(Number(redo.sheets[0].cells['4:1']),0.5);
    const bold=await session.action({action:'bold',ri:4,ci:1});
    const cell=bold.data[0].rows['4'].cells['1'];assert.ok(bold.data[0].styles[cell.style].font.bold);
    const themed=await session.action({action:'theme'});assert.equal(typeof themed.theme,'boolean');
    await session.recalculate();
    const editedCalc:any=await wait(()=>session.calculation,'edited calculation',120000);
    assert.equal(editedCalc.sheets[0].cells['4:7'],'FAIL','Recalculation uses unsaved inputs');
    await vscode.commands.executeCommand('workbench.action.files.save');
    await wait(()=>!session.isDirty(),'save acknowledgement');
    assert.notEqual(hash(await fs.readFile(saved.fsPath)),before);
    const copy=vscode.Uri.file(join(folder,'edited-copy.xlsx'));await session.saveAs(copy);assert.ok((await fs.stat(copy.fsPath)).size>0);
    const replaced=await session.action({action:'replace',find:'Heat allocation',replace:'Edited title'});
    assert.equal(replaced.sheets[0].cells['0:0'],'Edited title');
    // Discard this final test-only edit through VS Code's revert lifecycle.
    await vscode.commands.executeCommand('workbench.action.files.revert');
    const missing=vscode.Uri.file(join(folder,'missing.xlsx'));
    await vscode.commands.executeCommand('vscode.openWith',missing,api.viewType);
    const missingSession=await wait(()=>api.sessions.get(missing.toString()),'missing-result webview session');
    const missingRender:any=await bounded(missingSession.ready,'missing-result preview render');
    assert.ok(!missingRender.error,missingRender.error);
    assert.equal(missingRender.sheets[0].cells['4:7'],'Not calculated');
    assert.equal(missingSession.nativeStarted(),false);
    assert.ok(vscode.workspace.isTrusted,'Isolated synthetic integration workspace is trusted');
    await missingSession.recalculate();
    const calculated:any=await wait(()=>{if(missingSession.lastStatus?.includes('unavailable:'))throw new Error(missingSession.lastStatus);return missingSession.calculation;},'native result render',120000);
    assert.equal(calculated.sheets[0].cells['4:7'],'OK');

    const nativeUri=vscode.Uri.file(join(folder,'native.xlsx'));
    await vscode.commands.executeCommand('vscode.openWith',nativeUri,api.viewType);
    const nativeSession=await wait(()=>api.sessions.get(nativeUri.toString()),'native document session');
    const nativeReady:any=await bounded(nativeSession.ready,'native document');
    assert.ok(nativeReady.nativeEditing);
    const nativeSourceHash=hash(await fs.readFile(nativeUri.fsPath));
    const nativeEdit=await nativeSession.action({action:'edit',sheetIndex:1,ri:0,ci:0,text:'0.9'});
    assert.ok(!nativeEdit.error,nativeEdit.error);
    await nativeSession.recalculate();
    await wait(()=>nativeSession.calculation?.revision===nativeEdit.revision?nativeSession.calculation:undefined,'native named-range recalculation');
    assert.equal(nativeSession.calculation.sheets[0].cells['1:9'],100);
    assert.equal(nativeSession.calculation.sheets[0].cells['2:11'],2);
    assert.equal(nativeSession.calculation.sheets[0].cells['3:11'],3);
    assert.equal(hash(await fs.readFile(nativeUri.fsPath)),nativeSourceHash);
    const dateEdit=await nativeSession.action({action:'edit',sheetIndex:0,ri:2,ci:14,text:'2024-02-01'});
    await nativeSession.recalculate();
    const dated:any=await wait(()=>nativeSession.calculation?.revision===dateEdit.revision?nativeSession.calculation:undefined,'native date formatting');
    assert.equal(typeof dated.sheets[0].cells['2:14'],'string','Native date formatting is displayed');
    const undoDate=await nativeSession.action({action:'undo'});
    assert.ok(!undoDate.error,undoDate.error);assert.ok(!undoDate.sheets[0].cells['2:14'],'Undo removes the actual input, not a display-format refresh');
    const inserted=await nativeSession.action({action:'nativeOperation',command:{action:'insertRows',sheet:1,index:0,count:1}});
    assert.ok(!inserted.error,inserted.error);assert.equal(inserted.sheets[0].cells['1:9'],100);
    assert.ok((await nativeSession.nativeInspect()).names.includes('Efficiency'));
    await nativeSession.cancel();
    assert.ok((await nativeSession.nativeInspect()).names.includes('Efficiency'),'Reconnect from the native recovery checkpoint');
    await nativeSession.save();
    const nativeBytes=new Uint8Array(await fs.readFile(nativeUri.fsPath)),savedNative=await JSZip.loadAsync(nativeBytes);
    assert.match(await savedNative.file('xl/workbook.xml').async('string'),/name="Efficiency"/);
    assert.match(await savedNative.file('xl/worksheets/sheet1.xml').async('string'),/conditionalFormatting/);
    assert.equal((await readResults(nativeBytes)).QUICK_SPLITS.J2.value,100);
    const ods=vscode.Uri.file(join(folder,'native-copy.ods'));await nativeSession.saveAs(ods);
    assert.ok((await JSZip.loadAsync(await fs.readFile(ods.fsPath))).file('content.xml'));
    const csv=vscode.Uri.file(join(folder,'literal.csv'));
    await vscode.commands.executeCommand('vscode.openWith',csv,api.viewType);
    const csvSession=await wait(()=>api.sessions.get(csv.toString()),'CSV session');
    const csvRender:any=await bounded(csvSession.ready,'CSV preview render');
    assert.ok(!csvRender.error,csvRender.error);
    assert.equal(csvRender.sheets[0].cells['0:0'],'=ABS(-1)');

    const original=process.env.AHMED_ORIGINAL_WORKBOOK;
    let originalVerified=false;
    if(original){
        const uri=vscode.Uri.file(original);
        await vscode.commands.executeCommand('vscode.openWith',uri,api.viewType);
        const originalSession=await wait(()=>api.sessions.get(uri.toString()),'original workbook session');
        const data:any=await bounded(originalSession.ready,'original workbook preview');
        assert.ok(!data.error,data.error);
        assert.equal(data.sheets.length,6);
        assert.equal(data.sheets.find((s:any)=>s.name==='QUICK_SPLITS').cells['4:7'],'OK');
        originalVerified=true;
    }
    await fs.writeFile(join(folder,'result.json'),JSON.stringify({previewMs,lazyNative:true,browsingClean:true,webview:true,canvas:true,sheets:true,savedResults:true,missingResults:true,nativeCalculation:true,toolbar:true,editing:true,undoRedo:true,formatting:true,findReplace:true,saving:true,nativeNames:true,nativeArrays:true,nativeStructure:true,nativeRecovery:true,originalVerified},null,2));
    await vscode.commands.executeCommand('workbench.action.closeAllEditors');
    console.log('Actual VS Code webview integration checks passed.');
}
