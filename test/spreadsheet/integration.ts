import * as vscode from 'vscode';
import assert from 'node:assert/strict';
import {promises as fs} from 'node:fs';
import {join} from 'node:path';
import {createHash} from 'node:crypto';
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


    const before=hash(await fs.readFile(saved.fsPath));
    const matches=await session.action({action:'find',text:'Heat allocation'});assert.equal(matches.matches.length,1);
    const edited=await session.action({action:'edit',ri:4,ci:1,text:'0.5'});
    assert.equal(edited.sheets[0].cells['4:1'],'0.5');assert.equal(edited.sheets[0].cells['4:7'],'Not calculated');
    await wait(()=>session.isDirty(),'VS Code dirty state');
    assert.equal(hash(await fs.readFile(saved.fsPath)),before,'Editing cannot save implicitly');
    const undo=await session.action({action:'undo'});assert.equal(undo.sheets[0].cells['4:1'],'0.4');
    const redo=await session.action({action:'redo'});assert.equal(redo.sheets[0].cells['4:1'],'0.5');
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
    assert.ok(vscode.workspace.isTrusted,'Isolated synthetic integration workspace is trusted');
    await missingSession.recalculate();
    const calculated:any=await wait(()=>{if(missingSession.lastStatus?.includes('unavailable:'))throw new Error(missingSession.lastStatus);return missingSession.calculation;},'native result render',120000);
    assert.equal(calculated.sheets[0].cells['4:7'],'OK');

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
    await fs.writeFile(join(folder,'result.json'),JSON.stringify({webview:true,canvas:true,sheets:true,savedResults:true,missingResults:true,nativeCalculation:true,toolbar:true,editing:true,undoRedo:true,formatting:true,findReplace:true,saving:true,originalVerified},null,2));
    await vscode.commands.executeCommand('workbench.action.closeAllEditors');
    console.log('Actual VS Code webview integration checks passed.');
}
