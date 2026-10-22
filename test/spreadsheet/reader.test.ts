import {test} from 'node:test';
import assert from 'node:assert/strict';
import {createHash} from 'node:crypto';
import JSZip from 'jszip';
import {fixture} from './fixture';
import {normalizePackage,inspectWorkbook,resolvePart} from '../../src/spreadsheet/package';
import {loadSheets} from '../../src/react/view/excel/excel_reader';
import {displayCellText} from '../../src/react/view/excel/excel_cell_display';
import {findAllInSheets} from '../../src/react/view/excel/excel_find';
import {findBackend,runProcess} from '../../src/spreadsheet/native';

const hash=(b:Uint8Array)=>createHash('sha256').update(b).digest('hex');
const get=(s:any,r:number,c:number)=>s.rows[r].cells[c];
const display=(s:any,r:number,c:number)=>{const cell=get(s,r,c);return displayCellText(cell,s.styles?.[cell.style]??{},(x,y)=>s.rows[y]?.cells[x]);};
test('arbitrary comment paths load without touching worksheets, formulas or source',async()=>{
    const source=await fixture({unusualComments:true}), before=hash(source);
    const normalized=await normalizePackage(source);
    const a=await JSZip.loadAsync(source),b=await JSZip.loadAsync(normalized);
    for(const name of Object.keys(a.files).filter(n=>n.startsWith('xl/worksheets/')&&n.endsWith('.xml')))assert.equal(await a.file(name)!.async('string'),await b.file(name)!.async('string'));
    const info=await inspectWorkbook(source);assert.deepEqual(info.sheets.map(s=>s.state),['visible','visible','hidden']);
    const data=await loadSheets(source.buffer as ArrayBuffer,'.xlsx'),sheet=data.sheets[0];
    assert.equal(data.sheets.length,3);assert.equal(sheet.name,'QUICK_SPLITS');
    assert.equal(display(sheet,4,7),'OK');
    assert.match(get(sheet,4,7).text,/ABS/);
    assert.match(get(sheet,4,7).comment,/percentages/);
    assert.equal(display(sheet,4,6),'100%');
    assert.deepEqual(sheet.merges,['A1:C1']);assert.ok((sheet.cols?.[0] as any).width>250);
    assert.equal(data.warnings?.length,0);
    assert.equal(hash(source),before);
});
test('saved zero, false, empty text, errors, cross-sheet formula and literal equals are distinct',async()=>{
    const source=await fixture();const data=await loadSheets(source.buffer as ArrayBuffer,'.xlsx');const s=data.sheets[0];
    assert.equal(display(s,7,0),0);assert.equal(display(s,8,0),'FALSE');assert.equal(display(s,9,0),'');
    assert.equal(display(s,10,0),'#N/A');assert.equal(display(s,11,0),'=this is literal text');assert.equal(display(s,12,0),2);
    assert.equal(display(s,13,0),1);
    assert.equal(display(s,14,0),'2024-01-01');
    assert.equal((await inspectWorkbook(source)).missingResults,0);
    assert.ok(findAllInSheets(data.sheets,'OK',{allSheets:true}).length);
});
test('missing results are explicit, not invented errors or zero',async()=>{
    const source=await fixture({missing:true});assert.equal((await inspectWorkbook(source)).missingResults,1);
    const data=await loadSheets(source.buffer as ArrayBuffer,'.xlsx');assert.equal(display(data.sheets[0],4,7),'Not calculated');
});
test('broken packages and missing referenced comments produce useful errors',async()=>{
    await assert.rejects(loadSheets(new Uint8Array([1,2,3]).buffer,'.xlsx'));
    const source=await fixture({unusualComments:true});const zip=await JSZip.loadAsync(source);
    zip.remove('xl/notes/custom-note.xml');
    await assert.rejects(normalizePackage(await zip.generateAsync({type:'uint8array'})),/Missing comment/);
});
test('relationship traversal stays inside the package',()=>{
    assert.equal(resolvePart('xl/worksheets/sheet1.xml','../comments1.xml'),'xl/comments1.xml');
    assert.equal(resolvePart('xl/worksheets/sheet1.xml','/xl/notes/a.xml'),'xl/notes/a.xml');
    assert.throws(()=>resolvePart('xl/a.xml','../../escape'),/escapes/);
});
test('disabled backend and process cancellation are explicit',async()=>{
    assert.equal(await findBackend({backend:'none'}),undefined);
    const controller=new AbortController();controller.abort();
    await assert.rejects(runProcess(process.execPath,['-e','setTimeout(()=>{},10000)'],{signal:controller.signal}),/cancelled/);
    await assert.rejects(runProcess(process.execPath,['-e','setTimeout(()=>{},10000)'],{timeoutMs:100}),/timed out/);
});

test('namespace-prefixed workbook and relationship elements retain sheet identity',async()=>{
    const source=await fixture(),zip=await JSZip.loadAsync(source);
    let workbook=await zip.file('xl/workbook.xml')!.async('string');
    workbook=workbook.replace(/<(\/?)([A-Za-z][A-Za-z0-9]*)(?=[\s/>])/g,'<$1x:$2').replace('xmlns="','xmlns:x="');
    zip.file('xl/workbook.xml',workbook);
    const buffer=await zip.generateAsync({type:'uint8array'});
    assert.deepEqual((await inspectWorkbook(buffer)).sheets.map(s=>s.name),['QUICK_SPLITS','Inputs','Hidden notes']);
    const data=await loadSheets(buffer.buffer as ArrayBuffer,'.xlsx');
    assert.equal(data.sheets.length,3);assert.equal(display(data.sheets[0],4,7),'OK');
});
test('CSV cells beginning with equals are literal values',async()=>{
    const data=await loadSheets(new TextEncoder().encode('=ABS(-1),label').buffer,'.csv');
    assert.equal(display(data.sheets[0],0,0),'=ABS(-1)');
});

test('missing optional comment metadata recovers sheets with an explicit warning',async()=>{
    const source=await fixture({unusualComments:true}),zip=await JSZip.loadAsync(source);
    zip.remove('xl/notes/custom-note.xml');
    const bytes=await zip.generateAsync({type:'uint8array'});
    const data=await loadSheets(bytes.buffer as ArrayBuffer,'.xlsx');
    assert.equal(data.sheets.length,3);assert.equal(display(data.sheets[0],4,7),'OK');
    assert.ok(data.warnings?.some(w=>w.startsWith('Fallback reader')));
});
