import {test} from 'node:test';
import assert from 'node:assert/strict';
import {fixture} from './fixture';
import {loadSheets} from '../../src/react/view/excel/excel_reader';
import {findAllInSheets,replaceAllInSheets} from '../../src/react/view/excel/excel_find';
import {valueSnapshot,invalidateEditedResults,overlayResults,summarize} from '../../src/spreadsheet/editing';

test('editing inputs invalidates cross-sheet formula caches but formatting does not',async()=>{
    const source=await fixture(),data=await loadSheets(source.buffer as ArrayBuffer,'.xlsx');
    const previous=valueSnapshot(data.sheets);
    data.sheets[0].rows![4].cells[1].style=100;
    assert.equal(invalidateEditedResults(data.sheets,previous),false);
    data.sheets[1].rows![0].cells[0].text='3';
    assert.equal(invalidateEditedResults(data.sheets,previous),true);
    assert.equal(summarize(data.sheets)[0].cells['12:0'],'Not calculated');
    overlayResults(data.sheets,{QUICK_SPLITS:{A13:{value:6,type:'n'}}});
    assert.equal(summarize(data.sheets)[0].cells['12:0'],6);
    assert.equal(summarize(data.sheets)[1].cells['0:0'],'3');
    const beforeSheets=valueSnapshot(data.sheets);data.sheets.push({name:'Empty new sheet'});
    assert.equal(invalidateEditedResults(data.sheets,beforeSheets),true,'Sheet structure can affect formulas');
});
test('XLSX serialization preserves original formulas/results, types, comments and hidden sheets',async()=>{
    // The writer's VS Code message transport is unused by buildXlsx.
    Object.defineProperty(globalThis,'window',{value:{addEventListener(){}},configurable:true});
    Object.defineProperty(globalThis,'navigator',{value:{userAgent:''},configurable:true});
    const {buildXlsx}=await import('../../src/react/view/excel/excel_writer');
    const source=await fixture(),data=await loadSheets(source.buffer as ArrayBuffer,'.xlsx');
    const output=await buildXlsx(data.sheets),reopened=await loadSheets(output.buffer as ArrayBuffer,'.xlsx');
    const cells=summarize(reopened.sheets)[0].cells;
    assert.equal(cells['4:7'],'OK');assert.equal(cells['4:6'],'100%');
    assert.equal(cells['7:0'],0);assert.equal(cells['8:0'],'FALSE');assert.equal(cells['9:0'],'');
    assert.equal(cells['10:0'],'#N/A');assert.equal(cells['11:0'],'=this is literal text');
    assert.equal(cells['14:0'],'2024-01-01');
    assert.match(reopened.sheets[0].rows![4].cells[7].comment!,/percentages/);
    assert.equal(reopened.sheets[2].state,'hidden');
    const previous=valueSnapshot(data.sheets);data.sheets[0].rows![4].cells[1].text='0.5';
    invalidateEditedResults(data.sheets,previous);
    const changed=await buildXlsx(data.sheets),again=await loadSheets(changed.buffer as ArrayBuffer,'.xlsx');
    assert.equal(summarize(again.sheets)[0].cells['4:7'],'Not calculated');
});

test('find matches whole saved results and replacement changes the intended sheet',async()=>{
    const source=await fixture(),data=await loadSheets(source.buffer as ArrayBuffer,'.xlsx');
    assert.equal(findAllInSheets(data.sheets,'OK',{allSheets:true,wholeCell:true}).length,1);
    assert.equal(findAllInSheets(data.sheets,'Heat allocation',{allSheets:true,wholeCell:true}).length,1);
    const before=valueSnapshot(data.sheets);
    assert.equal(replaceAllInSheets(data.sheets,'Heat allocation','New title',{allSheets:true,wholeCell:true}),1);
    invalidateEditedResults(data.sheets,before);
    assert.equal(summarize(data.sheets)[0].cells['0:0'],'New title');
});
