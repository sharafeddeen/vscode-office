import {test} from 'node:test';
import assert from 'node:assert/strict';
import {markDerived,applyNativeResults} from '../../src/spreadsheet/native-model';
import type {SheetData} from '../../src/react/view/excel/x-spreadsheet';
test('native overlays include array outputs, protect derived cells and clear removed outputs',()=>{
    const sheets:SheetData[]=[{name:'Sheet',rows:{len:3,0:{cells:{0:{text:'=SEQUENCE(3,1)',previewCell:true,formulaPreview:true}}},1:{cells:{0:{text:'0',previewCell:true}}},2:{cells:{0:{text:'0',previewCell:true}}}},cols:{len:1}}];
    const arrays={Sheet:[{anchor:'A1',ref:'A1:A3'}]};
    markDerived(sheets,arrays);
    assert.equal(sheets[0].rows![1].cells[0].editable,false);
    applyNativeResults(sheets,{Sheet:{A1:{value:1,type:'n'},A2:{value:2,type:'n'},A3:{value:3,type:'n'}}},arrays);
    assert.equal(sheets[0].rows![1].cells[0].savedValue,2);
    assert.equal(sheets[0].rows![2].cells[0].savedValue,3);
    applyNativeResults(sheets,{Sheet:{A1:{value:1,type:'n'}}},{Sheet:[{anchor:'A1',ref:'A1'}]});
    assert.equal(sheets[0].rows![1].cells[0],undefined);
    assert.equal(sheets[0].rows![2].cells[0],undefined);
});
