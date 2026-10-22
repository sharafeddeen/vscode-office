import type {CellData, SheetData} from '../react/view/excel/x-spreadsheet';
import {displayCellText} from '../react/view/excel/excel_cell_display';
import type {Results} from './results';

export function visitCells(sheets:SheetData[], visit:(cell:CellData, sheet:SheetData, ri:number, ci:number)=>void) {
    for (const sheet of sheets) for (const [ri,row] of Object.entries(sheet.rows ?? {})) {
        if (ri==='len') continue;
        for (const [ci,cell] of Object.entries((row as any).cells ?? {})) visit(cell as CellData,sheet,Number(ri),Number(ci));
    }
}
export function valueSnapshot(sheets:SheetData[]) {
    const values = new Map<string,string>();
    sheets.forEach((sheet,i)=>values.set(JSON.stringify(['sheet',i]),JSON.stringify([sheet.name,sheet.state,sheet.rows?.len,sheet.cols?.len])));
    visitCells(sheets,(cell,sheet,ri,ci)=>values.set(JSON.stringify([sheet.name,ri,ci]),cell.text ?? ''));
    return values;
}
/** Navigation, zoom, cached values and spare viewport rows/columns are not edits. */
export function editSnapshot(sheets:SheetData[]):string {
    const normalize=(value:any):any=>{
        if(Array.isArray(value))return value.map(normalize);
        if(value&&typeof value==='object')return Object.fromEntries(Object.keys(value).sort().filter(k=>value[k]!==undefined).map(k=>[k,normalize(value[k])]));
        return value;
    };
    return JSON.stringify(normalize(sheets.map(sheet=>{
        const {styles,rows,cols,...properties}=sheet;
        const contentRows=Object.fromEntries(Object.entries(rows ?? {}).filter(([ri])=>ri!=='len').flatMap(([ri,row]:any)=>{
            const {cells,...rowProperties}=row;
            const contentCells=Object.fromEntries(Object.entries(cells ?? {}).flatMap(([ci,value]:any)=>{
                const {previewCell,nativeDerived,date1904,savedValue,savedType,hasSavedValue,numFmtFormat,style,...cell}=value;
                const resolved={...cell,style:styles?.[style]};
                if(!resolved.text)delete resolved.text;
                if(resolved.formulaPreview===false)delete resolved.formulaPreview;
                return Object.keys(resolved).some(k=>resolved[k]!==undefined)?[[ci,resolved]]:[];
            }));
            return Object.keys(rowProperties).length||Object.keys(contentCells).length?[[ri,{...rowProperties,cells:contentCells}]]:[];
        }));
        return {...properties,rows:contentRows,cols:Object.fromEntries(Object.entries(cols ?? {}).filter(([ci])=>ci!=='len'))};
    })));
}
/** Any content/structure change may affect any formula, including cross-sheet references. */
export function invalidateEditedResults(sheets:SheetData[], previous:Map<string,string>, csv=false) {
    const next = valueSnapshot(sheets);
    const changed = next.size!==previous.size || [...next].some(([key,text])=>previous.get(key)!==text);
    if (!changed) return false;
    visitCells(sheets,(cell,sheet,ri,ci)=>{
        const key=JSON.stringify([sheet.name,ri,ci]);
        if (previous.get(key)!==(cell.text ?? '')) {
            cell.previewCell=true;
            cell.formulaPreview=!csv && (cell.text ?? '').startsWith('=');
            delete cell.savedValue;delete cell.savedType;cell.hasSavedValue=false;
        }
        if(cell.formulaPreview){cell.previewCell=true;cell.hasSavedValue=false;delete cell.savedValue;delete cell.savedType;}
    });
    return true;
}
export function overlayResults(sheets:SheetData[], results:Results) {
    let missing=0;
    visitCells(sheets,(cell,sheet,ri,ci)=>{
        if(!cell.formulaPreview)return;
        let n=ci+1,col='';while(n>0){n--;col=String.fromCharCode(65+n%26)+col;n=Math.floor(n/26);}
        const result=results[sheet.name!]?.[col+(ri+1)];
        cell.hasSavedValue=!!result;
        if(result){cell.savedValue=result.value;cell.savedType=result.type;}else{delete cell.savedValue;delete cell.savedType;missing++;}
    });
    return missing;
}
export function summarize(sheets:SheetData[]) {
    return sheets.map(sheet=>({name:sheet.name,cells:Object.fromEntries(Object.entries(sheet.rows ?? {}).filter(([ri])=>ri!=='len').flatMap(([ri,row]:any)=>Object.keys(row.cells ?? {}).slice(0,50).map(ci=>{
        const cell=row.cells[ci];
        return [ri+':'+ci,displayCellText(cell,sheet.styles?.[cell.style] ?? {},(x,y)=>(sheet.rows?.[y] as any)?.cells?.[x])];
    })).slice(0,2000))}));
}
