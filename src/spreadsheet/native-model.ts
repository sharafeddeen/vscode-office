import {utils} from 'xlsx';
import type {SheetData,CellData} from '../react/view/excel/x-spreadsheet';
import type {Results} from './results';
export type Arrays=Record<string,{anchor:string;ref:string}[]>;
export function markDerived(sheets:SheetData[],arrays:Arrays={}) {
    for(const sheet of sheets)for(const array of arrays[sheet.name!] ?? []) {
        const range=utils.decode_range(array.ref),anchor=utils.decode_cell(array.anchor);
        for(const [ri,row] of Object.entries(sheet.rows ?? {})) {
            if(ri==='len'||Number(ri)<range.s.r||Number(ri)>range.e.r)continue;
            for(const [ci,cell] of Object.entries((row as any).cells ?? {}) as [string,CellData][])
                if(Number(ci)>=range.s.c&&Number(ci)<=range.e.c&&(Number(ri)!==anchor.r||Number(ci)!==anchor.c)){cell.nativeDerived=true;cell.editable=false;}
        }
    }
}
export function applyNativeResults(sheets:SheetData[],results:Results,arrays:Arrays={},projection?:SheetData[]) {
    let missing=0;
    for(const sheet of sheets) {
        sheet.rows ??={len:0};sheet.cols ??={len:0};
        const values=results[sheet.name!] ?? {};
        for(const [ri,row] of Object.entries(sheet.rows)) {
            if(ri==='len')continue;
            for(const [ci,cell] of Object.entries((row as any).cells ?? {}) as [string,CellData][]) {
                const value=values[utils.encode_cell({r:Number(ri),c:Number(ci)})];
                if(cell.nativeDerived&&!value)delete (row as any).cells[ci];
                else if(cell.formulaPreview&&!value){cell.hasSavedValue=false;delete cell.savedValue;delete cell.savedType;missing++;}
            }
        }
        for(const [address,value] of Object.entries(values)) {
            const {r,c}=utils.decode_cell(address);
            const row=(sheet.rows[r] ??= {cells:{}}) as any;
            const cell:CellData=row.cells[c] ??= {text:String(value.value),previewCell:true,nativeDerived:true};
            const projected=projection?.find(s=>s.name===sheet.name);
            const nativeCell=(projected?.rows?.[r] as any)?.cells?.[c] as CellData|undefined;
            if(nativeCell){
                cell.numFmt=nativeCell.numFmt;cell.numFmtFormat=nativeCell.numFmtFormat;cell.date1904=nativeCell.date1904;
                const format=projected?.styles?.[nativeCell.style!]?.format;
                const style=sheet.styles?.[cell.style!] ?? {};
                if(style.format!==format){sheet.styles ??= [];cell.style=sheet.styles.length;sheet.styles.push({...style,format});}
            }
            cell.hasSavedValue=true;cell.savedValue=value.value;cell.savedType=value.type;cell.previewCell=true;
            sheet.rows.len=Math.max(sheet.rows.len ?? 0,r+1);sheet.cols.len=Math.max(sheet.cols.len ?? 0,c+1);
        }
    }
    markDerived(sheets,arrays);
    return missing;
}
