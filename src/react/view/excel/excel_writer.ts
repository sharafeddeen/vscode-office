import ExcelJS from '@cweijan/exceljs';
import JSZip from 'jszip';
import { handler } from "../../util/vscode";
import * as XLSX from 'xlsx';
import Spreadsheet from './x-spreadsheet/index';
import type { CellData, RowData, SheetData } from './x-spreadsheet/index';
import { CsvEncoding, encodeCsvText } from './csvEncoding';
import { DEFAULT_ROW_HEIGHT_PX, freezeExprToExcelView, pxToExcelRowHeight } from './excel_meta';
import { patchWorksheetSortStateXml } from './excel_sort_state';
import { applySpreadsheetStyle } from './excel_styles';
import { hyperlinkKey, writeCellHyperlink, type SpreadsheetHyperlink } from './excel_hyperlink';
import { writeWorksheetValidations } from './excel_validation';
import { writeWorksheetProtection } from './excel_protection';
import { writeWorksheetImages } from './excel_images';

const DEFAULT_COL_WIDTH = 100;

export { buildFormattingSnapshot, hasFormattingChanged } from './excel_meta';

export interface ExportOptions {
    /** 通过另存为对话框保存，而非覆盖当前文件 */
    saveAs?: boolean;
    /** saveAs 时指定目标格式 */
    saveAsExt?: string;
}

function isRowData(row: RowData | number | undefined): row is RowData {
    return row != null && typeof row === 'object';
}

function getColWidth(cols: SheetData['cols'], ci: number) {
    const col = cols?.[ci];
    if (col && typeof col === 'object' && col.width != null) return col.width;
    return DEFAULT_COL_WIDTH;
}

function pxToExcelColWidth(px: number) {
    return Math.max((px - 5) / 7, 0);
}

function setCellValue(cell: ExcelJS.Cell, data: CellData) {
    const text=data.text ?? '';
    if(data.previewCell && data.hasSavedValue) {
        const result=data.savedType==='e' ? {error:data.savedValue} : data.savedValue;
        if(data.formulaPreview)cell.value={formula:text.slice(1),result} as any;
        else cell.value=result as any;
        return;
    }
    if(data.previewCell && !data.formulaPreview && text.startsWith('=')){cell.value=text;return;}
    if (!text) {
        cell.value = null;
        return;
    }
    if (text.startsWith('=')) {
        cell.value = { formula: text.slice(1) };
        return;
    }
    const num = Number(text);
    if (text.trim() !== '' && !Number.isNaN(num) && String(num) === text.trim()) {
        cell.value = num;
        return;
    }
    cell.value = text;
}

function applySheetMeta(worksheet: ExcelJS.Worksheet, sheetData: SheetData) {
    if (sheetData.freeze) {
        const frozen = freezeExprToExcelView(sheetData.freeze);
        if (frozen) {
            worksheet.views = [{
                state: 'frozen',
                xSplit: frozen.xSplit,
                ySplit: frozen.ySplit,
                topLeftCell: sheetData.freeze,
            }];
        }
    }

    const autofilter = sheetData.autofilter;
    if (autofilter?.ref) {
        worksheet.autoFilter = autofilter.ref;
    }
}

function writeRowHeights(worksheet: ExcelJS.Worksheet, rows: SheetData['rows']) {
    if (!rows?.len) return;
    for (let ri = 0; ri < rows.len; ri += 1) {
        const row = rows[ri];
        if (!isRowData(row) || row.height == null) continue;
        if (Math.abs(row.height - DEFAULT_ROW_HEIGHT_PX) < 1) continue;
        worksheet.getRow(ri + 1).height = pxToExcelRowHeight(row.height);
    }
}

async function writeSheetToExcelJs(worksheet: ExcelJS.Worksheet, workbook: ExcelJS.Workbook, sheetData: SheetData) {
    const { rows, cols, styles = [], merges = [], hyperlinks = {}, validations } = sheetData;

    if (cols?.len) {
        for (let ci = 0; ci < cols.len; ci += 1) {
            worksheet.getColumn(ci + 1).width = pxToExcelColWidth(getColWidth(cols, ci));
        }
    }

    for (let i = 0; i < merges.length; i += 1) {
        worksheet.mergeCells(merges[i]);
    }

    applySheetMeta(worksheet, sheetData);
    writeRowHeights(worksheet, rows);
    writeWorksheetValidations(worksheet, validations);

    if (!rows) return;

    const rowLen = rows.len ?? 0;
    for (let ri = 0; ri < rowLen; ri += 1) {
        const row = rows[ri];
        if (!isRowData(row) || !row.cells) continue;
        for (const ciKey of Object.keys(row.cells)) {
            const ci = Number(ciKey);
            if (Number.isNaN(ci)) continue;
            const cellData = row.cells[ci];
            const excelCell = worksheet.getCell(ri + 1, ci + 1);
            const hl = hyperlinks[hyperlinkKey(ri, ci)] as SpreadsheetHyperlink | undefined;
            if (hl?.link) {
                writeCellHyperlink(excelCell, cellData.text ?? '', hl);
            } else {
                setCellValue(excelCell, cellData);
                if(cellData.comment)excelCell.note=cellData.comment;
            }
            if (cellData.style != null && styles[cellData.style]) {
                applySpreadsheetStyle(excelCell, styles[cellData.style]);
                if(cellData.numFmt && cellData.numFmtFormat===styles[cellData.style].format)excelCell.numFmt=cellData.numFmt;
            }
        }
    }

    writeWorksheetImages(worksheet, workbook, sheetData.images, sheetData.backgroundImage);
    await writeWorksheetProtection(worksheet, sheetData);
}

async function emitSave(buffer: Uint8Array, options?: ExportOptions) {
    const content = [...buffer];
    if (options?.saveAs) {
        handler.emit('saveAs', { content, ext: options.saveAsExt ?? 'xlsx' });
        return;
    }
    handler.emit('save', content);
}

async function patchWorkbookSortStates(buffer: Uint8Array, sheets: SheetData[]) {
    const zip = await JSZip.loadAsync(buffer);
    for (let i = 0; i < sheets.length; i += 1) {
        const file = zip.file(`xl/worksheets/sheet${i + 1}.xml`);
        if (!file) continue;
        const xml = await file.async('string');
        zip.file(`xl/worksheets/sheet${i + 1}.xml`, patchWorksheetSortStateXml(xml, sheets[i].autofilter));
    }
    return new Uint8Array(await zip.generateAsync({ type: 'uint8array' }));
}

export async function buildXlsx(sheets: SheetData[]) {
    const workbook = new ExcelJS.Workbook();
    workbook.properties.date1904=sheets.some(s=>Object.values(s.rows ?? {}).some((r:any)=>Object.values(r?.cells ?? {}).some((c:any)=>c.date1904)));
    workbook.calcProperties.fullCalcOnLoad=true;
    for (let i = 0; i < sheets.length; i += 1) {
        const sheetData = sheets[i];
        const worksheet = workbook.addWorksheet(sheetData.name || `Sheet${i + 1}`,{state:sheetData.state ?? 'visible'});
        await writeSheetToExcelJs(worksheet, workbook, sheetData);
    }
    const buffer = new Uint8Array(await workbook.xlsx.writeBuffer());
    const patched = await patchWorkbookSortStates(buffer, sheets);
    return patched;
}

async function exportWithExcelJs(sheets:SheetData[],options?:ExportOptions) {
    await emitSave(await buildXlsx(sheets),options);
}

function applyColWidths(ws: XLSX.WorkSheet, xws: SheetData) {
    const cols = xws.cols;
    if (!cols?.len) return;
    const colWidths = [];
    for (let ci = 0; ci < cols.len; ci += 1) {
        colWidths.push({ wpx: getColWidth(cols, ci) });
    }
    ws['!cols'] = colWidths;
}

function dataToSheetJs(xws: SheetData) {
    const aoa: string[][] = [];
    const rowobj = xws.rows;
    if (!rowobj?.len) {
        const ws = XLSX.utils.aoa_to_sheet(aoa);
        applyColWidths(ws, xws);
        return ws;
    }
    for (let ri = 0; ri < rowobj.len; ri += 1) {
        const row = rowobj[ri];
        if (!isRowData(row)) continue;
        aoa[ri] = [];
        for (const ciKey of Object.keys(row.cells ?? {})) {
            const ci = Number(ciKey);
            if (Number.isNaN(ci)) continue;
            aoa[ri][ci] = row.cells[ci].text;
        }
    }
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    for(const [ri,row] of Object.entries(rowobj)) {
        if(!isRowData(row as any))continue;
        for(const [ci,data] of Object.entries((row as RowData).cells ?? {})) {
            const cell:any=ws[XLSX.utils.encode_cell({r:Number(ri),c:Number(ci)})];
            if(!cell)continue;
            if(data.formulaPreview){cell.f=data.text.slice(1);if(data.hasSavedValue){cell.v=data.savedValue;cell.t=data.savedType==='str'?'s':(data.savedType ?? 'n');}else delete cell.v;}
            else if(data.hasSavedValue){cell.v=data.savedValue;cell.t=data.savedType==='str'?'s':(data.savedType ?? 'n');}
        }
    }
    applyColWidths(ws, xws);
    return ws;
}

function exportWithSheetJs(sheets: SheetData[], bookType: XLSX.BookType) {
    const workbook = XLSX.utils.book_new();
    for (let i = 0; i < sheets.length; i += 1) {
        const sheetData = sheets[i];
        XLSX.utils.book_append_sheet(workbook, dataToSheetJs(sheetData), sheetData.name || `Sheet${i + 1}`);
    }
    const buffer = XLSX.write(workbook, { bookType, type: 'array' });
    handler.emit('save', [...new Uint8Array(buffer)]);
}

export async function exportSaveAs(
    spreadSheet: Spreadsheet,
    targetExt: string,
    csvEncoding: CsvEncoding = 'utf8',
    csvDelimiter: string = ',',
) {
    const ext = targetExt.replace('.', '').toLowerCase();
    const sheets = spreadSheet.getData();
    if (ext === 'xlsx' || ext === 'xlsm') {
        await exportWithExcelJs(sheets, { saveAs: true, saveAsExt: ext });
        return;
    }
    if (ext === 'xls' || ext === 'ods') {
        const wb = XLSX.utils.book_new();
        sheets.forEach((s, i) => {
            const ws = dataToSheetJs(s);
            XLSX.utils.book_append_sheet(wb, ws, s.name || `Sheet${i + 1}`);
        });
        const buf = XLSX.write(wb, { bookType: ext as XLSX.BookType, type: 'array' });
        handler.emit('saveAs', { content: [...new Uint8Array(buf)], ext });
        return;
    }
    if (ext === 'csv' || ext === 'tsv') {
        const fs = ext === 'tsv' ? '\t' : csvDelimiter;
        const csvContent = XLSX.utils.sheet_to_csv(dataToSheetJs(sheets[0]), { FS: fs });
        const bytes = encodeCsvText(csvContent, csvEncoding);
        handler.emit('saveAs', { content: [...bytes], ext });
    }
}

export async function export_xlsx(
    spreadSheet: Spreadsheet,
    extName: string,
    csvEncoding: CsvEncoding = 'utf8',
    options?: ExportOptions,
    csvDelimiter: string = ',',
) {
    const ext = extName.replace('.', '').toLowerCase();
    const sheets = spreadSheet.getData();

    if (ext === 'xlsx' || ext === 'xlsm' || options?.saveAs) {
        await exportWithExcelJs(sheets, options?.saveAs ? { saveAs: true } : undefined);
        return;
    }
    if (ext === 'xls' || ext === 'ods') {
        exportWithSheetJs(sheets, ext);
        return;
    }
    if (ext === 'csv' || ext === 'tsv') {
        const fs = ext === 'tsv' ? '\t' : csvDelimiter;
        const csvContent = XLSX.utils.sheet_to_csv(dataToSheetJs(sheets[0]), { FS: fs });
        const bytes = encodeCsvText(csvContent, csvEncoding);
        handler.emit('save', [...bytes]);
    }
}

// Shared by VS Code's Save/backup lifecycle and native calculation of current edits.
export async function serializeSpreadsheet(sheets:SheetData[],ext='xlsx',encoding:CsvEncoding='utf8',delimiter=','):Promise<Uint8Array> {
    ext=ext.replace(/^\./,'').toLowerCase();
    if(ext==='xlsx'||ext==='xlsm')return buildXlsx(sheets);
    if(ext==='csv'||ext==='tsv')return encodeCsvText(XLSX.utils.sheet_to_csv(dataToSheetJs(sheets[0]),{FS:ext==='tsv'?'\t':delimiter}),encoding);
    const workbook=XLSX.utils.book_new();
    sheets.forEach((s,i)=>XLSX.utils.book_append_sheet(workbook,dataToSheetJs(s),s.name || 'Sheet'+(i+1)));
    return new Uint8Array(XLSX.write(workbook,{bookType:ext as XLSX.BookType,type:'array'}));
}
