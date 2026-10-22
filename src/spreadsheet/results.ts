import JSZip from 'jszip';
import { XMLParser } from 'fast-xml-parser';
import { inspectWorkbook } from './package';
export interface FormulaResult { value: string | number | boolean; type: string }
export type Results = Record<string, Record<string, FormulaResult>>;
const xml = new XMLParser({ignoreAttributes:false,parseTagValue:false,parseAttributeValue:false,trimValues:false,removeNSPrefix:true});
const array = (v: any): any[] => v === undefined ? [] : Array.isArray(v) ? v : [v];
const text = (v: any): string => typeof v === 'object' && v ? String(v['#text'] ?? '') : String(v ?? '');
export async function readResults(buffer: Uint8Array): Promise<Results> {
    const info = await inspectWorkbook(buffer), zip = await JSZip.loadAsync(buffer), results: Results = {};
    const strings = zip.file('xl/sharedStrings.xml');
    const shared = strings ? array(xml.parse(await strings.async('string'))?.sst?.si).map(si => si.t !== undefined ? text(si.t) : array(si.r).map(r=>text(r.t)).join('')) : [];
    for (const sheet of info.sheets) {
        const document = xml.parse(await zip.file(sheet.path)!.async('string'));
        results[sheet.name] = {};
        for (const row of array(document?.worksheet?.sheetData?.row)) for (const cell of array(row.c)) {
            // Include cached values for formula cells and results that a native engine expanded into scalar cells.
            if (!Object.hasOwn(cell,'v')) continue;
            const type = cell['@_t'] ?? 'n', raw = text(cell.v);
            if (raw === '' && type !== 'str') continue;
            const value = type === 's' ? shared[Number(raw)] : type === 'b' ? raw === '1' : type === 'n' ? Number(raw) : raw;
            if (value !== undefined) results[sheet.name][cell['@_r']] = {type:type==='s'?'str':type,value};
        }
    }
    return results;
}
