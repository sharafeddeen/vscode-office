import JSZip from 'jszip';
import { XMLParser, XMLBuilder } from 'fast-xml-parser';

const parser = new XMLParser({ ignoreAttributes: false, parseTagValue: false, parseAttributeValue: false });
const reader = new XMLParser({ignoreAttributes:false,parseTagValue:false,parseAttributeValue:false,removeNSPrefix:true});
const builder = new XMLBuilder({ ignoreAttributes: false });
const list = <T>(value: T | T[] | undefined): T[] => value === undefined ? [] : Array.isArray(value) ? value : [value];

export function resolvePart(source: string, target: string): string {
    const raw = target.startsWith('/') ? target.slice(1) : source.slice(0, source.lastIndexOf('/') + 1) + target;
    const parts: string[] = [];
    for (const part of raw.split('/')) {
        if (part === '..') { if (!parts.length) throw new Error('Package relationship escapes root'); parts.pop(); }
        else if (part && part !== '.') parts.push(part);
    }
    return parts.join('/');
}
function relativePart(source: string, target: string): string {
    const from = source.split('/'); from.pop();
    const to = target.split('/');
    while (from.length && to.length && from[0] === to[0]) { from.shift(); to.shift(); }
    return [...from.map(() => '..'), ...to].join('/');
}
const relSource = (name: string) => name === '_rels/.rels' ? '' : name.replace('/_rels/', '/').replace(/\.rels$/, '');
const local = (object: any, name: string): any => Object.entries(object ?? {}).find(([key])=>key.split(':').pop()===name)?.[1];
const relationships = (document: any) => list<any>(local(local(document,'Relationships'),'Relationship'));

export interface WorkbookInfo {
    sheets: { name: string; state: string; path: string }[];
    missingResults: number;
    formulaCount: number;
    date1904?: boolean;
    warnings: string[];
}
export async function inspectWorkbook(buffer: Uint8Array | ArrayBuffer): Promise<WorkbookInfo> {
    const zip = await JSZip.loadAsync(buffer);
    const workbook = zip.file('xl/workbook.xml');
    const rels = zip.file('xl/_rels/workbook.xml.rels');
    if (!workbook || !rels) throw new Error('Missing XLSX workbook or relationships');
    const relMap = new Map(relationships(reader.parse(await rels.async('string'))).map(r => [r['@_Id'], r]));
    const workbookXml = reader.parse(await workbook.async('string'))?.workbook;
    const date1904 = ['1','true'].includes(String(workbookXml?.workbookPr?.['@_date1904']));
    const sheets = list<any>(workbookXml?.sheets?.sheet).map(s => {
        const rel = relMap.get(s['@_r:id'] ?? s['@_id']);
        if (!rel || rel['@_TargetMode'] === 'External') throw new Error('Invalid worksheet relationship');
        return { name: s['@_name'], state: s['@_state'] ?? 'visible', path: resolvePart('xl/workbook.xml', rel['@_Target']) };
    });
    let missingResults = 0, formulaCount = 0;
    for (const sheet of sheets) {
        const part = zip.file(sheet.path);
        if (!part) throw new Error('Missing worksheet: ' + sheet.name);
        const xml = reader.parse(await part.async('string'));
        for (const row of list<any>(xml?.worksheet?.sheetData?.row)) {
            for (const cell of list<any>(row.c)) {
                if (!Object.hasOwn(cell, 'f')) continue;
                formulaCount++;
                if (!Object.hasOwn(cell, 'v') || (cell.v === '' && cell['@_t'] !== 'str')) missingResults++;
            }
        }
    }
    const names = Object.keys(zip.files);
    const warnings: string[] = [];
    if (names.some(n => /^xl\/charts\//.test(n))) warnings.push('Charts are available in the desktop spreadsheet application.');
    if (names.some(n => /^xl\/pivotTables\//.test(n))) warnings.push('Pivot definitions are not rendered; saved worksheet values are shown.');
    if (names.some(n => /^xl\/externalLinks\//.test(n))) warnings.push('External links are not refreshed.');
    if (names.some(n => /vbaProject\.bin$/i.test(n))) warnings.push('Macros are not executed.');
    if (names.some(n => /^xl\/threadedComments\//.test(n))) warnings.push('Threaded comments are not displayed.');
    return { sheets, formulaCount, missingResults, warnings, date1904 };
}

/** Normalize a disposable ZIP for ExcelJS. The source bytes are never written back.
 * Comments/VML must be located by relationship type, not generator-specific names.
 */
export async function normalizePackage(buffer: ArrayBuffer | Uint8Array): Promise<Uint8Array> {
    const zip = await JSZip.loadAsync(buffer);
    const relDocs = new Map<string, any>();
    const renames = new Map<string, string>();
    let commentIndex = 900001, vmlIndex = 900001;
    for (const name of Object.keys(zip.files).filter(n => n.endsWith('.rels'))) {
        const xml = parser.parse(await zip.file(name)!.async('string'));
        relDocs.set(name, xml);
        for (const rel of relationships(xml)) {
            if (rel['@_TargetMode'] === 'External') continue;
            const type = String(rel['@_Type'] ?? '');
            if (!/\/(comments|vmlDrawing)$/.test(type)) continue;
            const old = resolvePart(relSource(name), rel['@_Target']);
            if (!zip.file(old)) throw new Error('Missing comment/drawing part: ' + old);
            if (renames.has(old)) continue;
            let dest: string;
            do { dest = type.endsWith('/comments') ? 'xl/comments' + commentIndex++ + '.xml' : 'xl/drawings/vmlDrawing' + vmlIndex++ + '.vml'; }
            while (zip.file(dest));
            renames.set(old, dest);
        }
    }
    for (const [name, xml] of relDocs) {
        const source = relSource(name), newSource = renames.get(source) ?? source;
        for (const rel of relationships(xml)) {
            if (rel['@_TargetMode'] === 'External') continue;
            const old = resolvePart(source, rel['@_Target']);
            rel['@_Target'] = relativePart(newSource, renames.get(old) ?? old);
        }
        zip.file(name, builder.build(xml));
    }
    const types = zip.file('[Content_Types].xml');
    if (types) {
        const xml = parser.parse(await types.async('string'));
        for (const entry of list<any>(xml?.Types?.Override)) {
            const old = String(entry['@_PartName']).replace(/^\//, '');
            if (renames.has(old)) entry['@_PartName'] = '/' + renames.get(old);
        }
        zip.file('[Content_Types].xml', builder.build(xml));
    }
    for (const [old, dest] of renames) {
        const bytes = await zip.file(old)!.async('uint8array');
        zip.file(dest, bytes); zip.remove(old);
    }
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}
