import ExcelJS from '@cweijan/exceljs';
import JSZip from 'jszip';
export async function fixture(options: {unusualComments?:boolean; missing?:boolean} = {}): Promise<Uint8Array> {
    const book = new ExcelJS.Workbook();
    const sheet = book.addWorksheet('QUICK_SPLITS');
    const inputs = book.addWorksheet('Inputs');
    inputs.getCell('A1').value = 1;
    const secret = book.addWorksheet('Hidden notes',{state:'hidden'});
    secret.getCell('A1').value = 'Hidden content';
    sheet.getCell('A1').value='Heat allocation';sheet.mergeCells('A1:C1');
    sheet.getColumn('A').width=42;
    sheet.getCell('A1').font={bold:true,size:14};
    sheet.getCell('B5').value=.4;sheet.getCell('C5').value=.15;sheet.getCell('D5').value=.3;sheet.getCell('E5').value=0;sheet.getCell('F5').value=.15;
    sheet.getCell('G5').value={formula:'SUM(B5:F5)',result:1};
    sheet.getCell('G5').numFmt='0%';
    sheet.getCell('H5').value={formula:'IF(ABS(G5-1)<0.001,"OK","FAIL")',...!options.missing?{result:'OK'}:{}};
    sheet.getCell('H5').note='Verify percentages sum to 100%.';
    sheet.getCell('A8').value={formula:'0',result:0};
    sheet.getCell('A9').value={formula:'FALSE()',result:false};
    sheet.getCell('A10').value={formula:'""',result:''};
    sheet.getCell('A11').value={formula:'NA()',result:{error:'#N/A'}};
    sheet.getCell('A12').value='=this is literal text';
    sheet.getCell('A13').value={formula:"Inputs!A1*2",result:2};
    sheet.getCell('A14').value={formula:'IFERROR(INDEX(Inputs!A1:A2,MATCH(1,Inputs!A1:A2,0)),0)',result:1};
    sheet.getCell('A15').value=new Date('2024-01-01T00:00:00Z');
    sheet.getCell('A15').numFmt='yyyy-mm-dd';
    const bytes = new Uint8Array(await book.xlsx.writeBuffer());
    if(!options.unusualComments)return bytes;
    const zip=await JSZip.loadAsync(bytes);
    // Simulate openpyxl and nonstandard metadata locations, with absolute OPC targets.
    const changes=new Map<string,string>();
    for(const name of Object.keys(zip.files)) {
        if(/^xl\/comments\d+\.xml$/.test(name))changes.set(name,'xl/notes/custom-note.xml');
        if(/^xl\/drawings\/vmlDrawing\d+\.vml$/.test(name))changes.set(name,'xl/notes/custom-drawing.vml');
    }
    for(const [old,dest] of changes) {zip.file(dest,await zip.file(old)!.async('uint8array'));zip.remove(old);}
    for(const name of Object.keys(zip.files)) {
        if(!name.endsWith('.rels')&&name!=='[Content_Types].xml')continue;
        let data=await zip.file(name)!.async('string');
        for(const [old,dest] of changes) {
            data=data.replaceAll('/'+old,'/'+dest).replaceAll('../'+old.replace(/^xl\//,''),'/'+dest);
        }
        zip.file(name,data);
    }
    return zip.generateAsync({type:'uint8array'});
}
