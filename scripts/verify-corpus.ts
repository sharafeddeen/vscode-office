import {promises as fs} from 'node:fs';
import {join,extname} from 'node:path';
import {createHash} from 'node:crypto';
import {inspectWorkbook} from '../src/spreadsheet/package';
import {loadSheets} from '../src/react/view/excel/excel_reader';
import {displayCellText} from '../src/react/view/excel/excel_cell_display';
async function main() {
const root=process.argv[2];
if(!root)throw new Error('Usage: npm run test:corpus -- /external/evidence/directory');
async function* files(dir:string):AsyncGenerator<string>{for(const entry of await fs.readdir(dir,{withFileTypes:true})){const path=join(dir,entry.name);if(entry.isDirectory())yield*files(path);else if(extname(path).toLowerCase()==='.xlsx')yield path;}}
let count=0,sheets=0,formulas=0,missing=0,fallbacks=0;const failures:any[]=[];
for await(const path of files(root)){
    try{
        const bytes=new Uint8Array(await fs.readFile(path)),hash=createHash('sha256').update(bytes).digest('hex');
        const info=await inspectWorkbook(bytes);
        const data=await loadSheets(bytes.buffer as ArrayBuffer,'.xlsx');
        if(data.sheets.length!==info.sheets.length||data.sheets.some((s,i)=>s.name!==info.sheets[i].name))throw new Error('Sheet names/order mismatch: '+JSON.stringify({expected:info.sheets.map(s=>s.name),actual:data.sheets.map(s=>s.name),warnings:data.warnings}));
        for(const s of data.sheets)for(const row of Object.values(s.rows??{}))if(typeof row==='object')for(const cell of Object.values(row?.cells??{})){
            if(cell.formulaPreview){formulas++;if(!cell.hasSavedValue)missing++;
                const value=displayCellText(cell,s.styles?.[cell.style!]??{},()=>null);
                if(cell.hasSavedValue&&value==='Not calculated')throw new Error('Saved result disappeared');
            }
        }
        if(createHash('sha256').update(await fs.readFile(path)).digest('hex')!==hash)throw new Error('Source hash changed');
        fallbacks+=data.warnings?.some(w=>w.startsWith('Fallback reader'))?1:0;
        count++;sheets+=data.sheets.length;
        if(count%25===0)console.log(JSON.stringify({count,sheets,formulas,missing,fallbacks}));
    }catch(e){failures.push({path,message:String(e)});}
}
const report={count,sheets,formulas,missing,fallbacks,failures};
await fs.mkdir('.cache',{recursive:true});await fs.writeFile('.cache/corpus-report.json',JSON.stringify(report,null,2));
console.log(JSON.stringify({...report,failures:failures.length}));
if(failures.length)process.exitCode=1;

}
main().catch(e=>{console.error(e);process.exitCode=1;});
