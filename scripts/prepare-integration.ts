import {fixture} from '../test/spreadsheet/fixture';
import {promises as fs} from 'node:fs';
import {join} from 'node:path';
async function main(){
    const root=process.argv[2];await fs.mkdir(root,{recursive:true});
    await fs.writeFile(join(root,'literal.csv'),'=ABS(-1),label');
    await fs.writeFile(join(root,'saved.xlsx'),await fixture({unusualComments:true}));
    await fs.writeFile(join(root,'missing.xlsx'),await fixture({missing:true,unusualComments:true}));
}
main().catch(e=>{console.error(e);process.exitCode=1;});
