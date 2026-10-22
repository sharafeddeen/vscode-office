import '../react/polyfills/buffer';
import React, {useEffect, useRef, useState} from 'react';
import {createRoot} from 'react-dom/client';
import Spreadsheet, {type SheetData} from '../react/view/excel/x-spreadsheet';
import {loadSheets} from '../react/view/excel/excel_reader';
import {displayCellText} from '../react/view/excel/excel_cell_display';
import {findAllInSheets} from '../react/view/excel/excel_find';
import {handler} from '../react/util/vscode';
import {initExcelLocale} from '../react/view/excel/excel_i18n';
import '../react/view/excel/Excel.less';
import './preview.css';
import type {Results} from './results';

initExcelLocale();
const api = window as any;
function App() {
    const container = useRef<HTMLDivElement>(null);
    const grid = useRef<Spreadsheet>();
    const sheets = useRef<SheetData[]>([]);
    const info = useRef<any>();
    const displayed = useRef<SheetData[]>([]);
    const mode = useRef('Saved results');
    const [status,setStatus] = useState('Loading workbook…');
    const [warnings,setWarnings] = useState<string[]>([]);
    const [error,setError] = useState('');
    const [query,setQuery] = useState('');
    const [hidden,setHidden] = useState(false);
    const [selection,setSelection] = useState<any>(null);
    const matchIndex = useRef(0);
    const currentQuery = useRef('');
    const documentKey = useRef('');
    function cellText(sheet:SheetData,ri:number,ci:number) {
        const cell = (sheet.rows?.[ri] as any)?.cells?.[ci];
        if(!cell) return '';
        return displayCellText(cell,sheet.styles?.[cell.style] ?? {},(x,y)=>(sheet.rows?.[y] as any)?.cells?.[x]);
    }
    function summary() {
        return sheets.current.map(sheet=>({name:sheet.name,cells:Object.fromEntries(Object.entries(sheet.rows ?? {}).filter(([key])=>key!=='len').flatMap(([ri,row]:any)=>Object.keys(row.cells ?? {}).filter(ci=>row.cells[ci].text).slice(0,50).map(ci=>[ri+':'+ci,cellText(sheet,Number(ri),Number(ci))])).slice(0,2000))}));
    }
    function render(showHidden=hidden) {
        if(!container.current) return;
        const visible = sheets.current.filter((_,i)=>showHidden || (info.current?.sheets[i]?.state ?? 'visible')==='visible');
        displayed.current = visible;
        if(!visible.length) {setError('No visible sheets. Enable Show hidden sheets.'); return;}
        const spreadsheet = grid.current ?? new Spreadsheet(container.current,{
            mode:'read',showToolbar:false,showContextmenu:false,
            view:{height:()=>window.innerHeight-138,width:()=>window.innerWidth},
            row:{len:Math.max(40,...visible.map(s=>s.rows?.len ?? 0)),height:30},
            col:{len:Math.max(26,...visible.map(s=>s.cols?.len ?? 0))}
        });
        spreadsheet.loadData(visible);
        grid.current = spreadsheet;
        const select = (cell:any,ri:number,ci:number) => {
            setSelection({cell,ri,ci,sheet:displayed.current[spreadsheet.getActiveSheetIndex()]});
            try {localStorage.setItem(documentKey.current,JSON.stringify(spreadsheet.getSelection()));} catch {}
        };
        if (!(spreadsheet as any).ahmedBound) {
            (spreadsheet as any).ahmedBound = true;
            spreadsheet.on('cell-selected',select);
            spreadsheet.on('find',()=>document.getElementById('find')?.focus());
        }
        try {
            const saved=JSON.parse(localStorage.getItem(documentKey.current) ?? 'null');
            if(saved) spreadsheet.scrollToCell(saved.ri,saved.ci,Math.min(saved.sheetIndex,visible.length-1));
        } catch {}
        requestAnimationFrame(()=>handler.emit('previewReady',{warnings:info.current?.readerWarnings,sheets:summary(),canvasCount:container.current?.querySelectorAll('canvas').length ?? 0}));
    }
    useEffect(()=>{
        handler.on('open',async payload=>{
            try {
                setError('');setStatus('Loading workbook…');info.current=payload.info;
                documentKey.current='ahmed-spreadsheet:'+payload.documentCacheId;
                const bytes=new Uint8Array(payload.buffer);
                const data=await loadSheets(bytes.buffer,payload.ext);
                sheets.current=data.sheets;
                info.current = {...info.current,readerWarnings:data.warnings};
                mode.current='Saved results';
                setWarnings([...(payload.info?.warnings ?? []),...(data.warnings ?? [])]);
                setStatus('Saved results'+(payload.info?.missingResults?' · '+payload.info.missingResults+' results need calculation':''));
                render(false);setHidden(false);
            } catch(e) {const message=String(e instanceof Error?e.message:e);setError(message);setStatus('Failed to load workbook');handler.emit('renderError',message);}
        });
        handler.on('loadError',message=>{setError(message);setStatus('Failed to load workbook');handler.emit('renderError',message);});
        handler.on('status',payload=>{setStatus(payload.value);if(payload.warnings)setWarnings(previous=>Array.from(new Set([...previous,...payload.warnings])));});
        handler.on('calculated',payload=>{
            const results:Results=payload.results;
            let missing=0;
            for(const sheet of sheets.current) {
                for(const [ri,row] of Object.entries(sheet.rows ?? {})) {
                    if(ri==='len')continue;
                    for(const [ci,cell] of Object.entries((row as any).cells ?? {}) as any) {
                        if(!cell.formulaPreview)continue;
                        let n=Number(ci)+1,col='';while(n>0){n--;col=String.fromCharCode(65+n%26)+col;n=Math.floor(n/26);}
                        const result=results[sheet.name!]?.[col+(Number(ri)+1)];
                        if(result) {cell.savedValue=result.value;cell.hasSavedValue=true;cell.savedType=result.type;}
                        else missing++;
                    }
                }
            }
            mode.current='Recalculated with '+payload.backend;
            setStatus(mode.current+(missing?' · '+missing+' results unavailable':''));
            grid.current?.reRender();
            setSelection(previous=>previous?{...previous}:null);
            handler.emit('calculationRendered',{sheets:summary(),missing});
        });
        handler.emit('init');
        const resize=()=>grid.current?.reRender();
        const keyboard=(event:KeyboardEvent)=>{if((event.ctrlKey||event.metaKey)&&event.key.toLowerCase()==='f'){event.preventDefault();document.getElementById('find')?.focus();}};
        window.addEventListener('resize',resize);window.addEventListener('keydown',keyboard);
        return ()=>{window.removeEventListener('resize',resize);window.removeEventListener('keydown',keyboard);};
    },[]);
    function findNext() {
        if(!grid.current||!query)return;
        const matches=findAllInSheets(grid.current.getData(),query,{allSheets:true});
        if(currentQuery.current!==query){currentQuery.current=query;matchIndex.current=0;}
        if(!matches.length){setStatus(mode.current+' · no matches');return;}
        const match=matches[matchIndex.current++ % matches.length];grid.current.scrollToCell(match.ri,match.ci,match.sheetIndex);
        setStatus(mode.current+' · '+matches.length+' matches');
    }
    const cell=selection?.cell;
    return <main className="excel-viewer ahmed-preview">
        <header><strong>Ahmed Spreadsheet Viewer</strong><span>Read only</span>
            <input id="find" aria-label="Find in workbook" placeholder="Find text or formula" value={query} onChange={e=>setQuery(e.target.value)} onKeyDown={e=>{if(e.key==='Enter')findNext();}}/>
            <button onClick={findNext}>Find next</button>
            <button onClick={()=>handler.emit('recalculate')}>Recalculate</button>
            <button onClick={()=>handler.emit('cancel')}>Cancel</button>
            <button onClick={()=>handler.emit('openDesktop')}>Open in desktop app</button>
        </header>
        <section className="preview-status" role="status">{status}
            {!!info.current?.sheets.some((s:any)=>s.state!=='visible')&&<label><input type="checkbox" checked={hidden} onChange={e=>{setHidden(e.target.checked);render(e.target.checked);}}/>Show hidden sheets</label>}
            <label>Zoom <select aria-label="Zoom" defaultValue="1" onChange={e=>{const g:any=grid.current;if(g){g.data.setZoomScale(Number(e.target.value));g.reRender();}}}>{[0.5,0.75,1,1.25,1.5,2].map(z=><option key={z} value={z}>{z*100}%</option>)}</select></label>
        </section>
        {!!warnings.length&&<details className="preview-warnings"><summary>Workbook details ({warnings.length})</summary>{warnings.map((w,i)=><p key={i}>{w}</p>)}</details>}
        <section className="preview-formula"><span>{selection?selection.sheet?.name+' · row '+(selection.ri+1):'Select a cell'}</span><input aria-label="Formula or value" readOnly value={cell?.text ?? ''}/><span title={cell?.comment ?? ''}>{cell?.comment?'Comment: '+cell.comment:''}</span></section>
        {error&&<div role="alert" className="preview-error"><h2>Unable to display workbook</h2><p>{error}</p><button onClick={()=>handler.emit('openDesktop')}>Open in desktop app</button></div>}
        <div ref={container} className="preview-grid" style={{display:error?'none':undefined}}/>
    </main>;
}
createRoot(document.getElementById('root')!).render(<App/>);
