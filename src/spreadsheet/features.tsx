import React,{useEffect,useRef,useState} from 'react';
import type Spreadsheet from '../react/view/excel/x-spreadsheet';
import {handler} from '../react/util/vscode';
import {getConfigs} from '../react/util/vscodeConfig';
import {serializeSpreadsheet} from '../react/view/excel/excel_writer';
import {valueSnapshot,invalidateEditedResults,overlayResults,summarize} from './editing';

export function useSpreadsheetFeatures() {
    const grid=useRef<Spreadsheet|null>(null),payload=useRef<any>({});
    const previous=useRef(new Map<string,string>()),revision=useRef(0);
    const [status,setStatus]=useState('Saved results');
    const [warnings,setWarnings]=useState<string[]>([]);
    const [comment,setComment]=useState('');
    const [calculating,setCalculating]=useState(false);
    function loaded(spreadsheet:Spreadsheet,open:any,readerWarnings?:string[]) {
        grid.current=spreadsheet;payload.current=open;revision.current=0;
        previous.current=valueSnapshot(spreadsheet.getData());
        setWarnings([...(open.info?.warnings ?? []),...(readerWarnings ?? [])]);
        setStatus('Saved results'+(open.info?.missingResults?' · '+open.info.missingResults+' results need calculation':''));
        spreadsheet.on('cell-selected',(cell:any)=>setComment(cell?.comment ?? ''));
        setTimeout(()=>handler.emit('previewReady',{
            warnings:readerWarnings,sheets:summarize(spreadsheet.getData()),canvasCount:document.querySelectorAll('canvas').length,
            toolbarCount:document.querySelectorAll('.x-spreadsheet-toolbar').length,
            formulaBarCount:document.querySelectorAll('.x-spreadsheet-formula-bar').length,
            toolbarButtons:document.querySelectorAll('.x-spreadsheet-toolbar-btn').length,
            editable:!open.readOnly,themeToggle:!!document.querySelector('.dark-mode-toggle')
        }),50);
    }
    function changed() {
        if(!grid.current)return;
        const sheets=grid.current.getData();
        if(invalidateEditedResults(sheets,previous.current,/^\.?(csv|tsv)$/i.test(payload.current.ext ?? ''))) {
            revision.current++;setStatus('Edited · recalculate to refresh formula results');
            previous.current=valueSnapshot(sheets);grid.current.reRender();
        }
        handler.emit('change',{revision:revision.current});
    }
    useEffect(()=>{
        handler.on('status',p=>{setStatus(p.value);setCalculating(p.value.startsWith('Calculating'));});
        handler.on('calculated',p=>{
            if(!grid.current || p.revision!==revision.current)return;
            const missing=overlayResults(grid.current.getData(),p.results);
            setCalculating(false);setStatus('Recalculated with '+p.backend+(missing?' · '+missing+' results unavailable':''));
            grid.current.reRender();handler.emit('calculationRendered',{sheets:summarize(grid.current.getData()),missing});
        });
        handler.on('exportRequest',async p=>{
            try {
                if(!grid.current)throw new Error('Workbook is still loading');
                const bytes=await serializeSpreadsheet(grid.current.getData(),p.ext ?? 'xlsx',payload.current.csvEncoding,payload.current.csvDelimiter);
                handler.emit('exportResponse',{id:p.id,buffer:Array.from(bytes),revision:revision.current});
            } catch(e){handler.emit('exportResponse',{id:p.id,error:String(e instanceof Error?e.message:e)});}
        });
        handler.on('recalculateRequest',()=>handler.emit('recalculate',{revision:revision.current,edited:revision.current>0}));
        // Only present in the official VS Code integration-test host.
        if((getConfigs() as any)?.testMode)handler.on('testAction',async p=>{
            try {
                if(!grid.current)throw new Error('No grid');
                const g:any=grid.current;
                if(p.action==='edit'){
                    grid.current.scrollToCell(p.ri,p.ci,p.sheetIndex ?? 0);
                    const input=document.querySelector('.x-spreadsheet-formula-bar-input') as HTMLInputElement;
                    input.focus();input.dispatchEvent(new FocusEvent('focus'));input.value=p.text;input.dispatchEvent(new Event('input',{bubbles:true}));input.blur();input.dispatchEvent(new FocusEvent('blur'));
                }
                else if(p.action==='undo')g.sheet.toolbar.change('undo');
                else if(p.action==='redo')g.sheet.toolbar.change('redo');
                else if(p.action==='bold'){grid.current.scrollToCell(p.ri,p.ci);g.sheet.toolbar.change('font-bold',true);}
                else if(p.action==='find')return handler.emit('testResponse',{id:p.id,matches:grid.current.findAll(p.text,{allSheets:true})});
                else if(p.action==='replace')grid.current.replaceAll(p.find,p.replace,{allSheets:true});
                else if(p.action==='theme')(document.querySelector('.dark-mode-toggle') as HTMLButtonElement)?.click();
                grid.current.reRender();
                handler.emit('testResponse',{id:p.id,sheets:summarize(grid.current.getData()),data:grid.current.getData(),
                    theme:document.body.classList.contains('office-adaptive'),toolbar:!!document.querySelector('.x-spreadsheet-toolbar')});
            }catch(e){handler.emit('testResponse',{id:p.id,error:String(e)});}
        });
    },[]);
    return {loaded,changed,error:(message:string)=>handler.emit('renderError',message),
        bar:<div className="ahmed-calculation-bar">
            <span role="status">{status}</span>
            <span className="ahmed-cell-comment" title={comment}>{comment && 'Comment: '+comment}</span>
            {!!warnings.length&&<details><summary>Workbook details ({warnings.length})</summary>{warnings.map((w,i)=><p key={i}>{w}</p>)}</details>}
            <button onClick={()=>handler.emit('recalculate',{revision:revision.current,edited:revision.current>0})}>Recalculate</button>
            {calculating&&<button onClick={()=>handler.emit('cancel')}>Cancel</button>}
            <label>Zoom <select aria-label="Zoom" defaultValue="1" onChange={e=>{const g:any=grid.current;if(g){g.data.setZoomScale(Number(e.target.value));g.reRender();}}}>
                {[0.5,0.75,1,1.25,1.5,2].map(z=><option key={z} value={z}>{z*100}%</option>)}
            </select></label>
            <button onClick={()=>handler.emit('openDesktop')}>Open in desktop app</button>
        </div>};
}
