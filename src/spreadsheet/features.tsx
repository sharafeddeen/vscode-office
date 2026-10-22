import React,{useEffect,useRef,useState} from 'react';
import type Spreadsheet from '../react/view/excel/x-spreadsheet';
import {handler} from '../react/util/vscode';
import {getConfigs} from '../react/util/vscodeConfig';
import {serializeSpreadsheet} from '../react/view/excel/excel_writer';
import {loadSheets} from '../react/view/excel/excel_reader';
import {valueSnapshot,invalidateEditedResults,overlayResults,summarize} from './editing';
import {applyNativeResults,markDerived} from './native-model';

export function useSpreadsheetFeatures() {
    const grid=useRef<Spreadsheet|null>(null),payload=useRef<any>({});
    const previous=useRef(new Map<string,string>()),revision=useRef(0);
    const timer=useRef<ReturnType<typeof setTimeout>>();
    const busy=useRef(false),nativeDone=useRef<{resolve:()=>void;reject:(e:Error)=>void}>();
    const [status,setStatus]=useState('Saved results'),[warnings,setWarnings]=useState<string[]>([]);
    const [comment,setComment]=useState(''),[calculating,setCalculating]=useState(false);
    const [blocking,setBlocking]=useState(false),[error,setError]=useState('');
    const enabled=()=>!!payload.current.nativeEditing;
    function history(state:any){
        if(!state||!grid.current)return;
        const g:any=grid.current;
        for(const data of g.datas)data.nativeHistory={undo:state.undo,redo:state.redo};
        g.sheet.toolbar.reset();
    }
    function model(){return grid.current?.getData() ?? [];}
    function installNativeCommands(){
        grid.current?.setNativeCommandHandler(enabled()?command=>{void nativeCommand(command).catch(()=>{});}:undefined);
        if(enabled())history({undo:false,redo:false});
    }
    function loaded(spreadsheet:Spreadsheet,open:any,readerWarnings?:string[]) {
        clearTimeout(timer.current);grid.current=spreadsheet;payload.current=open;revision.current=0;
        markDerived(spreadsheet.getData(),open.info?.arrays);
        previous.current=valueSnapshot(spreadsheet.getData());installNativeCommands();
        setWarnings([...(open.info?.warnings ?? []),...(readerWarnings ?? []),...(open.nativeWarning?[open.nativeWarning]:[])]);
        setStatus('Saved results'+(open.info?.missingResults?' · '+open.info.missingResults+' results need calculation':''));
        setError('');spreadsheet.on('cell-selected',(cell:any)=>setComment(cell?.comment ?? ''));
        setTimeout(()=>handler.emit('previewReady',{
            model:model(),warnings:readerWarnings,sheets:summarize(spreadsheet.getData()),canvasCount:document.querySelectorAll('canvas').length,
            toolbarCount:document.querySelectorAll('.x-spreadsheet-toolbar').length,
            formulaBarCount:document.querySelectorAll('.x-spreadsheet-formula-bar').length,
            toolbarButtons:document.querySelectorAll('.x-spreadsheet-toolbar-btn').length,
            editable:!open.readOnly,nativeEditing:open.nativeEditing,themeToggle:!!document.querySelector('.dark-mode-toggle')
        }),50);
    }
    function changed() {
        if(!grid.current||busy.current)return;
        const sheets=grid.current.getData();
        if(invalidateEditedResults(sheets,previous.current,/^\.?(csv|tsv)$/i.test(payload.current.ext ?? ''))) {
            setStatus('Edited · waiting for native calculation');
            previous.current=valueSnapshot(sheets);grid.current.reRender();
        }
        revision.current++;handler.emit('change',{revision:revision.current});
        if(enabled()){
            clearTimeout(timer.current);
            timer.current=setTimeout(()=>handler.emit('nativeSync',{model:model(),revision:revision.current}),180);
        }
    }
    function nativeCommand(command:any):Promise<void> {
        if(!enabled()||busy.current)return Promise.resolve();
        clearTimeout(timer.current);busy.current=true;setBlocking(true);setError('');
        revision.current++;
        const completion=new Promise<void>((resolve,reject)=>{nativeDone.current={resolve,reject};});
        handler.emit('nativeCommand',{command,model:model(),revision:revision.current,selection:grid.current?.getSelection()});
        return completion;
    }
    useEffect(()=>{
        const blockInput=(event:Event)=>{if(busy.current){event.preventDefault();event.stopImmediatePropagation();}};
        for(const type of ['keydown','beforeinput','paste'])window.addEventListener(type,blockInput,true);
        handler.on('status',p=>{setStatus(p.value);setCalculating(p.value.startsWith('Calculating'));});
        handler.on('calculated',async p=>{
            if(!grid.current||p.revision!==revision.current)return;
            const projection=p.native&&p.buffer?await loadSheets(new Uint8Array(p.buffer).buffer,'.xlsx'):undefined;
            if(p.revision!==revision.current)return;
            const sheets=grid.current.getData();
            const missing=p.native?applyNativeResults(sheets,p.results,p.arrays,projection?.sheets):overlayResults(sheets,p.results);
            previous.current=valueSnapshot(sheets);history(p.history);
            setCalculating(false);setError('');
            setStatus('Recalculated with '+p.backend+(missing?' · '+missing+' results unavailable':''));
            grid.current.reRender();
            if(p.native)handler.emit('nativeProjectionAccepted',{model:model(),revision:revision.current});
            handler.emit('calculationRendered',{sheets:summarize(sheets),missing,revision:p.revision});
        });
        handler.on('modelRequest',p=>{
            clearTimeout(timer.current);
            handler.emit('modelResponse',{id:p.id,model:model(),revision:revision.current});
        });
        handler.on('exportRequest',async p=>{
            try {
                if(!grid.current)throw new Error('Workbook is still loading');
                const bytes=await serializeSpreadsheet(grid.current.getData(),p.ext ?? 'xlsx',payload.current.csvEncoding,payload.current.csvDelimiter);
                handler.emit('exportResponse',{id:p.id,buffer:Array.from(bytes),revision:revision.current});
            }catch(e){handler.emit('exportResponse',{id:p.id,error:String(e instanceof Error?e.message:e)});}
        });
        handler.on('nativeSynced',p=>{if(p.revision===revision.current)history(p.history);});
        handler.on('nativeSnapshot',async p=>{
            try{
                const data=await loadSheets(new Uint8Array(p.buffer).buffer,'.xlsx');
                markDerived(data.sheets,p.info?.arrays);
                grid.current!.loadData(data.sheets);
                payload.current.info=p.info;
                previous.current=valueSnapshot(model());
                const selection=p.selection;
                if(selection)grid.current!.scrollToCell(selection.ri,selection.ci,Math.min(selection.sheetIndex,data.sheets.length-1));
                history(p.history);
                handler.emit('modelResponse',{id:p.id,model:model(),revision:revision.current});
            }catch(e){handler.emit('modelResponse',{id:p.id,error:String(e)});setError(String(e));busy.current=false;setBlocking(false);nativeDone.current?.reject(e as Error);nativeDone.current=undefined;}
        });
        handler.on('nativeReady',p=>{
            history(p.history);busy.current=false;setBlocking(false);setError('');
            setStatus('LibreOffice UNO · edits applied');
            handler.emit('nativeRendered',{sheets:summarize(model()),revision:revision.current});
            nativeDone.current?.resolve();nativeDone.current=undefined;
        });
        handler.on('nativeFailure',message=>{
            setCalculating(false);setError(message);busy.current=false;setBlocking(false);
            nativeDone.current?.reject(new Error(message));nativeDone.current=undefined;
        });
        handler.on('recalculateRequest',()=>{clearTimeout(timer.current);handler.emit('recalculate',{revision:revision.current});});
        if((getConfigs() as any)?.testMode)handler.on('testAction',async p=>{
            try {
                if(!grid.current)throw new Error('No grid');
                const g:any=grid.current;
                if(p.action==='edit'){
                    grid.current.scrollToCell(p.ri,p.ci,p.sheetIndex ?? 0);
                    const input=document.querySelector('.x-spreadsheet-formula-bar-input') as HTMLInputElement;
                    input.focus();input.dispatchEvent(new FocusEvent('focus'));input.value=p.text;input.dispatchEvent(new Event('input',{bubbles:true}));input.blur();input.dispatchEvent(new FocusEvent('blur'));
                }else if(p.action==='undo'||p.action==='redo'){
                    if(enabled())await nativeCommand({action:p.action});else g.sheet.toolbar.change(p.action);
                }else if(p.action==='nativeOperation')await nativeCommand(p.command);
                else if(p.action==='bold'){grid.current.scrollToCell(p.ri,p.ci);g.sheet.toolbar.change('font-bold',true);}
                else if(p.action==='find')return handler.emit('testResponse',{id:p.id,matches:grid.current.findAll(p.text,{allSheets:true})});
                else if(p.action==='replace')grid.current.replaceAll(p.find,p.replace,{allSheets:true});
                else if(p.action==='theme')(document.querySelector('.dark-mode-toggle') as HTMLButtonElement)?.click();
                grid.current.reRender();
                handler.emit('testResponse',{id:p.id,sheets:summarize(model()),data:model(),revision:revision.current,
                    theme:document.body.classList.contains('office-adaptive'),toolbar:!!document.querySelector('.x-spreadsheet-toolbar')});
            }catch(e){handler.emit('testResponse',{id:p.id,error:String(e)});}
        });
        return()=>{clearTimeout(timer.current);for(const type of ['keydown','beforeinput','paste'])window.removeEventListener(type,blockInput,true);};
    },[]);
    return {loaded,changed,native:()=>!!payload.current.nativeAvailable,
        error:(message:string)=>handler.emit('renderError',message),
        bar:<>
            {blocking&&<div className="ahmed-native-blocker" role="status">Applying edit in LibreOffice…</div>}
            <div className="ahmed-calculation-bar">
                <span role="status" title={error || status}>{error || status}</span>
                <span className="ahmed-cell-comment" title={comment}>{comment && 'Comment: '+comment}</span>
                {!!warnings.length&&<details><summary>Workbook details ({warnings.length})</summary>{warnings.map((w,i)=><p key={i}>{w}</p>)}</details>}
                {error&&enabled()&&<button onClick={()=>void nativeCommand({action:'restore'}).catch(()=>{})}>Reload native state</button>}
                <button disabled={blocking} onClick={()=>{clearTimeout(timer.current);handler.emit('recalculate',{revision:revision.current});}}>Recalculate</button>
                {calculating&&<button onClick={()=>handler.emit('cancel')}>Cancel</button>}
                <label>Zoom <select aria-label="Zoom" defaultValue="1" onChange={e=>{const g:any=grid.current;if(g){g.data.setZoomScale(Number(e.target.value));g.reRender();}}}>
                    {[0.5,0.75,1,1.25,1.5,2].map(z=><option key={z} value={z}>{z*100}%</option>)}
                </select></label>
                <button onClick={()=>handler.emit('openDesktop')}>Open in desktop app</button>
            </div>
        </>};
}
