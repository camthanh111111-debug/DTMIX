from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from pathlib import Path
from uuid import uuid4
import json, shutil
from engine import DTMIXWebEngine

BASE=Path(__file__).resolve().parent; UP=BASE/'uploads'; OUT=BASE/'outputs'; UP.mkdir(exist_ok=True); OUT.mkdir(exist_ok=True)
app=FastAPI(title='DTMIX Web 1.3')
app.mount('/static',StaticFiles(directory=BASE/'static'),name='static')

@app.get('/',response_class=HTMLResponse)
def home(): return (BASE/'templates'/'index.html').read_text(encoding='utf-8')

def save_upload(f):
    if not f.filename.lower().endswith('.docx'): raise HTTPException(400,'Chỉ chấp nhận file .docx')
    sid=uuid4().hex; p=UP/f'{sid}.docx'
    with p.open('wb') as o: shutil.copyfileobj(f.file,o)
    return sid,p

@app.post('/api/analyze')
def analyze(file:UploadFile=File(...), youngmix:bool=Form(False)):
    sid,p=save_upload(file)
    try:
        e=DTMIXWebEngine(p,OUT,youngmix=youngmix); data=e.load(); data['session_id']=sid; return data
    except Exception as ex: raise HTTPException(500,f'Lỗi đọc đề: {ex}')

@app.post('/api/mix')
def mix(session_id:str=Form(...), payload:str=Form(...)):
    p=UP/f'{session_id}.docx'
    if not p.exists(): raise HTTPException(404,'Phiên tải đề không còn tồn tại.')
    try: cfg=json.loads(payload); e=DTMIXWebEngine(p,OUT,info=cfg.get('info'),youngmix=cfg.get('youngmix',False),config=cfg); e.load(); z=e.mix(); return FileResponse(z,media_type='application/zip',filename=z.name)
    except Exception as ex: raise HTTPException(500,f'Lỗi trộn đề: {ex}')

@app.get('/health')
def health(): return {'status':'ok','app':'DTMIX Web 1.3'}
