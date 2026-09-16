import os, io, re, json, shutil, tempfile, zipfile, traceback
from pathlib import Path
import docx
import dtmix_legacy
from dtmix_legacy import DTMIXApp
# Không mở hộp thoại Tk trên server headless
dtmix_legacy.messagebox.showerror=lambda *a,**k: None
dtmix_legacy.messagebox.showwarning=lambda *a,**k: None
dtmix_legacy.messagebox.showinfo=lambda *a,**k: None

class V:
    def __init__(self, value=None): self.value=value; self._original_val=str(value) if value is not None else ''
    def get(self): return self.value
    def set(self,v): self.value=v
    def delete(self,*a): self.value=''
    def insert(self,*a): self.value=a[-1] if a else ''
    def select(self): self.value=True
    def deselect(self): self.value=False
    def configure(self,**kw): pass
    def winfo_exists(self): return True
    def pack(self,*a,**k): pass
    def pack_forget(self,*a,**k): pass
    def grid(self,*a,**k): pass

class DTMIXWebEngine:
    def __init__(self, filepath, output_root, info=None, youngmix=False, config=None):
        self.app = DTMIXApp.__new__(DTMIXApp)
        a=self.app; info=info or {}; config=config or {}
        a.filepath=str(filepath); a.output_dir=str(output_root)
        a.parsed_data={'header':[],'parts':[]}; a.errors=[]; a.youngmix_data=[]; a.youngmix_total_q=0
        a.muc_entries={}; a.muc_shuffle_vars={}; a.muc_fix_entries={}; a.muc_order_shuffle_vars={}; a.ym_ui_controls=[]
        a.youngmix_var=V(bool(youngmix)); a.master_fix_var=V(bool(config.get('master_fix', True)))
        a.continuous_numbering_var=V(bool(config.get('continuous_numbering', False))); a.keep_muc_var=V(bool(config.get('keep_group_titles', False)))
        a.auto_code_var=V(True); a.auto_code_var_ym=V(True); a.thu_cong_var=V(False); a.thu_cong_var_ym=V(False)
        codes=config.get('codes') or ['111','222','333','444']; codes=', '.join(map(str,codes))
        a.entry_made=V(codes); a.entry_made_ym=V(codes); a.combo_somade=V(str(len(config.get('codes') or [1,2,3,4]))); a.combo_somade_ym=V(a.combo_somade.get())
        defaults={'so':'SỞ GIÁO DỤC VÀ ĐÀO TẠO THÀNH PHỐ HỒ CHÍ MINH','truong':'TRƯỜNG THPT DƯƠNG BẠCH MAI','kythi':'KIỂM TRA HỌC KỲ II','namhoc':'NĂM HỌC 2025 - 2026','monthi':'Môn: HÓA HỌC','thoigian':'Thời gian làm bài: 45 phút'}
        vals={k:info.get(k,defaults[k]) for k in defaults}
        a.val_so=V(vals['so']); a.val_truong=V(vals['truong']); a.val_kythi=V(vals['kythi']); a.val_namhoc=V(vals['namhoc']); a.val_mon=V(vals['monthi']); a.val_thoigian=V(vals['thoigian'])
        a.after=lambda delay, fn=None: fn() if fn else None; a.update=lambda:None; a.update_idletasks=lambda:None
        a.update_progress=lambda *args,**kwargs:None; a.show_success_dialog=lambda d:setattr(a,'_web_out_dir',d)
        # tron_de may touch these only on error/finally
        a.btn_tron=V(); a.btn_tron_ym=V(); a.progress_bar=V(); a.progress_bar_ym=V(); a.lbl_status=V(); a.lbl_status_ym=V()
        self.config=config

    def load(self):
        a=self.app
        a.global_doc=docx.Document(a.filepath)
        a.convert_auto_numbering(a.global_doc); a.convert_floating_to_inline(a.global_doc); a.split_soft_returns(a.global_doc); a.preprocess_textboxes(a.global_doc); a.preprocess_layout_tables(a.global_doc)
        a.parsed_data,a.errors=a.parse_document_structure(a.global_doc)
        if a.youngmix_var.get(): a.parse_youngmix_structure(a.global_doc)
        self._make_controls()
        return self.summary()

    def _make_controls(self):
        a=self.app; cfg=self.config; std=cfg.get('standard_groups',{}); ym=cfg.get('youngmix_groups',[])
        for pi,p in enumerate(a.parsed_data.get('parts',[])):
            for mi,m in enumerate(p.get('mucs',[])):
                key=f'{pi}:{mi}'; c=std.get(key,{})
                default_shuffle=p.get('type')!=4
                a.muc_shuffle_vars[(pi,mi)]=V(c.get('shuffle_questions',default_shuffle))
                a.muc_order_shuffle_vars[(pi,mi)]=V(c.get('shuffle_group', default_shuffle and not m.get('is_fixed',False)))
                fixed=c.get('fixed_questions')
                if fixed is None: fixed=', '.join(str(i+1) for i,q in enumerate(m.get('questions',[])) if q.get('is_fixed'))
                a.muc_fix_entries[(pi,mi)]=V(fixed)
                a.muc_entries[(pi,mi)]=V(str(c.get('pick',len(m.get('questions',[])))))
        if a.youngmix_var.get():
            # mirror desktop auto-detection before creating controls
            all_m=[]
            for p in a.parsed_data.get('parts',[]): all_m.extend(p.get('mucs',[]))
            for i,g in enumerate(a.youngmix_data):
                if i < len(all_m):
                    m=all_m[i]; t1=t2=essay=short=0
                    for q in m.get('questions',[]):
                        qt=q.get('ym_type',4); n=len(q.get('answers',[]))
                        if n==0:
                            if qt==3: short+=1
                            else: essay+=1
                        elif n>=2:
                            if qt==2:t2+=1
                            else:t1+=1
                    if essay and not(t1 or t2 or short): g['q_type']='Tự luận'; g['mix_type']='Chỉ câu hỏi' if g['mix_type']=='Câu hỏi & đáp án' else g['mix_type']
                    elif t2 and not(t1 or essay or short): g['q_type']='TN 2025 - Phần 2'
                    elif short and not(t1 or t2 or essay): g['q_type']='TN 2025 - Phần 3'; g['mix_type']='Chỉ câu hỏi' if g['mix_type']=='Câu hỏi & đáp án' else g['mix_type']
            a.ym_ui_controls=[]
            for i,g in enumerate(a.youngmix_data):
                c=ym[i] if i<len(ym) else {}
                a.ym_ui_controls.append({'type':V(c.get('type',g.get('q_type','Trắc nghiệm'))),'mix':V(c.get('mix',g.get('mix_type','Câu hỏi & đáp án'))),'fix':V(c.get('fix', g.get('is_fixed',False) or a.master_fix_var.get())),'start_q1':V(c.get('start_q1', i==0 if a.continuous_numbering_var.get() else True)),'pick':V(str(c.get('pick',g.get('q_count',0))))})

    def summary(self):
        a=self.app; parts=[]; total=0; has_ans=0
        for pi,p in enumerate(a.parsed_data.get('parts',[])):
            groups=[]
            for mi,m in enumerate(p.get('mucs',[])):
                qs=m.get('questions',[]); total+=len(qs)
                for q in qs:
                    if any(x.get('is_true',False) for x in q.get('answers',[])) or q.get('ym_type')==3: has_ans+=1
                groups.append({'key':f'{pi}:{mi}','title':m.get('title') or f'Nhóm {mi+1}','count':len(qs),'fixed':m.get('is_fixed',False),'fixed_questions':[i+1 for i,q in enumerate(qs) if q.get('is_fixed')]})
            parts.append({'index':pi,'title':p.get('title',''),'type':p.get('type'),'groups':groups})
        y=[]
        for i,g in enumerate(a.youngmix_data): y.append({'index':i,'name':g.get('name'), 'tag':g.get('tag'),'count':g.get('q_count'),'type':g.get('q_type'),'mix':g.get('mix_type'),'fixed':g.get('is_fixed')})
        return {'parts':parts,'youngmix_groups':y,'total_questions':total,'recognized_answers':has_ans,'errors':[str(e) for e in a.errors]}

    def mix(self):
        self.app.tron_de()
        outdir=Path(getattr(self.app,'_web_out_dir',''))
        if not outdir.exists(): raise RuntimeError('DTMIX không tạo được thư mục kết quả.')
        zip_path=outdir.parent/(outdir.name+'.zip')
        with zipfile.ZipFile(zip_path,'w',zipfile.ZIP_DEFLATED) as z:
            for f in outdir.rglob('*'):
                if f.is_file(): z.write(f, f.relative_to(outdir))
        return zip_path
