from __future__ import annotations
import os
import queue
import sys
import threading
import traceback
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog,messagebox,ttk

from .service import run_batch

class Application:
    def __init__(self,root):
        self.root=root
        self.pdfs=[]
        self.events=queue.Queue()
        self.cancel=threading.Event()
        self.running=False
        self.root.title('HYBE PDF 자동 입력 v1.3.1')
        self.root.geometry('980x730')
        self.root.minsize(780,620)
        self.root.protocol('WM_DELETE_WINDOW',self.close)
        base=Path(sys.executable).parent if getattr(sys,'frozen',False) else Path(__file__).resolve().parents[2]
        self.template=tk.StringVar(value=str(base/'Template.xlsx') if (base/'Template.xlsx').exists() else '')
        self.reference=tk.StringVar()
        self.reviewed=tk.StringVar()
        self.output=tk.StringVar(value=str(Path.home()/'Downloads'/f'HYBE_결과_{datetime.now():%Y%m%d_%H%M%S}.xlsx'))
        self.status=tk.StringVar(value='PDF와 양식을 선택하세요. 참고 엑셀은 선택 사항입니다.')
        frame=ttk.Frame(root,padding=18)
        frame.pack(fill='both',expand=True)
        frame.columnconfigure(1,weight=1)
        ttk.Label(frame,text='HYBE PDF 자동 입력',font=('맑은 고딕',19,'bold')).grid(row=0,column=0,columnspan=3,sticky='w')
        ttk.Label(frame,text='원문 근거가 있는 값만 입력 · 판단되지 않는 값은 빈칸',font=('맑은 고딕',10)).grid(row=1,column=0,columnspan=3,sticky='w',pady=(3,16))
        self.controls=[]
        self.pdf_label=ttk.Label(frame,text='선택한 PDF 없음')
        self.pdf_label.grid(row=2,column=1,sticky='w',padx=8)
        self.button(frame,'PDF 선택',self.pick_pdfs,2,0)
        self.button(frame,'폴더 선택',self.pick_folder,2,2)
        for row,label,variable,command in [(3,'양식 엑셀',self.template,lambda:self.pick_excel(self.template)),
                (4,'비교할 엑셀 (선택)',self.reference,lambda:self.pick_excel(self.reference)),
                (5,'직접 검토확정 파일',self.reviewed,self.pick_reviewed),
                (6,'결과 저장',self.output,self.pick_output)]:
            ttk.Label(frame,text=label).grid(row=row,column=0,sticky='w',pady=7)
            entry=ttk.Entry(frame,textvariable=variable)
            entry.grid(row=row,column=1,sticky='ew',padx=8)
            self.controls.append(entry)
            self.button(frame,'찾아보기',command,row,2)
        ttk.Label(frame,text='엑셀은 차이 비교에만 사용합니다. 스캔 인보이스 금액·수량은 원문을 직접 대조한 검토확정 파일이 있어야 입력됩니다.',wraplength=870).grid(row=7,column=0,columnspan=3,sticky='w',pady=(6,13))
        bar=ttk.Frame(frame)
        bar.grid(row=8,column=0,columnspan=3,sticky='ew')
        self.start_button=ttk.Button(bar,text='검증 후 엑셀 생성',command=self.start)
        self.start_button.pack(side='left')
        self.controls.append(self.start_button)
        self.stop_button=ttk.Button(bar,text='중지',command=self.cancel.set,state='disabled')
        self.stop_button.pack(side='left',padx=10)
        self.open_button=ttk.Button(bar,text='결과 폴더 열기',command=self.open_output)
        self.open_button.pack(side='right')
        self.progress=ttk.Progressbar(frame,mode='indeterminate')
        self.progress.grid(row=9,column=0,columnspan=3,sticky='ew',pady=12)
        ttk.Label(frame,textvariable=self.status,wraplength=870).grid(row=10,column=0,columnspan=3,sticky='w')
        self.text=tk.Text(frame,wrap='word',font=('맑은 고딕',10),state='disabled')
        self.text.grid(row=11,column=0,columnspan=3,sticky='nsew',pady=(10,0))
        frame.rowconfigure(11,weight=1)
        self.root.after(100,self.poll)

    def button(self,parent,label,command,row,col):
        button=ttk.Button(parent,text=label,command=command)
        button.grid(row=row,column=col,sticky='ew',pady=5)
        self.controls.append(button)

    def pick_pdfs(self):
        values=filedialog.askopenfilenames(title='처리할 PDF 선택',filetypes=[('PDF','*.pdf')])
        if values:
            self.pdfs=list(values)
            self.pdf_label.config(text=f'{len(values)}개 PDF 선택')

    def pick_folder(self):
        folder=filedialog.askdirectory(title='PDF 폴더 선택')
        if folder:
            self.pdfs=[str(p) for p in sorted(Path(folder).rglob('*.pdf'))]
            self.pdf_label.config(text=f'{len(self.pdfs)}개 PDF 선택')

    def pick_excel(self,variable):
        path=filedialog.askopenfilename(filetypes=[('Excel','*.xlsx *.xlsm')])
        if path:
            variable.set(path)

    def pick_output(self):
        extension=Path(self.template.get()).suffix or '.xlsx'
        path=filedialog.asksaveasfilename(defaultextension=extension,filetypes=[('Excel','*'+extension)])
        if path:
            self.output.set(path)

    def pick_reviewed(self):
        path=filedialog.askopenfilename(title='원문 직접 검토확정 JSON',filetypes=[('검토확정 파일','*.json')])
        if path:
            self.reviewed.set(path)

    def append(self,message):
        self.status.set(message)
        self.text.config(state='normal')
        self.text.insert('end',message+'\n')
        self.text.see('end')
        self.text.config(state='disabled')

    def start(self):
        if not self.pdfs or not self.template.get().strip() or not self.output.get().strip():
            messagebox.showerror('입력 확인','PDF, 양식 엑셀, 결과 저장 경로가 필요합니다.')
            return
        self.running=True
        self.cancel.clear()
        for control in self.controls:
            control.configure(state='disabled')
        self.stop_button.config(state='normal')
        self.progress.start(12)
        arguments=(list(self.pdfs),self.template.get(),self.output.get(),self.reference.get().strip() or None,self.reviewed.get().strip() or None)
        threading.Thread(target=self.worker,args=arguments,daemon=True).start()

    def worker(self,pdfs,template,output,reference,reviewed):
        try:
            result=run_batch(pdfs,template,output,reference_path=reference,reviewed_path=reviewed,
                cancel=self.cancel,progress=lambda s:self.events.put(('progress',s)))
            self.events.put(('done',result))
        except Exception as exc:
            self.events.put(('error',(str(exc),traceback.format_exc())))

    def poll(self):
        try:
            while True:
                kind,value=self.events.get_nowait()
                if kind=='progress':
                    self.append(value)
                    continue
                self.running=False
                self.progress.stop()
                for control in self.controls:
                    control.configure(state='normal')
                self.stop_button.config(state='disabled')
                if kind=='error':
                    self.append(value[1])
                    messagebox.showerror('처리 실패',value[0])
                else:
                    counts=value['counts']
                    message=f'저장 {value["written"]}건 / 확인 필요 {counts["review"]}건 / 실패 {counts["error"]}건 / 중복 {counts["skipped_duplicate"]}건'
                    if value['cancelled']:
                        message='사용자 중지 — '+message
                    self.append(message+'\n'+value['output'])
                    messagebox.showinfo('처리 결과',message+'\n\n결과 엑셀의 검토내역과 .review.csv를 확인하세요.')
        except queue.Empty:
            pass
        self.root.after(100,self.poll)

    def open_output(self):
        folder=Path(self.output.get()).parent
        if folder.exists():
            os.startfile(folder)

    def close(self):
        if self.running:
            self.cancel.set()
            self.append('중지 요청됨. 현재 페이지 처리 후 저장이 끝날 때까지 기다려 주세요.')
            return
        self.root.destroy()

def launch_gui():
    root=tk.Tk()
    root.option_add('*Font',('맑은 고딕',10))
    Application(root)
    root.mainloop()
