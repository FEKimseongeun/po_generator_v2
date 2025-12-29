"""
PO Generator GUI v2.0
플레이스홀더 기반 PO 자동 생성 인터페이스
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pathlib import Path
import threading
import json
import os
from typing import Optional, Dict
from datetime import datetime

import sys
sys.path.insert(0, str(Path(__file__).parent.parent))

from core.mom_parser import parse_mom, MOMData
from core.po_generator import generate_po, POGenerator


class POGeneratorApp:
    """PO Generator 메인 애플리케이션"""
    
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("MOM to PO Generator v2.0")
        self.root.geometry("1000x750")
        self.root.minsize(900, 650)
        
        # 변수
        self.mom_path = tk.StringVar()
        self.template_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.mom_data: Optional[MOMData] = None
        self.template_placeholders: list = []
        
        # UI 구성
        self._create_ui()
        self._update_status("프로그램 준비 완료. MOM 파일과 템플릿을 선택하세요.")
    
    def _create_ui(self):
        """UI 구성"""
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # === 파일 선택 ===
        file_frame = ttk.LabelFrame(main_frame, text="📁 파일 선택", padding="10")
        file_frame.pack(fill=tk.X, pady=(0, 10))
        
        # MOM 파일
        ttk.Label(file_frame, text="MOM 파일:", width=12, anchor='e').grid(row=0, column=0, padx=5, pady=5)
        ttk.Entry(file_frame, textvariable=self.mom_path, width=65).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="찾아보기", command=self._browse_mom, width=10).grid(row=0, column=2, padx=5)
        ttk.Button(file_frame, text="분석", command=self._analyze_mom, width=8).grid(row=0, column=3, padx=5)
        
        # 템플릿 파일
        ttk.Label(file_frame, text="PO 템플릿:", width=12, anchor='e').grid(row=1, column=0, padx=5, pady=5)
        ttk.Entry(file_frame, textvariable=self.template_path, width=65).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="찾아보기", command=self._browse_template, width=10).grid(row=1, column=2, padx=5)
        ttk.Button(file_frame, text="분석", command=self._analyze_template, width=8).grid(row=1, column=3, padx=5)
        
        # 출력 파일
        ttk.Label(file_frame, text="출력 파일:", width=12, anchor='e').grid(row=2, column=0, padx=5, pady=5)
        ttk.Entry(file_frame, textvariable=self.output_path, width=65).grid(row=2, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="저장 위치", command=self._browse_output, width=10).grid(row=2, column=2, padx=5)
        
        # === 데이터 미리보기 ===
        preview_frame = ttk.LabelFrame(main_frame, text="📋 데이터 미리보기", padding="10")
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        paned = ttk.PanedWindow(preview_frame, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True)
        
        # MOM 추출 데이터
        mom_frame = ttk.LabelFrame(paned, text="MOM 추출 필드", padding="5")
        paned.add(mom_frame, weight=1)
        
        self.mom_tree = ttk.Treeview(mom_frame, columns=('field', 'value'), show='headings', height=12)
        self.mom_tree.heading('field', text='필드명')
        self.mom_tree.heading('value', text='값')
        self.mom_tree.column('field', width=150)
        self.mom_tree.column('value', width=300)
        
        mom_scroll = ttk.Scrollbar(mom_frame, orient=tk.VERTICAL, command=self.mom_tree.yview)
        self.mom_tree.configure(yscrollcommand=mom_scroll.set)
        self.mom_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        mom_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 템플릿 플레이스홀더
        tpl_frame = ttk.LabelFrame(paned, text="템플릿 플레이스홀더", padding="5")
        paned.add(tpl_frame, weight=1)
        
        self.tpl_tree = ttk.Treeview(tpl_frame, columns=('placeholder', 'status'), show='headings', height=12)
        self.tpl_tree.heading('placeholder', text='플레이스홀더')
        self.tpl_tree.heading('status', text='상태')
        self.tpl_tree.column('placeholder', width=180)
        self.tpl_tree.column('status', width=120)
        
        tpl_scroll = ttk.Scrollbar(tpl_frame, orient=tk.VERTICAL, command=self.tpl_tree.yview)
        self.tpl_tree.configure(yscrollcommand=tpl_scroll.set)
        self.tpl_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tpl_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # === 상세 내용 보기 ===
        detail_frame = ttk.LabelFrame(main_frame, text="📝 선택 필드 상세 내용", padding="5")
        detail_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        self.detail_text = scrolledtext.ScrolledText(detail_frame, height=6, font=('Consolas', 9), wrap=tk.WORD)
        self.detail_text.pack(fill=tk.BOTH, expand=True)
        
        # 트리뷰 선택 이벤트
        self.mom_tree.bind('<<TreeviewSelect>>', self._on_mom_select)
        
        # === 버튼 ===
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Button(btn_frame, text="🔄 새로고침", command=self._refresh_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="📄 PO 생성", command=self._generate_po).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_frame, text="📂 출력 폴더", command=self._open_output_folder).pack(side=tk.RIGHT, padx=5)
        
        # === 상태바 ===
        self.status_bar = ttk.Label(main_frame, text="준비", relief=tk.SUNKEN, anchor='w', padding=(5, 2))
        self.status_bar.pack(fill=tk.X, side=tk.BOTTOM)
        
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
    
    def _browse_mom(self):
        path = filedialog.askopenfilename(
            title="MOM 파일 선택",
            filetypes=[("Word 문서", "*.docx"), ("모든 파일", "*.*")]
        )
        if path:
            self.mom_path.set(path)
            self._auto_output_path()
            self._analyze_mom()
    
    def _browse_template(self):
        path = filedialog.askopenfilename(
            title="PO 템플릿 선택",
            filetypes=[("Word 문서", "*.docx"), ("모든 파일", "*.*")]
        )
        if path:
            self.template_path.set(path)
            self._analyze_template()
    
    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            title="PO 저장 위치",
            defaultextension=".docx",
            filetypes=[("Word 문서", "*.docx")]
        )
        if path:
            self.output_path.set(path)
    
    def _auto_output_path(self):
        mom = self.mom_path.get()
        if mom:
            mom_file = Path(mom)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output = mom_file.parent / f"PO_{mom_file.stem}_{timestamp}.docx"
            self.output_path.set(str(output))
    
    def _analyze_mom(self):
        mom_path = self.mom_path.get()
        if not mom_path:
            messagebox.showwarning("경고", "MOM 파일을 선택하세요.")
            return
        
        self._show_progress(True)
        self._update_status("MOM 분석 중...")
        
        def analyze():
            try:
                self.mom_data = parse_mom(mom_path)
                self.root.after(0, self._update_mom_tree)
                self.root.after(0, self._update_placeholder_status)
                self.root.after(0, lambda: self._update_status(
                    f"MOM 분석 완료: {len(self.mom_data.fields)}개 필드 추출"
                ))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("오류", f"MOM 분석 실패:\n{e}"))
            finally:
                self.root.after(0, lambda: self._show_progress(False))
        
        threading.Thread(target=analyze, daemon=True).start()
    
    def _analyze_template(self):
        tpl_path = self.template_path.get()
        if not tpl_path:
            messagebox.showwarning("경고", "템플릿 파일을 선택하세요.")
            return
        
        self._show_progress(True)
        self._update_status("템플릿 분석 중...")
        
        def analyze():
            try:
                gen = POGenerator(tpl_path)
                self.template_placeholders = gen.get_template_placeholders()
                self.root.after(0, self._update_template_tree)
                self.root.after(0, self._update_placeholder_status)
                self.root.after(0, lambda: self._update_status(
                    f"템플릿 분석 완료: {len(self.template_placeholders)}개 플레이스홀더"
                ))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("오류", f"템플릿 분석 실패:\n{e}"))
            finally:
                self.root.after(0, lambda: self._show_progress(False))
        
        threading.Thread(target=analyze, daemon=True).start()
    
    def _update_mom_tree(self):
        for item in self.mom_tree.get_children():
            self.mom_tree.delete(item)
        
        if self.mom_data:
            for field, value in sorted(self.mom_data.fields.items()):
                preview = value[:60] + "..." if len(value) > 60 else value
                preview = preview.replace('\n', ' ')
                self.mom_tree.insert('', tk.END, values=(field, preview))
    
    def _update_template_tree(self):
        for item in self.tpl_tree.get_children():
            self.tpl_tree.delete(item)
        
        for ph in self.template_placeholders:
            status = "⏳ 대기"
            self.tpl_tree.insert('', tk.END, values=(f"{{{{{ph}}}}}", status))
    
    def _update_placeholder_status(self):
        """플레이스홀더 상태 업데이트"""
        if not self.template_placeholders:
            return
        
        for item in self.tpl_tree.get_children():
            self.tpl_tree.delete(item)
        
        for ph in self.template_placeholders:
            if self.mom_data and ph in self.mom_data.fields and self.mom_data.fields[ph]:
                status = "✓ 매칭됨"
            elif self.mom_data:
                status = "✗ 데이터 없음"
            else:
                status = "⏳ MOM 필요"
            
            self.tpl_tree.insert('', tk.END, values=(f"{{{{{ph}}}}}", status))
    
    def _on_mom_select(self, event):
        """MOM 필드 선택 시 상세 내용 표시"""
        selection = self.mom_tree.selection()
        if selection and self.mom_data:
            item = self.mom_tree.item(selection[0])
            field_name = item['values'][0]
            full_value = self.mom_data.fields.get(field_name, "")
            
            self.detail_text.delete('1.0', tk.END)
            self.detail_text.insert('1.0', f"[{field_name}]\n\n{full_value}")
    
    def _refresh_all(self):
        if self.mom_path.get():
            self._analyze_mom()
        if self.template_path.get():
            self._analyze_template()
    
    def _generate_po(self):
        if not self.mom_path.get():
            messagebox.showwarning("경고", "MOM 파일을 선택하세요.")
            return
        if not self.template_path.get():
            messagebox.showwarning("경고", "PO 템플릿을 선택하세요.")
            return
        if not self.output_path.get():
            messagebox.showwarning("경고", "출력 파일 위치를 지정하세요.")
            return
        if not self.mom_data:
            messagebox.showwarning("경고", "먼저 MOM 파일을 분석하세요.")
            return
        
        self._show_progress(True)
        self._update_status("PO 생성 중...")
        
        def generate():
            try:
                result_path, replacements = generate_po(
                    self.template_path.get(),
                    self.mom_data,
                    self.output_path.get()
                )
                
                msg = f"PO 생성 완료!\n\n저장 위치: {result_path}\n\n"
                msg += f"교체된 필드 ({len(replacements)}개):\n"
                for r in replacements[:10]:
                    msg += f"  • {{{{{r.placeholder}}}}}\n"
                if len(replacements) > 10:
                    msg += f"  ... 외 {len(replacements) - 10}개"
                
                self.root.after(0, lambda: messagebox.showinfo("완료", msg))
                self.root.after(0, lambda: self._update_status(f"PO 생성 완료: {result_path}"))
                
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("오류", f"PO 생성 실패:\n{e}"))
            finally:
                self.root.after(0, lambda: self._show_progress(False))
        
        threading.Thread(target=generate, daemon=True).start()
    
    def _open_output_folder(self):
        output = self.output_path.get()
        if output:
            folder = Path(output).parent
            if folder.exists():
                os.startfile(str(folder)) if os.name == 'nt' else os.system(f'open "{folder}"')
    
    def _show_progress(self, show: bool):
        if show:
            self.progress.pack(fill=tk.X, side=tk.BOTTOM, before=self.status_bar, pady=(5, 0))
            self.progress.start(10)
        else:
            self.progress.stop()
            self.progress.pack_forget()
    
    def _update_status(self, msg: str):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.status_bar.config(text=f"[{timestamp}] {msg}")


def main():
    root = tk.Tk()
    style = ttk.Style()
    style.theme_use('clam')
    app = POGeneratorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
