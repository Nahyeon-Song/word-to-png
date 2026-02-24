import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import TkinterDnD, DND_FILES
import os
import sys
import shutil
import subprocess
import threading
import tempfile
from pathlib import Path


def convert_word_to_png(docx_path, output_dir, dpi=150, progress_callback=None, value_callback=None, cancel_event=None):
    """Word 문서를 페이지별 PNG 이미지로 변환합니다."""
    import pythoncom
    import win32com.client
    import fitz  # PyMuPDF

    # 별도 스레드에서 COM 사용 시 반드시 필요
    pythoncom.CoInitialize()

    docx_path = Path(docx_path)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory() as tmp_dir:
        # 파일명의 특수문자(괄호, 이중점 등)로 인한 Word COM 오류를 방지하기 위해
        # 단순한 이름으로 임시 복사 후 변환
        safe_docx = Path(tmp_dir) / "input.docx"
        shutil.copy2(str(docx_path), str(safe_docx))
        pdf_path = Path(tmp_dir) / "input.pdf"

        if progress_callback:
            progress_callback(f"PDF 변환 중: {docx_path.name}")
        if value_callback:
            value_callback(0)

        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        try:
            doc = word.Documents.Open(str(safe_docx.resolve()))
            doc.SaveAs(str(pdf_path.resolve()), FileFormat=17)  # 17 = wdFormatPDF
            doc.Close(False)
        finally:
            word.Quit()

        if value_callback:
            value_callback(40)
        if progress_callback:
            progress_callback("PNG 변환 중...")

        doc = fitz.open(str(pdf_path))
        total_pages = len(doc)

        for i in range(total_pages):
            if cancel_event and cancel_event.is_set():
                break

            page = doc[i]
            mat = fitz.Matrix(dpi / 72, dpi / 72)
            pix = page.get_pixmap(matrix=mat)
            output_file = output_dir / f"{docx_path.stem}_page_{i + 1:03d}.png"
            pix.save(str(output_file))

            if value_callback:
                value_callback(40 + int((i + 1) / total_pages * 60))
            if progress_callback:
                progress_callback(f"페이지 {i + 1}/{total_pages} 저장 완료")

        doc.close()

    return total_pages


class App(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("Word → PNG 변환기")
        self.resizable(False, False)
        self.configure(bg="#f5f5f5")

        self.files = []
        self.output_dir = tk.StringVar()
        self._cancel_event = threading.Event()

        style = ttk.Style(self)
        style.theme_use("vista")
        style.configure("App.TButton", font=("Malgun Gothic", 9), padding=(10, 4))
        style.configure("Convert.TButton", font=("Malgun Gothic", 10, "bold"), padding=(20, 6))
        style.configure("Bold.TButton", font=("Malgun Gothic", 9, "bold"), padding=(10, 4))

        self._build_ui()
        self.update_idletasks()
        self.minsize(620, self.winfo_reqheight())

    def _btn(self, parent, text, command, **kwargs):
        style = kwargs.pop("style", "App.TButton")
        return ttk.Button(parent, text=text, command=command, style=style, **kwargs)

    def _build_ui(self):
        BG = "#f5f5f5"

        # 파일 선택 영역
        file_frame = tk.LabelFrame(
            self,
            text="  변환할 파일  ",
            font=("Malgun Gothic", 9),
            bg=BG,
            padx=12,
            pady=8,
        )
        file_frame.pack(fill="x", padx=20, pady=(10, 5))

        btn_row = tk.Frame(file_frame, bg=BG)
        btn_row.pack(fill="x", pady=(0, 6))

        self._btn(btn_row, "파일 추가", self.add_files).pack(side="left", padx=(0, 4))
        self._btn(btn_row, "선택 삭제", self.remove_selected).pack(side="left", padx=(0, 4))
        self._btn(btn_row, "전체 초기화", self.clear_files).pack(side="left")

        list_frame = tk.Frame(file_frame, bg=BG)
        list_frame.pack(fill="x")

        scrollbar = tk.Scrollbar(list_frame, orient="vertical")
        self.file_listbox = tk.Listbox(
            list_frame,
            height=6,
            font=("Malgun Gothic", 9),
            selectmode="extended",
            yscrollcommand=scrollbar.set,
            relief="solid",
            borderwidth=1,
        )
        scrollbar.config(command=self.file_listbox.yview)
        self.file_listbox.pack(side="left", fill="x", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 드래그 앤 드롭 등록
        self.file_listbox.drop_target_register(DND_FILES)
        self.file_listbox.dnd_bind("<<DragEnter>>", self._on_drag_enter)
        self.file_listbox.dnd_bind("<<DragLeave>>", self._on_drag_leave)
        self.file_listbox.dnd_bind("<<Drop>>", self._on_drop)

        self.drop_hint = tk.Label(
            file_frame,
            text=".docx / .doc 파일을 끌어다 놓을 수 있습니다",
            font=("Malgun Gothic", 8),
            bg=BG,
            fg="#999",
        )
        self.drop_hint.pack(anchor="w", pady=(3, 0))

        self.file_count_label = tk.Label(
            file_frame,
            text="0개 파일 선택됨",
            font=("Malgun Gothic", 8),
            bg=BG,
            fg="#999",
        )
        self.file_count_label.pack(anchor="e", pady=(2, 0))

        # 저장 폴더 선택
        out_frame = tk.LabelFrame(
            self,
            text="  저장 폴더  ",
            font=("Malgun Gothic", 9),
            bg=BG,
            padx=12,
            pady=8,
        )
        out_frame.pack(fill="x", padx=20, pady=5)

        out_row = tk.Frame(out_frame, bg=BG)
        out_row.pack(fill="x")

        tk.Entry(
            out_row,
            textvariable=self.output_dir,
            font=("Malgun Gothic", 9),
            relief="solid",
            borderwidth=1,
        ).pack(side="left", fill="x", expand=True, padx=(0, 8), ipady=3)

        self._btn(out_row, "폴더 선택", self.select_output_dir).pack(side="right")

        tk.Label(
            out_frame,
            text="* 비워두면 첫 번째 파일과 같은 폴더에 'output' 폴더가 생성됩니다",
            font=("Malgun Gothic", 8),
            bg=BG,
            fg="#999",
        ).pack(anchor="w", pady=(4, 0))

        # 해상도 고정 (150 DPI)
        self.dpi_var = tk.IntVar(value=150)

        # # 해상도 설정 UI (비활성화)
        # dpi_frame = tk.LabelFrame(
        #     self,
        #     text="  해상도 (DPI)  ",
        #     font=("Malgun Gothic", 9),
        #     bg=BG,
        #     padx=12,
        #     pady=6,
        # )
        # dpi_frame.pack(fill="x", padx=20, pady=5)
        #
        # dpi_inner = tk.Frame(dpi_frame, bg=BG)
        # dpi_inner.pack(anchor="w")
        #
        # for dpi_val, label, desc in [
        #     (72,  "72 DPI",  "파일 작음, 화질 낮음"),
        #     (150, "150 DPI", "권장 — 균형잡힌 화질"),
        #     (300, "300 DPI", "인쇄 품질, 파일 큼"),
        # ]:
        #     row = tk.Frame(dpi_inner, bg=BG)
        #     row.pack(anchor="w", pady=1)
        #     tk.Radiobutton(
        #         row,
        #         text=label,
        #         variable=self.dpi_var,
        #         value=dpi_val,
        #         font=("Malgun Gothic", 9),
        #         bg=BG,
        #     ).pack(side="left")
        #     tk.Label(
        #         row,
        #         text=f"({desc})",
        #         font=("Malgun Gothic", 8),
        #         bg=BG,
        #         fg="#999",
        #     ).pack(side="left")

        # 진행 상태
        progress_frame = tk.Frame(self, bg=BG)
        progress_frame.pack(fill="x", padx=20, pady=(8, 0))

        self.progress_label = tk.Label(
            progress_frame,
            text="",
            font=("Malgun Gothic", 8),
            bg=BG,
            fg="#666",
            anchor="w",
        )
        self.progress_label.pack(fill="x")

        self.progress_bar = ttk.Progressbar(
            progress_frame, mode="determinate", maximum=100, value=0
        )
        self.progress_bar.pack(fill="x", pady=(3, 0))

        # 버튼 행
        btn_bottom = tk.Frame(self, bg="#f5f5f5")
        btn_bottom.pack(anchor="e", padx=20, pady=(6, 10))

        self.cancel_btn = self._btn(
            btn_bottom, "강제종료", self.cancel_conversion,
        )
        self.cancel_btn.pack(side="left", padx=(0, 6))
        self.cancel_btn.config(state="disabled")

        self.convert_btn = self._btn(
            btn_bottom, "변환 시작", self.start_conversion,
            style="Bold.TButton",
        )
        self.convert_btn.pack(side="left")

    def _on_drag_enter(self, event):
        self.file_listbox.config(bg="#ddeeff", relief="solid")

    def _on_drag_leave(self, event):
        self.file_listbox.config(bg="white", relief="solid")

    def _on_drop(self, event):
        self.file_listbox.config(bg="white", relief="solid")
        paths = self.tk.splitlist(event.data)
        for p in paths:
            path = Path(p)
            if path.suffix.lower() in (".docx", ".doc") and str(path) not in self.files:
                self.files.append(str(path))
                self.file_listbox.insert("end", path.name)
        self._update_file_count()

    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="Word 파일 선택",
            filetypes=[("Word 문서", "*.docx *.doc"), ("모든 파일", "*.*")],
        )
        for p in paths:
            if p not in self.files:
                self.files.append(p)
                self.file_listbox.insert("end", Path(p).name)
        self._update_file_count()

    def remove_selected(self):
        selected = list(self.file_listbox.curselection())
        for idx in reversed(selected):
            self.files.pop(idx)
            self.file_listbox.delete(idx)
        self._update_file_count()

    def clear_files(self):
        self.files.clear()
        self.file_listbox.delete(0, "end")
        self._update_file_count()

    def _update_file_count(self):
        count = len(self.files)
        self.file_count_label.config(text=f"{count}개 파일 선택됨")

    def select_output_dir(self):
        path = filedialog.askdirectory(title="저장 폴더 선택")
        if path:
            self.output_dir.set(path)

    def start_conversion(self):
        if not self.files:
            messagebox.showwarning("경고", "변환할 파일을 추가해주세요.")
            return

        if not self.output_dir.get():
            default_out = str(Path(self.files[0]).parent / "output")
            self.output_dir.set(default_out)

        self._cancel_event.clear()
        self.progress_label.config(text="")
        self.progress_bar.config(value=0)
        self.convert_btn.config(state="disabled")
        self.cancel_btn.config(state="normal")

        thread = threading.Thread(target=self._run_conversion, daemon=True)
        thread.start()

    def _run_conversion(self):
        total_pages = 0
        errors = []
        total = len(self.files)

        for i, file_path in enumerate(self.files, 1):
            if self._cancel_event.is_set():
                break

            self._set_progress(f"[{i}/{total}] {Path(file_path).name} 처리 중...")
            base = (i - 1) / total * 100
            share = 100 / total

            def make_value_cb(b, s):
                def cb(pct_within):
                    self.after(0, self.progress_bar.config, {"value": b + pct_within * s / 100})
                return cb

            try:
                pages = convert_word_to_png(
                    file_path,
                    self.output_dir.get(),
                    dpi=self.dpi_var.get(),
                    progress_callback=self._set_progress,
                    value_callback=make_value_cb(base, share),
                    cancel_event=self._cancel_event,
                )
                total_pages += pages
            except Exception as e:
                if not self._cancel_event.is_set():
                    errors.append(f"{Path(file_path).name}: {e}")
                self.after(0, self.progress_bar.config, {"value": int(i / total * 100)})

        self.after(0, self._on_done, total_pages, errors)

    def cancel_conversion(self):
        self._cancel_event.set()
        self._set_progress("취소 중... (현재 작업 완료 후 중단됩니다)")
        # Word 프로세스 강제 종료
        subprocess.run(
            ["taskkill", "/F", "/IM", "WINWORD.EXE"],
            capture_output=True,
        )

    def _set_progress(self, msg):
        self.after(0, lambda: self.progress_label.config(text=msg))

    def _on_done(self, total_pages, errors):
        self.convert_btn.config(state="normal")
        self.cancel_btn.config(state="disabled")

        if self._cancel_event.is_set():
            self.progress_label.config(text="취소됨")
            self.progress_bar.config(value=0)
            return

        self.progress_bar.config(value=100)
        self.progress_label.config(text="완료!")

        if errors:
            err_msg = "\n".join(errors)
            messagebox.showerror(
                "오류 발생",
                f"일부 파일 변환에 실패했습니다:\n\n{err_msg}\n\n"
                "Microsoft Word가 설치되어 있는지 확인하세요.",
            )
        else:
            output_path = self.output_dir.get()
            open_folder = messagebox.askyesno(
                "변환 완료",
                f"총 {total_pages}페이지를 PNG로 저장했습니다.\n\n"
                f"저장 위치: {output_path}\n\n"
                "저장 폴더를 열까요?",
            )
            if open_folder:
                os.startfile(output_path)


if __name__ == "__main__":
    app = App()
    app.mainloop()
