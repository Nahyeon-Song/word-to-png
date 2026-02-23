import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import sys
import threading
import tempfile
from pathlib import Path


def convert_word_to_png(docx_path, output_dir, dpi=150, progress_callback=None):
    """Word 문서를 페이지별 PNG 이미지로 변환합니다."""
    from docx2pdf import convert
    import fitz  # PyMuPDF

    docx_path = Path(docx_path)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory() as tmp_dir:
        pdf_path = Path(tmp_dir) / (docx_path.stem + ".pdf")

        if progress_callback:
            progress_callback(f"PDF 변환 중: {docx_path.name}")

        convert(str(docx_path), str(pdf_path))

        if progress_callback:
            progress_callback("PNG 변환 중...")

        doc = fitz.open(str(pdf_path))
        total_pages = len(doc)

        for i in range(total_pages):
            page = doc[i]
            mat = fitz.Matrix(dpi / 72, dpi / 72)
            pix = page.get_pixmap(matrix=mat)
            output_file = output_dir / f"{docx_path.stem}_page_{i + 1:03d}.png"
            pix.save(str(output_file))

            if progress_callback:
                progress_callback(f"페이지 {i + 1}/{total_pages} 저장 완료")

        doc.close()

    return total_pages


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Word → PNG 변환기")
        self.geometry("620x560")
        self.resizable(False, False)
        self.configure(bg="#f5f5f5")

        self.files = []
        self.output_dir = tk.StringVar()

        self._build_ui()

    def _build_ui(self):
        # 헤더
        header = tk.Frame(self, bg="#2c3e50", height=70)
        header.pack(fill="x")
        header.pack_propagate(False)

        tk.Label(
            header,
            text="Word → PNG 변환기",
            font=("Malgun Gothic", 16, "bold"),
            bg="#2c3e50",
            fg="white",
        ).pack(side="left", padx=20, pady=15)

        tk.Label(
            header,
            text="페이지별 PNG 저장",
            font=("Malgun Gothic", 10),
            bg="#2c3e50",
            fg="#bdc3c7",
        ).pack(side="left", pady=15)

        # 파일 선택 영역
        file_frame = tk.LabelFrame(
            self,
            text="  변환할 파일  ",
            font=("Malgun Gothic", 10, "bold"),
            bg="#f5f5f5",
            fg="#2c3e50",
            padx=12,
            pady=10,
        )
        file_frame.pack(fill="x", padx=20, pady=(15, 5))

        btn_row = tk.Frame(file_frame, bg="#f5f5f5")
        btn_row.pack(fill="x", pady=(0, 6))

        tk.Button(
            btn_row,
            text="+ 파일 추가",
            command=self.add_files,
            font=("Malgun Gothic", 9, "bold"),
            bg="#3498db",
            fg="white",
            relief="flat",
            cursor="hand2",
            padx=10,
            pady=4,
        ).pack(side="left", padx=(0, 6))

        tk.Button(
            btn_row,
            text="선택 삭제",
            command=self.remove_selected,
            font=("Malgun Gothic", 9),
            bg="#e67e22",
            fg="white",
            relief="flat",
            cursor="hand2",
            padx=10,
            pady=4,
        ).pack(side="left", padx=(0, 6))

        tk.Button(
            btn_row,
            text="전체 초기화",
            command=self.clear_files,
            font=("Malgun Gothic", 9),
            bg="#e74c3c",
            fg="white",
            relief="flat",
            cursor="hand2",
            padx=10,
            pady=4,
        ).pack(side="left")

        list_frame = tk.Frame(file_frame, bg="#f5f5f5")
        list_frame.pack(fill="x")

        scrollbar = tk.Scrollbar(list_frame, orient="vertical")
        self.file_listbox = tk.Listbox(
            list_frame,
            height=6,
            font=("Malgun Gothic", 9),
            selectmode="extended",
            yscrollcommand=scrollbar.set,
            bg="white",
            relief="solid",
            borderwidth=1,
        )
        scrollbar.config(command=self.file_listbox.yview)
        self.file_listbox.pack(side="left", fill="x", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.file_count_label = tk.Label(
            file_frame,
            text="0개 파일 선택됨",
            font=("Malgun Gothic", 8),
            bg="#f5f5f5",
            fg="#7f8c8d",
        )
        self.file_count_label.pack(anchor="e", pady=(4, 0))

        # 저장 폴더 선택
        out_frame = tk.LabelFrame(
            self,
            text="  저장 폴더  ",
            font=("Malgun Gothic", 10, "bold"),
            bg="#f5f5f5",
            fg="#2c3e50",
            padx=12,
            pady=10,
        )
        out_frame.pack(fill="x", padx=20, pady=5)

        out_row = tk.Frame(out_frame, bg="#f5f5f5")
        out_row.pack(fill="x")

        tk.Entry(
            out_row,
            textvariable=self.output_dir,
            font=("Malgun Gothic", 9),
            bg="white",
            relief="solid",
            borderwidth=1,
        ).pack(side="left", fill="x", expand=True, padx=(0, 8), ipady=4)

        tk.Button(
            out_row,
            text="폴더 선택",
            command=self.select_output_dir,
            font=("Malgun Gothic", 9),
            bg="#3498db",
            fg="white",
            relief="flat",
            cursor="hand2",
            padx=10,
            pady=4,
        ).pack(side="right")

        tk.Label(
            out_frame,
            text="* 비워두면 첫 번째 파일과 같은 폴더에 'output' 폴더가 생성됩니다",
            font=("Malgun Gothic", 8),
            bg="#f5f5f5",
            fg="#7f8c8d",
        ).pack(anchor="w", pady=(4, 0))

        # 해상도 설정
        dpi_frame = tk.LabelFrame(
            self,
            text="  해상도 (DPI)  ",
            font=("Malgun Gothic", 10, "bold"),
            bg="#f5f5f5",
            fg="#2c3e50",
            padx=12,
            pady=8,
        )
        dpi_frame.pack(fill="x", padx=20, pady=5)

        self.dpi_var = tk.IntVar(value=150)
        dpi_inner = tk.Frame(dpi_frame, bg="#f5f5f5")
        dpi_inner.pack(anchor="w")

        for dpi_val, label, desc in [
            (72, "저해상도 (72 DPI)", "파일 작음, 화질 낮음"),
            (150, "중해상도 (150 DPI)", "권장 — 균형잡힌 화질"),
            (300, "고해상도 (300 DPI)", "인쇄 품질, 파일 큼"),
        ]:
            row = tk.Frame(dpi_inner, bg="#f5f5f5")
            row.pack(anchor="w", pady=1)
            tk.Radiobutton(
                row,
                text=label,
                variable=self.dpi_var,
                value=dpi_val,
                font=("Malgun Gothic", 9),
                bg="#f5f5f5",
            ).pack(side="left")
            tk.Label(
                row,
                text=f"({desc})",
                font=("Malgun Gothic", 8),
                bg="#f5f5f5",
                fg="#7f8c8d",
            ).pack(side="left")

        # 진행 상태
        progress_frame = tk.Frame(self, bg="#f5f5f5")
        progress_frame.pack(fill="x", padx=20, pady=(8, 0))

        self.progress_label = tk.Label(
            progress_frame,
            text="",
            font=("Malgun Gothic", 9),
            bg="#f5f5f5",
            fg="#555",
            anchor="w",
        )
        self.progress_label.pack(fill="x")

        self.progress_bar = ttk.Progressbar(
            progress_frame, mode="indeterminate", length=580
        )
        self.progress_bar.pack(fill="x", pady=(3, 0))

        # 변환 버튼
        self.convert_btn = tk.Button(
            self,
            text="▶  변환 시작",
            command=self.start_conversion,
            font=("Malgun Gothic", 13, "bold"),
            bg="#27ae60",
            fg="white",
            relief="flat",
            cursor="hand2",
            height=2,
            width=18,
        )
        self.convert_btn.pack(pady=12)

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

        self.convert_btn.config(state="disabled")
        self.progress_bar.start(10)

        thread = threading.Thread(target=self._run_conversion, daemon=True)
        thread.start()

    def _run_conversion(self):
        total_pages = 0
        errors = []

        for i, file_path in enumerate(self.files, 1):
            self._set_progress(f"[{i}/{len(self.files)}] {Path(file_path).name} 처리 중...")
            try:
                pages = convert_word_to_png(
                    file_path,
                    self.output_dir.get(),
                    dpi=self.dpi_var.get(),
                    progress_callback=self._set_progress,
                )
                total_pages += pages
            except Exception as e:
                errors.append(f"{Path(file_path).name}: {e}")

        self.after(0, self._on_done, total_pages, errors)

    def _set_progress(self, msg):
        self.after(0, lambda: self.progress_label.config(text=msg))

    def _on_done(self, total_pages, errors):
        self.progress_bar.stop()
        self.convert_btn.config(state="normal")
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
