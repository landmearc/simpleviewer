import os
import io
import sys
import tempfile
import time
from PIL import Image, ImageTk
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD

import fitz  # PyMuPDF
try:
    import win32com.client
    import pythoncom
except ImportError:
    win32com = None

try:
    import docx
except ImportError:
    docx = None

class ViewerApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("マルチビューアー v3.5 (Word・Text・Python対応)")
        self.geometry("1200x900")

        self.current_file_path = ""
        self._doc_type = "none"
        self._page_index = 0
        self._total_pages = 1
        self._zoom = 1.0
        self._offx, self._offy = 0, 0
        self._drag_data = {"x": 0, "y": 0}
        
        self._tk_main_img = None
        self._active_doc = None 

        self._build_ui()

    def _build_ui(self):
        nav = ttk.Frame(self, padding=5)
        nav.pack(side="top", fill="x")
        
        ttk.Button(nav, text="◀ 前へ", command=self.prev_page, width=10).pack(side="left", padx=10)
        self.page_label = ttk.Label(nav, text="0 / 0", font=("Meiryo", 10, "bold"))
        self.page_label.pack(side="left", padx=20)
        ttk.Button(nav, text="次へ ▶", command=self.next_page, width=10).pack(side="left", padx=10)
        ttk.Button(nav, text="開く", command=self.pick_file_dialog).pack(side="right", padx=10)

        paned = ttk.Panedwindow(self, orient="vertical")
        paned.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(paned, bg="#2d2d2d", highlightthickness=0)
        paned.add(self.canvas, weight=3)

        self.text_area = tk.Text(paned, wrap="none", font=("Consolas", 10), bg="#f8f8f8", height=10)
        h_scroll = ttk.Scrollbar(paned, orient="horizontal", command=self.text_area.xview)
        self.text_area.configure(xscrollcommand=h_scroll.set)
        paned.add(self.text_area, weight=1)

        self.drop_target_register(DND_FILES)
        self.dnd_bind('<<Drop>>', self.on_drop)
        self.bind("<Configure>", self._on_window_resize)
        self.canvas.bind("<MouseWheel>", self._on_mouse_wheel)
        self.canvas.bind("<ButtonPress-1>", self._on_drag_start)
        self.canvas.bind("<B1-Motion>", self._on_drag_move)

    def on_drop(self, event):
        path = event.data.strip('{}').replace('"', '')
        if os.path.exists(path): self.load_file(path)

    def pick_file_dialog(self):
        p = filedialog.askopenfilename()
        if p: self.load_file(p)

    def load_file(self, path):
        if self._active_doc:
            try: self._active_doc.close()
            except: pass
        
        self.current_file_path = os.path.abspath(path)
        ext = os.path.splitext(self.current_file_path)[1].lower()
        self._page_index = 0
        self._doc_type = "none"
        full_text = ""

        try:
            # --- Wordファイルの処理 ---
            if ext in [".docx", ".doc"]:
                self._active_doc = self.convert_word_to_pdf_doc(self.current_file_path)
                self._doc_type = "paged"
                if docx:
                    d = docx.Document(self.current_file_path)
                    full_text = "\n".join([p.text for p in d.paragraphs])

            # --- PDFファイルの処理 ---
            elif ext == ".pdf":
                self._active_doc = fitz.open(self.current_file_path)
                self._doc_type = "paged"
                full_text = "".join([p.get_text() for p in self._active_doc])

            # --- 画像ファイルの処理 ---
            elif ext in [".jpg", ".jpeg", ".png", ".bmp"]:
                self._active_doc = Image.open(self.current_file_path)
                self._doc_type = "image"
                full_text = f"画像ファイル: {os.path.basename(path)}"

            # --- テキスト/プログラムファイルの処理 ---
            elif ext in [".txt", ".py", ".csv", ".log", ".md", ".ini"]:
                full_text = self._read_text_safe(self.current_file_path)
                self._doc_type = "text_only"
                self.canvas.delete("all")
                self.canvas.create_text(400, 200, text="テキストファイルを表示中", fill="white")

            self._total_pages = len(self._active_doc) if self._doc_type == "paged" else 1
            self.render_current(fit=True)
            self._display_text(full_text)
            self.page_label.config(text=f"{self._page_index + 1} / {self._total_pages}")

        except Exception as e:
            messagebox.showerror("Error", f"読込失敗:\n{e}")

    def _read_text_safe(self, path):
        for enc in ["utf-8", "cp932", "shift_jis"]:
            try:
                with open(path, "r", encoding=enc) as f:
                    return f.read()
            except:
                continue
        return "ファイルの読み込みに失敗しました（文字コード不明）"

    def convert_word_to_pdf_doc(self, path):
        pythoncom.CoInitialize()
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        try:
            doc = word.Documents.Open(path, ReadOnly=True)
            tmp_path = os.path.join(tempfile.gettempdir(), f"viewer_pv_{int(time.time())}.pdf")
            doc.ExportAsFixedFormat(tmp_path, 17)
            doc.Close(False)
            return fitz.open(tmp_path)
        finally:
            word.Quit()

    def _display_text(self, content):
        self.text_area.config(state="normal")
        self.text_area.delete("1.0", tk.END)
        self.text_area.insert(tk.END, content)
        self.text_area.config(state="disabled")

    def render_current(self, fit=False):
        if self._doc_type == "none" or self._doc_type == "text_only": return
        self.update_idletasks()
        cw, ch = self.canvas.winfo_width(), self.canvas.winfo_height()
        if cw < 50: cw, ch = 800, 450
        pil_img = None

        if self._doc_type == "paged":
            page = self._active_doc[self._page_index]
            if fit:
                self._zoom = min(cw/page.rect.width, ch/page.rect.height) * 0.95
                self._offx, self._offy = cw//2, ch//2
            pix = page.get_pixmap(matrix=fitz.Matrix(self._zoom, self._zoom))
            pil_img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        elif self._doc_type == "image":
            raw = self._active_doc.convert("RGB")
            iw, ih = raw.size
            if fit:
                self._zoom = min(cw/iw, ch/ih) * 0.95
                self._offx, self._offy = cw//2, ch//2
            pil_img = raw.resize((int(iw*self._zoom), int(ih*self._zoom)), Image.Resampling.LANCZOS)

        if pil_img:
            self.canvas.delete("all")
            self._tk_main_img = ImageTk.PhotoImage(pil_img)
            self.canvas.create_image(self._offx, self._offy, anchor="center", image=self._tk_main_img)

    def _on_mouse_wheel(self, event):
        scale = 1.1 if event.delta > 0 else 0.9
        self._zoom *= scale
        self.render_current(fit=False)

    def _on_drag_start(self, event):
        self._drag_data["x"], self._drag_data["y"] = event.x, event.y

    def _on_drag_move(self, event):
        self._offx += event.x - self._drag_data["x"]
        self._offy += event.y - self._drag_data["y"]
        self._drag_data["x"], self._drag_data["y"] = event.x, event.y
        self.render_current(fit=False)

    def prev_page(self):
        if self._page_index > 0:
            self._page_index -= 1
            self.render_current(fit=True)
            self.page_label.config(text=f"{self._page_index + 1} / {self._total_pages}")

    def next_page(self):
        if self._page_index < self._total_pages - 1:
            self._page_index += 1
            self.render_current(fit=True)
            self.page_label.config(text=f"{self._page_index + 1} / {self._total_pages}")

    def _on_window_resize(self, event):
        if self.current_file_path: self.render_current(fit=True)

if __name__ == "__main__":
    ViewerApp().mainloop()