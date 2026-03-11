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

class ViewerApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("マルチビューアー v4.2 (Office連携・画像書き出し版)")
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
        self._temp_pdf_path = None

        self._build_ui()

    def _build_ui(self):
        nav = ttk.Frame(self, padding=5)
        nav.pack(side="top", fill="x")
        
        ttk.Button(nav, text="◀ 前へ", command=self.prev_page, width=8).pack(side="left", padx=5)
        self.page_label = ttk.Label(nav, text="0 / 0", font=("Meiryo", 10, "bold"))
        self.page_label.pack(side="left", padx=15)
        ttk.Button(nav, text="次へ ▶", command=self.next_page, width=8).pack(side="left", padx=5)
        
        # --- 変更点：印刷の代わりに「画像として保存」 ---
        ttk.Button(nav, text="📸 プレビューを画像保存", command=self.save_current_as_image, width=20).pack(side="left", padx=20)
        
        ttk.Button(nav, text="開く", command=self.pick_file_dialog).pack(side="right", padx=10)

        paned = ttk.Panedwindow(self, orient="vertical")
        paned.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(paned, bg="#2d2d2d", highlightthickness=0)
        paned.add(self.canvas, weight=3)

        self.text_area = tk.Text(paned, wrap="none", font=("Consolas", 10), bg="#f8f8f8", height=8)
        paned.add(self.text_area, weight=1)

        self.drop_target_register(DND_FILES)
        self.dnd_bind('<<Drop>>', self.on_drop)
        self.bind("<Configure>", self._on_window_resize)
        self.canvas.bind("<MouseWheel>", self._on_mouse_wheel)
        self.canvas.bind("<ButtonPress-1>", self._on_drag_start)
        self.canvas.bind("<B1-Motion>", self._on_drag_move)

    def save_current_as_image(self):
        """表示中のページをPNG画像として保存する"""
        if not self._active_doc and self._doc_type != "image":
            messagebox.showwarning("警告", "保存するドキュメントが開かれていません。")
            return

        default_name = f"{os.path.splitext(os.path.basename(self.current_file_path))[0]}_p{self._page_index+1}.png"
        save_path = filedialog.asksaveasfilename(defaultextension=".png", 
                                                 initialfile=default_name,
                                                 filetypes=[("PNG files", "*.png")])
        if not save_path:
            return

        try:
            if self._doc_type == "paged":
                # PDF/Word/Excelから高画質(300dpi相当)で画像化
                page = self._active_doc[self._page_index]
                pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0)) # 2倍の解像度
                pix.save(save_path)
            elif self._doc_type == "image":
                # 元が画像ならそのまま保存
                self._active_doc.save(save_path)
            
            messagebox.showinfo("成功", "画像を保存しました。")
        except Exception as e:
            messagebox.showerror("エラー", f"保存に失敗しました:\n{e}")

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
        full_text = ""

        try:
            if ext in [".docx", ".doc"]:
                self._active_doc = self.convert_to_pdf_preview(self.current_file_path, "Word.Application")
                self._doc_type = "paged"
            elif ext in [".xlsx", ".xls"]:
                self._active_doc = self.convert_to_pdf_preview(self.current_file_path, "Excel.Application")
                self._doc_type = "paged"
            elif ext == ".pdf":
                self._active_doc = fitz.open(self.current_file_path)
                self._doc_type = "paged"
            elif ext in [".jpg", ".jpeg", ".png", ".bmp"]:
                self._active_doc = Image.open(self.current_file_path)
                self._doc_type = "image"
            elif ext in [".txt", ".py", ".csv", ".log"]:
                full_text = self._read_text_safe(self.current_file_path)
                self._doc_type = "text_only"
                self.canvas.delete("all")

            if self._doc_type == "paged":
                full_text = "".join([p.get_text() for p in self._active_doc])
            
            self._total_pages = len(self._active_doc) if self._doc_type == "paged" else 1
            self.render_current(fit=True)
            self._display_text(full_text)
            self.page_label.config(text=f"{self._page_index + 1} / {self._total_pages}")
        except Exception as e:
            messagebox.showerror("Error", f"読込失敗:\n{e}")

    def convert_to_pdf_preview(self, path, app_name):
        if not win32com: raise Exception("pywin32が必要です")
        pythoncom.CoInitialize()
        app = win32com.client.Dispatch(app_name)
        app.Visible = False
        tmp_path = os.path.join(tempfile.gettempdir(), f"pv_{int(time.time())}.pdf")
        try:
            if "Word" in app_name:
                doc = app.Documents.Open(path, ReadOnly=True)
                doc.ExportAsFixedFormat(tmp_path, 17)
                doc.Close(False)
            else:
                wb = app.Workbooks.Open(path, ReadOnly=True)
                wb.ExportAsFixedFormat(0, tmp_path)
                wb.Close(False)
            return fitz.open(tmp_path)
        finally:
            app.Quit()

    def _read_text_safe(self, path):
        for enc in ["utf-8", "cp932"]:
            try:
                with open(path, "r", encoding=enc) as f: return f.read()
            except: continue
        return "読み取り失敗"

    def _display_text(self, content):
        self.text_area.config(state="normal")
        self.text_area.delete("1.0", tk.END)
        self.text_area.insert(tk.END, content)
        self.text_area.config(state="disabled")

    def render_current(self, fit=False):
        if self._doc_type in ["none", "text_only"]: return
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
            if fit:
                self._zoom = min(cw/raw.size[0], ch/raw.size[1]) * 0.95
                self._offx, self._offy = cw//2, ch//2
            pil_img = raw.resize((int(raw.size[0]*self._zoom), int(raw.size[1]*self._zoom)), Image.Resampling.LANCZOS)

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