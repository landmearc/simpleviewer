import os
import io
import sys
import tempfile
import time
from PIL import Image, ImageTk, ImageSequence
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD

import fitz  # PyMuPDF
try:
    import win32com.client
    import pythoncom
except ImportError:
    win32com = None

# CAD用ライブラリ
try:
    import ezdxf
    from ezdxf.addons.drawing import RenderContext, Frontend
    from ezdxf.addons.drawing.matplotlib import MatplotlibBackend
    import matplotlib.pyplot as plt
    # 日本語フォント設定
    plt.rcParams['font.family'] = 'MS Gothic'
except ImportError:
    ezdxf = None

class ViewerApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("マルチビューアー v4.3.6 (超高精細DXF/GIF/TIFF対応)")
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
        self._preview_images = []

        self._build_ui()

    def _build_ui(self):
        nav = ttk.Frame(self, padding=5)
        nav.pack(side="top", fill="x")
        
        ttk.Button(nav, text="◀ 前へ", command=self.prev_page, width=8).pack(side="left", padx=5)
        self.page_label = ttk.Label(nav, text="0 / 0", font=("Meiryo", 10, "bold"))
        self.page_label.pack(side="left", padx=15)
        ttk.Button(nav, text="次へ ▶", command=self.next_page, width=8).pack(side="left", padx=5)
        
        ttk.Button(nav, text="📸 画像保存", command=self.save_current_as_image, width=12).pack(side="left", padx=10)
        
        # ステータス表示
        self.status_label = ttk.Label(nav, text="待機中", foreground="blue", font=("Meiryo", 9))
        self.status_label.pack(side="left", padx=20)

        ttk.Button(nav, text="開く", command=self.pick_file_dialog).pack(side="right", padx=10)

        paned = ttk.Panedwindow(self, orient="vertical")
        paned.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(paned, bg="#1a1a1a", highlightthickness=0)
        paned.add(self.canvas, weight=3)

        self.text_area = tk.Text(paned, wrap="none", font=("Consolas", 10), bg="#f8f8f8", height=8)
        paned.add(self.text_area, weight=1)

        self.drop_target_register(DND_FILES)
        self.dnd_bind('<<Drop>>', self.on_drop)
        self.bind("<Configure>", self._on_window_resize)
        self.canvas.bind("<MouseWheel>", self._on_mouse_wheel)
        self.canvas.bind("<ButtonPress-1>", self._on_drag_start)
        self.canvas.bind("<B1-Motion>", self._on_drag_move)

    def load_file(self, path):
        self.status_label.config(text="読込中...", foreground="red")
        self.update_idletasks() # 画面を更新して「読込中」を出す

        if self._active_doc is not None:
            try:
                if not getattr(self._active_doc, "is_closed", True):
                    self._active_doc.close()
            except: pass
        self._active_doc = None
        self._preview_images = []
        
        self.current_file_path = os.path.abspath(path)
        ext = os.path.splitext(self.current_file_path)[1].lower()
        self._page_index = 0
        full_text = ""

        try:
            if ext == ".dxf":
                # --- 600 DPI で超高精細レンダリング ---
                img = self._render_dxf_to_image(self.current_file_path)
                if img: 
                    self._preview_images = [img]
                    self._doc_type = "preview_imgs"
                full_text = f"DXF図面: {os.path.basename(path)} (600 DPI)"

            elif ext in [".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff"]:
                with Image.open(self.current_file_path) as img_obj:
                    self._preview_images = [frame.copy() for frame in ImageSequence.Iterator(img_obj)]
                self._doc_type = "preview_imgs"
                full_text = f"画像: {os.path.basename(path)}"

            elif ext in [".docx", ".doc"]:
                self._active_doc = self.convert_to_pdf_preview(self.current_file_path, "Word.Application")
                self._doc_type = "paged"
            elif ext in [".xlsx", ".xls"]:
                self._active_doc = self.convert_to_pdf_preview(self.current_file_path, "Excel.Application")
                self._doc_type = "paged"
            elif ext == ".pdf":
                self._active_doc = fitz.open(self.current_file_path)
                self._doc_type = "paged"
            elif ext in [".txt", ".py", ".csv", ".log"]:
                full_text = self._read_text_safe(self.current_file_path)
                self._doc_type = "text_only"
                self.canvas.delete("all")

            if self._doc_type == "paged":
                self._total_pages = len(self._active_doc)
                full_text = "".join([p.get_text() for p in self._active_doc])
            elif self._doc_type == "preview_imgs":
                self._total_pages = len(self._preview_images)
            else:
                self._total_pages = 1

            self.render_current(fit=True)
            self._display_text(full_text)
            self.page_label.config(text=f"{self._page_index + 1} / {self._total_pages}")
            self.status_label.config(text="完了", foreground="green")
        except Exception as e:
            self.status_label.config(text="エラー", foreground="red")
            messagebox.showerror("Error", f"読込失敗:\n{e}")

    def _render_dxf_to_image(self, path):
        if ezdxf is None: return None
        try:
            plt.close('all')
            doc = ezdxf.readfile(path)
            msp = doc.modelspace()
            
            # 高解像度を維持しつつ、キャンバスを広めに設定
            fig = plt.figure(figsize=(20, 12)) 
            ax = fig.add_axes([0, 0, 1, 1])
            
            ctx = RenderContext(doc)
            # --- 細線化のための設定 ---
            # 線の太さを一律で細く見せるため、背景とのコントラストを調整
            out = MatplotlibBackend(ax)
            
            # Frontendの設定で線幅の倍率を 0.1 程度に落とすと非常に細くなります
            frontend = Frontend(ctx, out)
            frontend.draw_layout(msp, finalize=True)
            
            # matplotlib自体のデフォルト線幅も細く設定
            for line in ax.get_lines():
                line.set_linewidth(0.5) 
            
            img_buf = io.BytesIO()
            # 600 DPI で保存
            fig.savefig(img_buf, format='png', dpi=600, bbox_inches='tight', facecolor='#1a1a1a')
            plt.close(fig)
            return Image.open(img_buf)
        except Exception as e:
            print(f"DXF Render Error: {e}")
            return None

    def render_current(self, fit=False):
        if self._doc_type in ["none", "text_only"]: return
        cw, ch = self.canvas.winfo_width(), self.canvas.winfo_height()
        if cw < 50: cw, ch = 800, 450
        raw = None
        
        if self._doc_type == "paged" and self._active_doc:
            page = self._active_doc[self._page_index]
            if fit: self._zoom = min(cw/page.rect.width, ch/page.rect.height) * 0.95; self._offx, self._offy = cw//2, ch//2
            pix = page.get_pixmap(matrix=fitz.Matrix(self._zoom, self._zoom))
            raw = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        elif self._doc_type == "preview_imgs":
            raw_orig = self._preview_images[self._page_index].convert("RGB")
            if fit: self._zoom = min(cw/raw_orig.size[0], ch/raw_orig.size[1]) * 0.95; self._offx, self._offy = cw//2, ch//2
            nw, nh = int(raw_orig.size[0]*self._zoom), int(raw_orig.size[1]*self._zoom)
            if nw > 0 and nh > 0:
                raw = raw_orig.resize((nw, nh), Image.Resampling.LANCZOS)

        if raw:
            self.canvas.delete("all")
            self._tk_main_img = ImageTk.PhotoImage(raw)
            self.canvas.create_image(self._offx, self._offy, anchor="center", image=self._tk_main_img)

    def save_current_as_image(self):
        if not self.current_file_path: return
        save_path = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG files", "*.png")])
        if save_path:
            try:
                if self._doc_type == "paged":
                    self._active_doc[self._page_index].get_pixmap(matrix=fitz.Matrix(3,3)).save(save_path)
                elif self._doc_type == "preview_imgs":
                    self._preview_images[self._page_index].save(save_path)
                messagebox.showinfo("成功", "保存しました")
            except Exception as e: messagebox.showerror("Error", e)

    def convert_to_pdf_preview(self, path, app_name):
        pythoncom.CoInitialize()
        app = win32com.client.Dispatch(app_name); app.Visible = False
        tmp_path = os.path.join(tempfile.gettempdir(), f"pv_{int(time.time())}.pdf")
        try:
            if "Word" in app_name:
                d = app.Documents.Open(path, ReadOnly=True); d.ExportAsFixedFormat(tmp_path, 17); d.Close(False)
            else:
                w = app.Workbooks.Open(path, ReadOnly=True); w.ExportAsFixedFormat(0, tmp_path); w.Close(False)
            return fitz.open(tmp_path)
        finally: app.Quit()

    def _display_text(self, content):
        self.text_area.config(state="normal"); self.text_area.delete("1.0", tk.END); self.text_area.insert(tk.END, content); self.text_area.config(state="disabled")
    def _read_text_safe(self, path):
        for enc in ["utf-8", "cp932"]:
            try:
                with open(path, "r", encoding=enc) as f: return f.read()
            except: continue
        return "読込不可"
    def on_drop(self, event):
        path = event.data.strip('{}').replace('"', ''); self.load_file(path) if os.path.exists(path) else None
    def pick_file_dialog(self):
        p = filedialog.askopenfilename(); self.load_file(p) if p else None
    def _on_mouse_wheel(self, event): self._zoom *= (1.1 if event.delta > 0 else 0.9); self.render_current()
    def _on_drag_start(self, event): self._drag_data["x"], self._drag_data["y"] = event.x, event.y
    def _on_drag_move(self, event):
        self._offx += event.x - self._drag_data["x"]; self._offy += event.y - self._drag_data["y"]
        self._drag_data["x"], self._drag_data["y"] = event.x, event.y; self.render_current()
    def prev_page(self):
        if self._page_index > 0: self._page_index -= 1; self.render_current(fit=True); self.page_label.config(text=f"{self._page_index + 1} / {self._total_pages}")
    def next_page(self):
        if self._page_index < self._total_pages - 1: self._page_index += 1; self.render_current(fit=True); self.page_label.config(text=f"{self._page_index + 1} / {self._total_pages}")
    def _on_window_resize(self, event):
        if self.current_file_path: self.render_current(fit=True)

if __name__ == "__main__":
    ViewerApp().mainloop()