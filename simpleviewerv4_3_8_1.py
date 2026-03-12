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

import matplotlib.pyplot as plt
plt.rcParams['font.family'] = 'MS Gothic'

class ViewerApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("マルチビューアー v4.3.8 (GIF/TIFF/SIMA対応)")
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
            if ext in [".sim", ".sima"]:
                pt_dict, polygons, text_lines = self._parse_sima(self.current_file_path)
                img = self._render_sima_to_image(pt_dict, polygons)
                if img:
                    self._preview_images = [img]
                    self._doc_type = "preview_imgs"
                full_text = "\n".join(text_lines)

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

    def _parse_sima(self, path):
        """SIMAファイルを解析。
        戻り値:
          pt_dict  : {seq_no(int): (name, X北, Y東)}
          polygons : [{"name": str, "pts": [(name,X,Y), ...]}, ...]
          text_lines: 生テキスト行リスト
        """
        pt_dict = {}      # {seq_no: (name, x, y)}
        polygons = []     # [{"name":..., "pts":[...]}]
        text_lines = []

        for enc in ["cp932", "utf-8", "utf-8-sig"]:
            try:
                with open(path, "r", encoding=enc) as f:
                    raw_lines = f.readlines()
                break
            except Exception:
                continue
        else:
            return pt_dict, polygons, ["SIMAファイルの読込に失敗しました"]

        cur_polygon = None  # 現在解析中の画地

        for line in raw_lines:
            line = line.rstrip("\r\n")
            text_lines.append(line)
            cols = [c.strip() for c in line.split(",")]
            rec = cols[0].upper() if cols else ""

            # --- 座標レコード (A01) ---
            if rec == "A01" and len(cols) >= 5:
                try:
                    seq = int(cols[1])
                    name = cols[2]
                    x = float(cols[3])   # X(北)
                    y = float(cols[4])   # Y(東)
                    pt_dict[seq] = (name, x, y)
                except ValueError:
                    pass

            # --- 画地開始 (D00) ---
            elif rec == "D00" and len(cols) >= 3:
                cur_polygon = {"name": cols[2], "pts": []}
                polygons.append(cur_polygon)

            # --- 画地の点参照 (B01) ---
            elif rec == "B01" and cur_polygon is not None and len(cols) >= 2:
                try:
                    seq = int(cols[1])
                    cur_polygon["pts"].append(seq)
                except ValueError:
                    pass

            # --- 画地終了 (D99) ---
            elif rec == "D99":
                cur_polygon = None

        return pt_dict, polygons, text_lines

    def _render_sima_to_image(self, pt_dict, polygons):
        """SIMA座標データを正しい結線で描画して PIL Image を返す。"""
        if not pt_dict:
            return None
        try:
            plt.close('all')
            # 白背景（正しい図面スタイル）
            fig, ax = plt.subplots(figsize=(16, 12), facecolor="white")
            ax.set_facecolor("white")
            ax.tick_params(colors="black")
            for spine in ax.spines.values():
                spine.set_edgecolor("#aaa")

            # 全点を散布図でプロット
            all_pts = list(pt_dict.values())
            all_x = [p[2] for p in all_pts]   # Y(東) → 画面横軸
            all_y = [p[1] for p in all_pts]   # X(北) → 画面縦軸
            ax.scatter(all_x, all_y, color="black", s=15, zorder=3)

            # 点名ラベル
            xrange = max(all_x) - min(all_x) if len(all_x) > 1 else 1
            yrange = max(all_y) - min(all_y) if len(all_y) > 1 else 1
            offset = max(xrange, yrange) * 0.006
            for name, px, py in all_pts:
                ax.text(py + offset, px + offset, name,
                        fontsize=5.5, color="#222",
                        fontfamily="MS Gothic", zorder=4)

            # 画地ポリゴンを閉合描画
            poly_colors = ["#1a3c8f", "#2060c0", "#1a5fa0", "#0a4070",
                           "#2255aa", "#163275", "#1848a0", "#1a3060"]
            for i, poly in enumerate(polygons):
                seqs = poly["pts"]
                coords = [pt_dict[s] for s in seqs if s in pt_dict]
                if len(coords) < 2:
                    continue
                ys_p = [c[2] for c in coords]   # Y(東)
                xs_p = [c[1] for c in coords]   # X(北)
                # 閉合（最終→最初を結ぶ）
                ys_p.append(ys_p[0])
                xs_p.append(xs_p[0])
                color = poly_colors[i % len(poly_colors)]
                ax.plot(ys_p, xs_p, color=color, linewidth=1.2, zorder=2)

                # 画地名を重心に表示
                cy = sum(ys_p[:-1]) / len(ys_p[:-1])
                cx = sum(xs_p[:-1]) / len(xs_p[:-1])
                ax.text(cy, cx, poly["name"],
                        fontsize=6, color="#333", ha="center",
                        fontfamily="MS Gothic", zorder=5,
                        bbox=dict(boxstyle="round,pad=0.1", fc="white", alpha=0.6, ec="none"))

            ax.set_aspect("equal")
            ax.grid(True, color="#ddd", linewidth=0.4)
            n_pts = len(pt_dict)
            n_poly = len(polygons)
            ax.set_title(f"SIMA座標データ  {n_pts}点 / {n_poly}画地",
                         fontsize=11, fontfamily="MS Gothic", color="black")
            ax.set_xlabel("Y（東）[m]", fontsize=9, fontfamily="MS Gothic")
            ax.set_ylabel("X（北）[m]", fontsize=9, fontfamily="MS Gothic")

            buf = io.BytesIO()
            fig.savefig(buf, format="png", dpi=200, bbox_inches="tight", facecolor="white")
            plt.close(fig)
            buf.seek(0)
            return Image.open(buf).copy()
        except Exception as e:
            print(f"SIMA Render Error: {e}")
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
        p = filedialog.askopenfilename(
            filetypes=[
                ("すべての対応ファイル", "*.pdf *.sim *.sima *.jpg *.jpeg *.png *.bmp *.gif *.tif *.tiff *.docx *.doc *.xlsx *.xls *.txt *.py *.csv *.log"),
                ("SIMAファイル", "*.sim *.sima"),
                ("PDF", "*.pdf"),
                ("画像", "*.jpg *.jpeg *.png *.bmp *.gif *.tif *.tiff"),
                ("すべて", "*.*"),
            ]
        )
        self.load_file(p) if p else None
    def _on_mouse_wheel(self, event):
        self._zoom *= (1.1 if event.delta > 0 else 0.9)
        self.render_current()
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