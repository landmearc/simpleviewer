import os
import math
import sys
from typing import List, Tuple, Optional, Dict

from PIL import Image, ImageTk

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from tkinterdnd2 import DND_FILES, TkinterDnD

import fitz  # PyMuPDF


# -------------------------
# Utilities
# -------------------------
def norm_paths_from_dnd(data: str) -> List[str]:
    paths, buff, in_brace = [], "", False
    for ch in data:
        if ch == "{":
            in_brace = True; buff = ""
        elif ch == "}":
            in_brace = False
            if buff: paths.append(buff); buff = ""
        elif ch == " " and not in_brace:
            if buff: paths.append(buff); buff = ""
        else:
            buff += ch
    if buff: paths.append(buff)
    return [p.strip().strip('"') for p in paths]


# -------------------------
# SIMA parser
# -------------------------
def read_text_guess_encoding(path: str) -> str:
    for enc in ("cp932", "shift_jis", "utf-8"):
        try:
            return open(path, "r", encoding=enc).read()
        except Exception:
            continue
    return open(path, "r", encoding="utf-8", errors="replace").read()


def parse_sima_file(path: str) -> Tuple[Dict[str, Tuple[float, float, Optional[float]]], List[List[str]]]:
    txt = read_text_guess_encoding(path)
    points: Dict[str, Tuple[float, float, Optional[float]]] = {}
    polylines: List[List[str]] = []
    in_d_block = False
    current_poly: List[str] = []

    for ln in [l.strip() for l in txt.splitlines() if l.strip()]:
        parts = [p.strip() for p in ln.split(",")]
        if not parts: continue
        head = parts[0].upper()

        if head == "A01" and len(parts) >= 5:
            name = (parts[2] if len(parts) > 2 else "").strip()
            if name:
                try:
                    x = float(parts[3]); y = float(parts[4])
                    z = None
                    if len(parts) >= 6 and parts[5]:
                        try: z = float(parts[5])
                        except: pass
                    points[name] = (x, y, z)
                except: pass
        elif head == "D00":
            in_d_block = True; current_poly = []
        elif head == "D99":
            if in_d_block and len(current_poly) >= 2:
                polylines.append(current_poly[:])
            in_d_block = False; current_poly = []
        elif in_d_block and head == "B01" and len(parts) >= 3:
            name = parts[2].strip()
            if name: current_poly.append(name)

    return points, polylines


# -------------------------
# Main App
# -------------------------
class ViewerApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF / SIMA ビューアー")
        self.geometry("1200x850")

        # PDF state
        self._pdf_doc: Optional[fitz.Document] = None
        self._pdf_page_index = 0
        self._pdf_zoom = 1.5
        self._pdf_min_zoom = 0.2
        self._pdf_max_zoom = 8.0
        self._pdf_offx = 20.0
        self._pdf_offy = 20.0
        self._pdf_cache: dict = {}
        self._pdf_pan_start = None
        self._pdf_pan_origin = None
        self._pdf_pan_moved = False
        self._pdf_preview_tk = None

        # SIMA state
        self.sima_points: Dict[str, Tuple[float, float, Optional[float]]] = {}
        self.sima_polylines: List[List[str]] = []
        self._sim_zoom = 1.0
        self._sim_min_zoom = 0.05
        self._sim_max_zoom = 50.0
        self._sim_offx = 0.0
        self._sim_offy = 0.0
        self._sim_bbox = None
        self._sim_pan_start = None
        self._sim_pan_origin = None
        self._sim_pan_moved = False

        self.preview_mode_var = tk.StringVar(value="pdf")
        self._build_ui()

    # -------------------------
    # UI
    # -------------------------
    def _build_ui(self):
        root = ttk.Frame(self, padding=6)
        root.pack(fill="both", expand=True)

        paned = ttk.Panedwindow(root, orient="horizontal")
        paned.pack(fill="both", expand=True)

        # ===== LEFT =====
        left = ttk.Frame(paned, padding=6)
        paned.add(left, weight=1)

        file_frm = ttk.LabelFrame(left, text="ファイルを開く", padding=6)
        file_frm.pack(fill="x", pady=(0, 8))

        ttk.Button(file_frm, text="PDF を開く…",  command=self.pick_pdf_dialog).pack(fill="x", pady=2)
        ttk.Button(file_frm, text="SIMA を開く…", command=self.pick_sim_dialog).pack(fill="x", pady=2)

        drop_pdf = ttk.Label(file_frm, text="PDF をここにドロップ",
                             relief="ridge", padding=8, anchor="center")
        drop_pdf.pack(fill="x", pady=(6, 2))
        drop_pdf.drop_target_register(DND_FILES)
        drop_pdf.dnd_bind("<<Drop>>", self.on_drop_pdf)

        drop_sim = ttk.Label(file_frm, text="SIMA をここにドロップ",
                             relief="ridge", padding=8, anchor="center")
        drop_sim.pack(fill="x", pady=2)
        drop_sim.drop_target_register(DND_FILES)
        drop_sim.dnd_bind("<<Drop>>", self.on_drop_sim)

        self.file_info_label = ttk.Label(left, text="ファイル未選択",
                                         foreground="gray", wraplength=220)
        self.file_info_label.pack(anchor="w", pady=(0, 8))

        mode_frm = ttk.LabelFrame(left, text="表示切替", padding=6)
        mode_frm.pack(fill="x", pady=(0, 8))
        ttk.Radiobutton(mode_frm, text="PDF",  variable=self.preview_mode_var,
                        value="pdf",  command=self.update_preview_mode).pack(side="left", padx=8)
        ttk.Radiobutton(mode_frm, text="SIMA", variable=self.preview_mode_var,
                        value="sim",  command=self.update_preview_mode).pack(side="left", padx=8)

        # PDF controls
        pdf_ctrl = ttk.LabelFrame(left, text="PDF 操作", padding=6)
        pdf_ctrl.pack(fill="x", pady=(0, 8))

        nav = ttk.Frame(pdf_ctrl)
        nav.pack(fill="x")
        ttk.Button(nav, text="◀ 前", command=self.pdf_prev_page).pack(side="left")
        ttk.Button(nav, text="次 ▶", command=self.pdf_next_page).pack(side="left", padx=4)
        self.pdf_page_label = ttk.Label(nav, text="page: -/-")
        self.pdf_page_label.pack(side="left", padx=8)

        zoom_row = ttk.Frame(pdf_ctrl)
        zoom_row.pack(fill="x", pady=(4, 0))
        ttk.Button(zoom_row, text="－", width=4, command=self.pdf_zoom_out).pack(side="left")
        ttk.Button(zoom_row, text="＋", width=4, command=self.pdf_zoom_in).pack(side="left", padx=2)
        ttk.Button(zoom_row, text="フィット", command=self.pdf_fit_to_canvas).pack(side="left", padx=6)
        self.pdf_zoom_label = ttk.Label(zoom_row, text="zoom: -")
        self.pdf_zoom_label.pack(side="left", padx=8)

        ttk.Label(pdf_ctrl, text="Ctrl+ホイール: ズーム　ドラッグ / ホイール: スクロール",
                  foreground="gray").pack(anchor="w", pady=(4, 0))

        # SIMA controls
        sim_ctrl = ttk.LabelFrame(left, text="SIMA 操作", padding=6)
        sim_ctrl.pack(fill="x", pady=(0, 8))

        sim_btn = ttk.Frame(sim_ctrl)
        sim_btn.pack(fill="x")
        ttk.Button(sim_btn, text="フィット", command=self.sim_fit_view).pack(side="left")
        ttk.Button(sim_btn, text="100%",    command=self.sim_zoom_100).pack(side="left", padx=4)

        self.sim_info_label = ttk.Label(sim_ctrl, text="SIMA: -", foreground="gray")
        self.sim_info_label.pack(anchor="w", pady=(4, 0))

        ttk.Label(sim_ctrl, text="ホイール: ズーム　ドラッグ: パン",
                  foreground="gray").pack(anchor="w", pady=(4, 0))

        # ===== RIGHT =====
        right = ttk.Frame(paned, padding=6)
        paned.add(right, weight=5)

        # PDF canvas
        self.pdf_preview_frame = ttk.LabelFrame(right, text="PDF プレビュー", padding=4)
        self.pdf_preview_frame.pack(fill="both", expand=True)

        pdf_cf = ttk.Frame(self.pdf_preview_frame)
        pdf_cf.pack(fill="both", expand=True)
        self.pdf_canvas = tk.Canvas(pdf_cf, bg="gray85", highlightthickness=0)
        self.pdf_canvas.pack(side="left", fill="both", expand=True)
        pdf_ys = ttk.Scrollbar(pdf_cf, orient="vertical",   command=self.pdf_canvas.yview)
        pdf_ys.pack(side="right", fill="y")
        pdf_xs = ttk.Scrollbar(self.pdf_preview_frame, orient="horizontal", command=self.pdf_canvas.xview)
        pdf_xs.pack(fill="x")
        self.pdf_canvas.configure(yscrollcommand=pdf_ys.set, xscrollcommand=pdf_xs.set)

        self.pdf_canvas.bind("<ButtonPress-1>",      self.on_pdf_pan_start)
        self.pdf_canvas.bind("<B1-Motion>",          self.on_pdf_pan_move)
        self.pdf_canvas.bind("<ButtonRelease-1>",    self.on_pdf_pan_end)
        self.pdf_canvas.bind("<MouseWheel>",         self.on_pdf_wheel)
        self.pdf_canvas.bind("<Control-MouseWheel>", self.on_pdf_zoom_wheel)
        self.pdf_canvas.bind("<Button-4>",           self.on_pdf_wheel_linux)
        self.pdf_canvas.bind("<Button-5>",           self.on_pdf_wheel_linux)
        self.pdf_canvas.bind("<Enter>",              lambda e: self.pdf_canvas.focus_set())
        self.pdf_canvas.create_text(20, 20, text="（PDF を開くと表示されます）",
                                    anchor="nw", fill="gray50")

        # SIMA canvas
        self.sim_preview_frame = ttk.LabelFrame(right, text="SIMA プレビュー", padding=4)

        sim_cf = ttk.Frame(self.sim_preview_frame)
        sim_cf.pack(fill="both", expand=True)
        self.sim_canvas = tk.Canvas(sim_cf, bg="white", highlightthickness=0)
        self.sim_canvas.pack(side="left", fill="both", expand=True)
        sim_ys = ttk.Scrollbar(sim_cf, orient="vertical",   command=self.sim_canvas.yview)
        sim_ys.pack(side="right", fill="y")
        sim_xs = ttk.Scrollbar(self.sim_preview_frame, orient="horizontal", command=self.sim_canvas.xview)
        sim_xs.pack(fill="x")
        self.sim_canvas.configure(yscrollcommand=sim_ys.set, xscrollcommand=sim_xs.set)

        self.sim_canvas.bind("<ButtonPress-1>",      self.on_sim_pan_start)
        self.sim_canvas.bind("<B1-Motion>",          self.on_sim_pan_move)
        self.sim_canvas.bind("<ButtonRelease-1>",    self.on_sim_pan_end)
        self.sim_canvas.bind("<MouseWheel>",         self.on_sim_zoom_wheel)
        self.sim_canvas.bind("<Control-MouseWheel>", self.on_sim_zoom_wheel)
        self.sim_canvas.bind("<Button-4>",           self.on_sim_zoom_wheel_linux)
        self.sim_canvas.bind("<Button-5>",           self.on_sim_zoom_wheel_linux)
        self.sim_canvas.create_text(20, 20, text="（SIMA を開くと表示されます）",
                                    anchor="nw", fill="gray50")

        self.update_preview_mode()

    # -------------------------
    # Mode switch
    # -------------------------
    def update_preview_mode(self):
        if self.preview_mode_var.get() == "pdf":
            self.sim_preview_frame.pack_forget()
            self.pdf_preview_frame.pack(fill="both", expand=True)
        else:
            self.pdf_preview_frame.pack_forget()
            self.sim_preview_frame.pack(fill="both", expand=True)

    # -------------------------
    # File open
    # -------------------------
    def pick_pdf_dialog(self):
        p = filedialog.askopenfilename(title="PDF を選択",
                                       filetypes=[("PDF", "*.pdf"), ("All files", "*.*")])
        if p: self.load_pdf(p)

    def pick_sim_dialog(self):
        p = filedialog.askopenfilename(title="SIMA を選択",
                                       filetypes=[("SIMA", "*.sim;*.sima"), ("All files", "*.*")])
        if p: self.load_sim(p)

    def on_drop_pdf(self, event):
        paths = [p for p in norm_paths_from_dnd(event.data)
                 if os.path.isfile(p) and p.lower().endswith(".pdf")]
        if paths: self.load_pdf(paths[0])

    def on_drop_sim(self, event):
        paths = [p for p in norm_paths_from_dnd(event.data)
                 if os.path.isfile(p) and os.path.splitext(p.lower())[1] in (".sim", ".sima")]
        if paths: self.load_sim(paths[0])

    def load_pdf(self, path: str):
        try:
            if self._pdf_doc:
                self._pdf_doc.close()
            self._pdf_doc = fitz.open(path)
            self._pdf_page_index = 0
            self._pdf_cache = {}
            self._pdf_zoom = 1.5
            self._pdf_offx = 20.0
            self._pdf_offy = 20.0
            self.file_info_label.config(text=f"PDF: {os.path.basename(path)}")
            self.preview_mode_var.set("pdf")
            self.update_preview_mode()
            self.after(50, self.pdf_fit_to_canvas)
        except Exception as e:
            messagebox.showerror("PDF 読込エラー", str(e))

    def load_sim(self, path: str):
        try:
            pts, polys = parse_sima_file(path)
            self.sima_points = pts
            self.sima_polylines = polys
            self._compute_sim_bbox()
            self.file_info_label.config(text=f"SIMA: {os.path.basename(path)}")
            self.preview_mode_var.set("sim")
            self.update_preview_mode()
            self.after(50, self.sim_fit_view)
        except Exception as e:
            messagebox.showerror("SIMA 読込エラー", str(e))

    # -------------------------
    # PDF rendering
    # -------------------------
    def _get_pdf_base_image(self) -> Optional[Image.Image]:
        if not self._pdf_doc: return None
        key = (self._pdf_page_index, round(self._pdf_zoom, 4))
        img = self._pdf_cache.get(key)
        if img is not None: return img
        page = self._pdf_doc[self._pdf_page_index]
        mat = fitz.Matrix(self._pdf_zoom, self._pdf_zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        self._pdf_cache[key] = img
        if len(self._pdf_cache) > 8:
            del self._pdf_cache[next(iter(self._pdf_cache))]
        return img

    def pdf_fit_to_canvas(self):
        if not self._pdf_doc: return
        self.pdf_canvas.update_idletasks()
        cw = max(1, self.pdf_canvas.winfo_width())
        ch = max(1, self.pdf_canvas.winfo_height())
        rect = self._pdf_doc[self._pdf_page_index].rect
        zx = (cw - 20) / max(1.0, rect.width)
        zy = (ch - 20) / max(1.0, rect.height)
        self._pdf_zoom = max(self._pdf_min_zoom, min(self._pdf_max_zoom, min(zx, zy)))
        self._pdf_cache.pop((self._pdf_page_index, round(self._pdf_zoom, 4)), None)
        img = self._get_pdf_base_image()
        if img:
            self._pdf_offx = (cw - img.width) / 2
            self._pdf_offy = (ch - img.height) / 2
        self.render_pdf_preview()

    def pdf_prev_page(self):
        if not self._pdf_doc: return
        self._pdf_page_index = max(0, self._pdf_page_index - 1)
        self.pdf_fit_to_canvas()

    def pdf_next_page(self):
        if not self._pdf_doc: return
        self._pdf_page_index = min(len(self._pdf_doc) - 1, self._pdf_page_index + 1)
        self.pdf_fit_to_canvas()

    def pdf_zoom_in(self):
        self.pdf_zoom_at(self.pdf_canvas.winfo_width()/2, self.pdf_canvas.winfo_height()/2, 1.15)

    def pdf_zoom_out(self):
        self.pdf_zoom_at(self.pdf_canvas.winfo_width()/2, self.pdf_canvas.winfo_height()/2, 1/1.15)

    def pdf_zoom_at(self, cx: float, cy: float, factor: float):
        if not self._pdf_doc: return
        old_z = self._pdf_zoom
        new_z = max(self._pdf_min_zoom, min(self._pdf_max_zoom, old_z * factor))
        if abs(new_z - old_z) < 1e-9: return
        self._pdf_offx = cx - (cx - self._pdf_offx) * (new_z / old_z)
        self._pdf_offy = cy - (cy - self._pdf_offy) * (new_z / old_z)
        self._pdf_zoom = new_z
        self.render_pdf_preview()

    def on_pdf_pan_start(self, event):
        self.pdf_canvas.focus_set()
        self._pdf_pan_start  = (event.x, event.y)
        self._pdf_pan_origin = (event.x, event.y)
        self._pdf_pan_moved  = False

    def on_pdf_pan_move(self, event):
        if not self._pdf_pan_start: return
        x0, y0 = self._pdf_pan_start
        ox, oy = self._pdf_pan_origin
        if abs(event.x - ox) > 4 or abs(event.y - oy) > 4:
            self._pdf_pan_moved = True
        self._pdf_offx += event.x - x0
        self._pdf_offy += event.y - y0
        self._pdf_pan_start = (event.x, event.y)
        self.render_pdf_preview()

    def on_pdf_pan_end(self, event):
        self._pdf_pan_start  = None
        self._pdf_pan_origin = None
        self._pdf_pan_moved  = False

    def on_pdf_wheel(self, event):
        if not self._pdf_doc: return "break"
        ctrl  = bool(getattr(event, "state", 0) & 0x0004)
        shift = bool(getattr(event, "state", 0) & 0x0001)
        if ctrl:
            self.pdf_zoom_at(event.x, event.y, 1.10 if event.delta > 0 else 1/1.10)
            return "break"
        step = 40 if event.delta > 0 else -40
        if shift: self._pdf_offx += step
        else:     self._pdf_offy += step
        self.render_pdf_preview()
        return "break"

    def on_pdf_zoom_wheel(self, event):
        return self.on_pdf_wheel(event)

    def on_pdf_wheel_linux(self, event):
        if not self._pdf_doc: return "break"
        up   = (event.num == 4)
        ctrl = bool(getattr(event, "state", 0) & 0x0004)
        if ctrl:
            self.pdf_zoom_at(event.x, event.y, 1.10 if up else 1/1.10)
            return "break"
        self._pdf_offy += 40 if up else -40
        self.render_pdf_preview()
        return "break"

    def render_pdf_preview(self):
        self.pdf_canvas.delete("all")
        if not self._pdf_doc:
            self.pdf_canvas.create_text(20, 20, text="（PDF を開くと表示されます）",
                                        anchor="nw", fill="gray50")
            self.pdf_page_label.config(text="page: -/-")
            self.pdf_zoom_label.config(text="zoom: -")
            return

        self.pdf_page_label.config(text=f"page: {self._pdf_page_index+1}/{len(self._pdf_doc)}")
        self.pdf_zoom_label.config(text=f"zoom: {int(self._pdf_zoom*100)}%")

        img = self._get_pdf_base_image()
        if img is None: return

        tkimg = ImageTk.PhotoImage(img)
        self._pdf_preview_tk = tkimg
        self.pdf_canvas.create_image(self._pdf_offx, self._pdf_offy, image=tkimg, anchor="nw")
        self.pdf_canvas.create_text(
            10, 10, text="Ctrl+ホイール: ズーム  ドラッグ / ホイール: スクロール",
            anchor="nw", fill="gray30")

    # -------------------------
    # SIMA rendering
    # -------------------------
    def _compute_sim_bbox(self):
        if not self.sima_points:
            self._sim_bbox = None; return
        xs = [v[0] for v in self.sima_points.values()]
        ys = [v[1] for v in self.sima_points.values()]
        self._sim_bbox = (min(xs), min(ys), max(xs), max(ys))

    def sim_zoom_100(self):
        self._sim_zoom = 1.0
        self._sim_offx = 40
        self._sim_offy = 40
        self.render_sim_view()

    def sim_fit_view(self):
        self.sim_canvas.update_idletasks()
        cw = max(1, self.sim_canvas.winfo_width())
        ch = max(1, self.sim_canvas.winfo_height())
        if not self._sim_bbox:
            self.sim_canvas.delete("all")
            self.sim_canvas.create_text(20, 20, text="（SIMA を開くと表示されます）",
                                        anchor="nw", fill="gray50")
            return
        min_n, min_e, max_n, max_e = self._sim_bbox
        margin = 40
        self._sim_zoom = max(self._sim_min_zoom,
                             min(self._sim_max_zoom,
                                 min((cw - 2*margin) / max(1e-9, max_e - min_e),
                                     (ch - 2*margin) / max(1e-9, max_n - min_n))))
        self._sim_offx = margin - min_e * self._sim_zoom
        self._sim_offy = margin + max_n * self._sim_zoom
        self.render_sim_view()

    def world_to_screen(self, nx: float, ey: float) -> Tuple[float, float]:
        return ey * self._sim_zoom + self._sim_offx, -nx * self._sim_zoom + self._sim_offy

    def screen_to_world(self, sx: float, sy: float) -> Tuple[float, float]:
        return -(sy - self._sim_offy) / self._sim_zoom, (sx - self._sim_offx) / self._sim_zoom

    def on_sim_pan_start(self, event):
        self._sim_pan_start  = (event.x, event.y)
        self._sim_pan_origin = (event.x, event.y)
        self._sim_pan_moved  = False

    def on_sim_pan_move(self, event):
        if not self._sim_pan_start: return
        x0, y0 = self._sim_pan_start
        ox, oy = self._sim_pan_origin
        if math.hypot(event.x - ox, event.y - oy) >= 5:
            self._sim_pan_moved = True
        self._sim_offx += event.x - x0
        self._sim_offy += event.y - y0
        self._sim_pan_start = (event.x, event.y)
        self.render_sim_view()

    def on_sim_pan_end(self, event):
        self._sim_pan_start  = None
        self._sim_pan_origin = None
        self._sim_pan_moved  = False

    def on_sim_zoom_wheel(self, event):
        if not self.sima_points: return
        mx, my = event.x, event.y
        wx0, wy0 = self.screen_to_world(mx, my)
        self._sim_zoom *= (1.10 if event.delta > 0 else 1/1.10)
        self._sim_zoom = max(self._sim_min_zoom, min(self._sim_max_zoom, self._sim_zoom))
        sx1, sy1 = self.world_to_screen(wx0, wy0)
        self._sim_offx += mx - sx1
        self._sim_offy += my - sy1
        self.render_sim_view()

    def on_sim_zoom_wheel_linux(self, event):
        if not self.sima_points: return
        mx, my = event.x, event.y
        wx0, wy0 = self.screen_to_world(mx, my)
        self._sim_zoom *= (1.10 if event.num == 4 else 1/1.10)
        self._sim_zoom = max(self._sim_min_zoom, min(self._sim_max_zoom, self._sim_zoom))
        sx1, sy1 = self.world_to_screen(wx0, wy0)
        self._sim_offx += mx - sx1
        self._sim_offy += my - sy1
        self.render_sim_view()

    def render_sim_view(self):
        self.sim_canvas.delete("all")
        if not self.sima_points:
            self.sim_canvas.create_text(20, 20, text="（SIMA を開くと表示されます）",
                                        anchor="nw", fill="gray50")
            self.sim_info_label.config(text="SIMA: -")
            return

        min_n, min_e, max_n, max_e = self._sim_bbox or (0, 0, 0, 0)
        p1 = self.world_to_screen(min_n, min_e)
        p2 = self.world_to_screen(max_n, max_e)
        self.sim_canvas.config(scrollregion=(
            min(p1[0], p2[0]) - 200, min(p1[1], p2[1]) - 200,
            max(p1[0], p2[0]) + 200, max(p1[1], p2[1]) + 200))

        for poly in self.sima_polylines:
            pts = []
            for name in poly:
                if name in self.sima_points:
                    sx, sy = self.world_to_screen(*self.sima_points[name][:2])
                    pts.extend([sx, sy])
            if len(pts) >= 4:
                self.sim_canvas.create_line(*pts, fill="navy", width=1.5)
                if len(pts) >= 6:
                    self.sim_canvas.create_line(pts[-2], pts[-1], pts[0], pts[1],
                                                fill="navy", width=1.5)

        r = 3
        for name, (nx, ey, _) in self.sima_points.items():
            sx, sy = self.world_to_screen(nx, ey)
            self.sim_canvas.create_oval(sx-r, sy-r, sx+r, sy+r,
                                        fill="black", outline="black")
            self.sim_canvas.create_text(sx+6, sy-6, text=name, anchor="sw",
                                        fill="black", font=("Meiryo", 9))

        self.sim_info_label.config(
            text=f"点: {len(self.sima_points)}  画地: {len(self.sima_polylines)}"
                 f"  zoom: {self._sim_zoom:.3f}x")


# -------------------------
# Entry point
# -------------------------
def main():
    app = ViewerApp()

    for arg in sys.argv[1:]:
        path = os.path.abspath(arg)
        if not os.path.isfile(path): continue
        ext = os.path.splitext(path.lower())[1]
        if ext == ".pdf":
            app.after(100, lambda p=path: app.load_pdf(p))
        elif ext in (".sim", ".sima"):
            app.after(100, lambda p=path: app.load_sim(p))

    app.mainloop()


if __name__ == "__main__":
    main()