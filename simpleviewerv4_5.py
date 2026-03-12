import os
import sys
import math
import tempfile
import time
from PIL import Image, ImageTk, ImageSequence, ImageDraw, ImageFont
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD


class ViewerApp(TkinterDnD.Tk):

    def __init__(self):
        super().__init__()

        self.title("SimpleViewer v4.5")
        self.geometry("1200x900")

        self.current_file_path = ""
        self._doc_type = "none"
        self._page_index = 0
        self._total_pages = 1

        self._zoom = 1.0
        self._offx = 0
        self._offy = 0

        self._tk_main_img = None
        self._preview_images = []
        self._active_doc = None

        self._drag_data = {"x":0,"y":0}

        # v4.5追加
        self._file_list = []
        self._file_index = -1
        self._sima_screen_pts = []

        self._build_ui()

    def _build_ui(self):

        nav = ttk.Frame(self,padding=5)
        nav.pack(side="top",fill="x")

        ttk.Button(nav,text="◀ 前へ",command=self.prev_page,width=8).pack(side="left",padx=5)

        self.page_label = ttk.Label(nav,text="0 / 0")
        self.page_label.pack(side="left",padx=15)

        ttk.Button(nav,text="次へ ▶",command=self.next_page,width=8).pack(side="left",padx=5)

        ttk.Button(nav,text="📸 画像保存",command=self.save_current_as_image,width=12).pack(side="left",padx=10)

        self.status_label = ttk.Label(nav,text="待機中",foreground="blue")
        self.status_label.pack(side="left",padx=20)

        ttk.Button(nav,text="開く",command=self.pick_file_dialog).pack(side="right",padx=10)

        paned = ttk.Panedwindow(self,orient="vertical")
        paned.pack(fill="both",expand=True)

        self.canvas = tk.Canvas(paned,bg="#1a1a1a",highlightthickness=0)
        paned.add(self.canvas,weight=3)

        self.text_area = tk.Text(paned,height=8)
        paned.add(self.text_area,weight=1)

        self.drop_target_register(DND_FILES)
        self.dnd_bind("<<Drop>>",self.on_drop)

        self.bind("<Configure>",self._on_window_resize)
        self.canvas.bind("<MouseWheel>",self._on_mouse_wheel)

        self.canvas.bind("<ButtonPress-1>",self._on_drag_start)
        self.canvas.bind("<B1-Motion>",self._on_drag_move)

        # v4.5追加
        self.bind("<Right>",self._next_file)
        self.bind("<Left>",self._prev_file)
        self.canvas.bind("<Button-1>",self._on_canvas_click)

    # -----------------------------
    # ファイル一覧
    # -----------------------------

    def _load_file_list(self,path):

        folder=os.path.dirname(path)

        exts=[
        ".pdf",".sim",".sima",
        ".jpg",".jpeg",".png",".bmp",".gif",".tif",".tiff",
        ".txt",".py",".csv",".log",
        ".doc",".docx",".xls",".xlsx"
        ]

        files=[]

        for f in sorted(os.listdir(folder)):
            if os.path.splitext(f)[1].lower() in exts:
                files.append(os.path.join(folder,f))

        self._file_list=files

        if path in files:
            self._file_index=files.index(path)

    def _next_file(self,event=None):

        if not self._file_list:
            return

        if self._file_index < len(self._file_list)-1:
            self._file_index+=1
            self.load_file(self._file_list[self._file_index])

    def _prev_file(self,event=None):

        if not self._file_list:
            return

        if self._file_index>0:
            self._file_index-=1
            self.load_file(self._file_list[self._file_index])

    # -----------------------------
    # ファイルロード
    # -----------------------------

    def load_file(self,path):

        self.status_label.config(text="読込中")

        self.current_file_path=os.path.abspath(path)

        ext=os.path.splitext(path)[1].lower()

        self._page_index=0

        try:

            if ext in [".sim",".sima"]:

                pt_dict,polygons,text_lines=self._parse_sima(path)

                img=self._render_sima_to_image(pt_dict,polygons)

                self._preview_images=[img]

                self._doc_type="preview_imgs"

                self._total_pages=1

                self._display_text("\n".join(text_lines))

            elif ext in [".jpg",".jpeg",".png",".bmp",".gif",".tif",".tiff"]:

                with Image.open(path) as img:

                    self._preview_images=[frame.copy() for frame in ImageSequence.Iterator(img)]

                self._doc_type="preview_imgs"

                self._total_pages=len(self._preview_images)

                self._display_text(os.path.basename(path))

            elif ext==".pdf":

                import fitz

                self._active_doc=fitz.open(path)

                self._doc_type="paged"

                self._total_pages=len(self._active_doc)

            elif ext in [".txt",".py",".csv",".log"]:

                with open(path,"r",encoding="utf-8",errors="ignore") as f:

                    txt=f.read()

                self._doc_type="text_only"

                self._display_text(txt)

                self.canvas.delete("all")

            else:

                self._display_text("未対応形式")

            self.render_current(fit=True)

            self.page_label.config(text=f"{self._page_index+1}/{self._total_pages}")

            self.status_label.config(text="完了")

            self._load_file_list(self.current_file_path)

        except Exception as e:

            messagebox.showerror("Error",str(e))

    # -----------------------------
    # SIMA解析
    # -----------------------------

    def _parse_sima(self,path):

        pt_dict={}
        polygons=[]
        text_lines=[]

        cur_polygon=None

        with open(path,"r",encoding="cp932",errors="ignore") as f:

            for line in f:

                line=line.rstrip()

                text_lines.append(line)

                cols=[c.strip() for c in line.split(",")]

                rec=cols[0].upper() if cols else ""

                if rec=="A01":

                    seq=int(cols[1])

                    name=cols[2]

                    x=float(cols[3])

                    y=float(cols[4])

                    pt_dict[seq]=(name,x,y)

                elif rec=="D00":

                    cur_polygon={"name":cols[2],"pts":[]}

                    polygons.append(cur_polygon)

                elif rec=="B01":

                    seq=int(cols[1])

                    cur_polygon["pts"].append(seq)

                elif rec=="D99":

                    cur_polygon=None

        return pt_dict,polygons,text_lines

    # -----------------------------
    # SIMA描画
    # -----------------------------

    def _render_sima_to_image(self,pt_dict,polygons):

        pts=list(pt_dict.values())

        east=[p[2] for p in pts]

        north=[p[1] for p in pts]

        min_e=min(east);max_e=max(east)

        min_n=min(north);max_n=max(north)

        span_e=max_e-min_e

        span_n=max_n-min_n

        img_w=2000

        img_h=1500

        margin=100

        scale=min((img_w-2*margin)/span_e,(img_h-2*margin)/span_n)

        img=Image.new("RGB",(img_w,img_h),"white")

        draw=ImageDraw.Draw(img)

        font=ImageFont.load_default()

        self._sima_screen_pts=[]

        def tr(n,e):

            x=margin+(e-min_e)*scale

            y=img_h-(margin+(n-min_n)*scale)

            return x,y

        for i,poly in enumerate(polygons):

            seqs=poly["pts"]

            coords=[pt_dict[s] for s in seqs if s in pt_dict]

            if len(coords)<2:
                continue

            pts_xy=[tr(c[1],c[2]) for c in coords]

            pts_xy.append(pts_xy[0])

            draw.line(pts_xy,fill=(0,0,200),width=2)

        for name,n,e in pts:

            x,y=tr(n,e)

            draw.ellipse((x-3,y-3,x+3,y+3),fill=(0,0,0))

            draw.text((x+5,y-5),name,font=font,fill=(0,0,0))

            self._sima_screen_pts.append((name,x,y))

        return img

    # -----------------------------
    # クリック点
    # -----------------------------

    def _on_canvas_click(self,event):

        if not self._sima_screen_pts:
            return

        min_d=15
        found=None

        for name,x,y in self._sima_screen_pts:

            d=((x-event.x)**2+(y-event.y)**2)**0.5

            if d<min_d:

                min_d=d

                found=(name,x,y)

        if found:

            self.status_label.config(text=f"点 {found[0]}")

    # -----------------------------

    def render_current(self,fit=False):

        if self._doc_type=="preview_imgs":

            img=self._preview_images[self._page_index]

            self._tk_main_img=ImageTk.PhotoImage(img)

            self.canvas.delete("all")

            self.canvas.create_image(0,0,anchor="nw",image=self._tk_main_img)

        elif self._doc_type=="paged":

            import fitz

            page=self._active_doc[self._page_index]

            pix=page.get_pixmap()

            img=Image.frombytes("RGB",[pix.width,pix.height],pix.samples)

            self._tk_main_img=ImageTk.PhotoImage(img)

            self.canvas.delete("all")

            self.canvas.create_image(0,0,anchor="nw",image=self._tk_main_img)

    # -----------------------------

    def prev_page(self):

        if self._page_index>0:

            self._page_index-=1

            self.render_current()

    def next_page(self):

        if self._page_index<self._total_pages-1:

            self._page_index+=1

            self.render_current()

    # -----------------------------

    def save_current_as_image(self):

        if not self._preview_images:

            return

        path=filedialog.asksaveasfilename(defaultextension=".png")

        if path:

            self._preview_images[self._page_index].save(path)

    # -----------------------------

    def _display_text(self,txt):

        self.text_area.delete("1.0",tk.END)

        self.text_area.insert(tk.END,txt)

    def on_drop(self,event):

        path=event.data.strip("{}")

        if os.path.exists(path):

            self.load_file(path)

    def pick_file_dialog(self):

        p=filedialog.askopenfilename()

        if p:

            self.load_file(p)

    def _on_mouse_wheel(self,event):

        pass

    def _on_drag_start(self,event):

        self._drag_data["x"]=event.x

        self._drag_data["y"]=event.y

    def _on_drag_move(self,event):

        pass

    def _on_window_resize(self,event):

        pass


# -----------------------------
# 起動
# -----------------------------

if __name__=="__main__":

    app=ViewerApp()

    if len(sys.argv)>1:

        path=sys.argv[1]

        if os.path.exists(path):

            app.after(200,lambda:app.load_file(path))

    app.mainloop()