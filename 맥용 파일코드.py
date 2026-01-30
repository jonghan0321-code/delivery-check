import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np
import os
import shutil
from datetime import datetime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import pickle
import threading
import time
import glob
import math
import platform
import base64
import io
from PIL import Image, ImageTk 
# [System Font Setting] (Mac 전용 설정)
import platform
import matplotlib.pyplot as plt

# 맥은 AppleGothic을 써야 한글이 안 깨집니다.
plt.rc('font', family='AppleGothic') 
plt.rcParams['axes.unicode_minus'] = False

# GUI에서 사용할 폰트 변수 지정
GUI_FONT = "AppleGothic"

# Library 확인
try:
    import tkintermapview
    MAP_AVAILABLE = True
except ImportError:
    MAP_AVAILABLE = False

try:
    import openpyxl
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False

# --- 구성 ---
DATA_FILE = "Logistics_Master_DB.pkl"
BACKUP_DIR = "backup"
   # 지역들 분류
US_REGIONS = {
    'Northeast': ['CT', 'ME', 'MA', 'NH', 'RI', 'VT', 'NJ', 'NY', 'PA'],
    'Midwest':   ['IL', 'IN', 'MI', 'OH', 'WI', 'IA', 'KS', 'MN', 'MO', 'NE', 'ND', 'SD'],
    'Southeast': ['DE', 'FL', 'GA', 'MD', 'NC', 'SC', 'VA', 'DC', 'WV', 'AL', 'KY', 'MS', 'TN', 'AR', 'LA'],
    'Southwest': ['AZ', 'NM', 'OK', 'TX'],
    'West':      ['CO', 'ID', 'MT', 'NV', 'UT', 'WY'],
    'Pacific':   ['CA', 'OR', 'WA', 'AK', 'HI']
}
   # 지역들 위치(지도에 버블차트 중앙 표시하기)
REGION_CENTERS = {
    'Northeast': (42.0, -74.5), 'Midwest': (41.5, -92.0), 'Southeast': (33.5, -84.0),
    'Southwest': (32.5, -100.0), 'West': (40.0, -112.0), 'Pacific': (37.0, -120.0)
}
   # 지역별 색깔 지정
REGION_COLORS = {
    'Northeast': "#4E79A7", 'Midwest': "#F28E2B", 'Southeast': "#E15759",
    'Southwest': "#76B7B2", 'West': "#59A14F", 'Pacific': "#EDC948", 'Others/Unknown': "#BAB0AC"
}
    # 각 Status가 뭘 뜻하는지 매칭
STATUS_MAP = {
    'Submitted': 'Pre-Transit', 'Requested': 'Pre-Transit', 'Scheduled': 'Pre-Transit',
    'Picked UP': 'In-Transit', 'In Transit': 'In-Transit', 'Out For Delivery': 'In-Transit',
    'Delivered': 'Completed',
    'Return to Shipper': 'Exception', 'Lost': 'Exception', 'Attempted': 'Exception',
    'Cancelled': 'Exception', 'Partial': 'Exception'
}

TRANSPARENT_ICON_B64 = "R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7"

# --- [Helper] 스크롤 가능하게 바꿈 ---
class ScrollableFrame(ttk.Frame):
    def __init__(self, container, *args, **kwargs):
        super().__init__(container, *args, **kwargs)
        self.canvas = tk.Canvas(self, bg="#F1F5F9", highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas, style="TFrame")

        self.window_id = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.scrollable_frame.bind("<Configure>", self._on_frame_configure)
        
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        self.scrollable_frame.bind('<Enter>', self._bind_mouse)
        self.scrollable_frame.bind('<Leave>', self._unbind_mouse)

    def _on_canvas_configure(self, event):
        self.canvas.itemconfig(self.window_id, width=event.width)

    def _on_frame_configure(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _bind_mouse(self, event=None):
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind_all("<Button-4>", self._on_mousewheel)
        self.canvas.bind_all("<Button-5>", self._on_mousewheel)

    def _unbind_mouse(self, event=None):
        self.canvas.unbind_all("<MouseWheel>")
        self.canvas.unbind_all("<Button-4>")
        self.canvas.unbind_all("<Button-5>")

    def _on_mousewheel(self, event):
        if self.winfo_exists():
            if event.num == 5 or event.delta < 0:
                self.canvas.yview_scroll(1, "units")
            elif event.num == 4 or event.delta > 0:
                self.canvas.yview_scroll(-1, "units")

# --- [2] 작업 창 ---
class ProgressWindow(tk.Toplevel):
    def __init__(self, parent, title="Processing..."):
        super().__init__(parent)
        self.title(title)
        self.geometry("450x200")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        self.configure(bg="#ffffff")
        
        # 중앙 창
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (w // 2)
        y = (self.winfo_screenheight() // 2) - (h // 2)
        self.geometry(f'{w}x{h}+{x}+{y}')

        self.lbl_title = tk.Label(self, text="Please Wait...", font=("Arial", 12, "bold"), bg="#ffffff", fg="#1E293B")
        self.lbl_title.pack(pady=(30, 5))
        self.lbl_status = tk.Label(self, text="Initializing...", font=("Arial", 10), bg="#ffffff", fg="#3B82F6")
        self.lbl_status.pack(pady=(0, 10))
        self.progress = ttk.Progressbar(self, orient="horizontal", length=350, mode="determinate")
        self.progress.pack(pady=5)
        self.lbl_percent = tk.Label(self, text="0%", font=("Arial", 10, "bold"), bg="#ffffff", fg="#64748B")
        self.lbl_percent.pack(pady=5)

    def update_progress_safe(self, step, total_steps, message):
        """Called from thread. Updates UI safely via after()."""
        pct = int((step / total_steps) * 100) if total_steps > 0 else 0
        self.after(0, lambda: self._update_ui(pct, message))

    def _update_ui(self, pct, message):
        if not self.winfo_exists(): return 
        try:
            self.progress['value'] = pct
            self.lbl_percent.config(text=f"{pct}%")
            self.lbl_status.config(text=message)
            self.update_idletasks() # 렉 안걸리게
        except: pass

    def close(self): self.destroy()

# --- [3] 데이터 관리
class DataManager:
    def __init__(self):
        self.df = pd.DataFrame()
        os.makedirs(BACKUP_DIR, exist_ok=True)
        self.load_data()
         # 데이터 불러오기
    def load_data(self):
        if os.path.exists(DATA_FILE):
            try:
                with open(DATA_FILE, 'rb') as f: self.df = pickle.load(f)
                self.repair_data()
            except: self.df = pd.DataFrame()
        else: self.df = pd.DataFrame()
         #리드타임 계산하기
    def calculate_lead_time_row(self, row):
        delivered = row.get('Delivered Date')
        if pd.isna(delivered): return np.nan
        picked_up = row.get('Picked Up Date')
        req_date = row.get('Pickup Req Date')
        if pd.notna(picked_up):
            lt = (delivered - picked_up).days
            if lt >= 0: return lt
        if pd.notna(req_date):
            lt = (delivered - req_date).days
            if lt >= 0: return lt
        return np.nan
       # 데이터 전처리하기
    def repair_data(self):
        if self.df.empty: return
        date_cols = ['Created Dt', 'Pickup Req Date', 'Pickup Appt Date', 'Picked Up Date', 'Delivered Date']
        for col in date_cols:
            if col in self.df.columns:
                self.df[col] = pd.to_datetime(self.df[col], errors='coerce').dt.normalize()
        self.df['Lead_Time_Days'] = self.df.apply(self.calculate_lead_time_row, axis=1)
        if 'Created Dt' in self.df.columns:
            self.df['Year'] = self.df['Created Dt'].dt.year
            self.df['Quarter'] = self.df['Created Dt'].dt.quarter
            self.df['Month_Num'] = self.df['Created Dt'].dt.month
            self.df['Day_of_Week'] = self.df['Created Dt'].dt.day_name()
            self.df['Week_Num'] = self.df['Created Dt'].dt.isocalendar().week
            self.df['Year_Month'] = self.df['Created Dt'].dt.to_period('M').astype(str)
            
        self.df['Calc_Pickup_Date'] = self.df.get('Picked Up Date', pd.Series(dtype='datetime64[ns]')).combine_first(self.df.get('Pickup Appt Date', pd.Series(dtype='datetime64[ns]'))).combine_first(self.df.get('Pickup Req Date', pd.Series(dtype='datetime64[ns]')))
        self.df['Origin_State'] = self.df.get('Pickup State', 'Missing').fillna('Missing')
        self.df['Station'] = self.df.get('Dest Station', 'Unknown').fillna('Unknown')
        self._apply_mappings()
      # 필요한 데이터 매칭하기
    def _apply_mappings(self):
        def get_region(state):
            st = str(state).strip().upper() 
            for r, s in US_REGIONS.items():
                if st in s: return r
            return 'Other'
        if 'Dest State' in self.df.columns: self.df['Region'] = self.df['Dest State'].apply(get_region)
        if 'Status' in self.df.columns: self.df['Status_Group'] = self.df['Status'].apply(lambda s: STATUS_MAP.get(str(s).strip(), "Unknown"))

    def save_data(self):
        with open(DATA_FILE, 'wb') as f: pickle.dump(self.df, f)

    def backup_data(self):
        if os.path.exists(DATA_FILE): shutil.copy(DATA_FILE, os.path.join(BACKUP_DIR, f"master_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pkl"))
     # 날짜 데이터 가져오는 로직
    def process_file(self, file_path):
        try:
            df = pd.read_excel(file_path)
            df.columns = [c.strip() for c in df.columns]
            df['Source_File'] = os.path.basename(file_path)
            date_cols = ['Created Dt', 'Pickup Req Date', 'Pickup Appt Date', 'Picked Up Date', 'Delivered Date']
            for col in date_cols:
                if col in df.columns: df[col] = pd.to_datetime(df[col], errors='coerce').dt.normalize()
            df['Lead_Time_Days'] = df.apply(self.calculate_lead_time_row, axis=1)
            df['Calc_Pickup_Date'] = df.get('Picked Up Date', pd.Series(dtype='datetime64[ns]')).combine_first(df.get('Pickup Appt Date', pd.Series(dtype='datetime64[ns]'))).combine_first(df.get('Pickup Req Date', pd.Series(dtype='datetime64[ns]')))
            if 'Created Dt' in df.columns:
                df['Year'] = df['Created Dt'].dt.year
                df['Quarter'] = df['Created Dt'].dt.quarter
                df['Month_Num'] = df['Created Dt'].dt.month
                df['Day_of_Week'] = df['Created Dt'].dt.day_name()
                df['Week_Num'] = df['Created Dt'].dt.isocalendar().week
                df['Year_Month'] = df['Created Dt'].dt.to_period('M').astype(str)
            # 지역 가져오는 로직
            def get_region(state):
                st = str(state).strip().upper() 
                for r, s in US_REGIONS.items():
                    if st in s: return r
                return 'Other'
            if 'Dest State' in df.columns: df['Region'] = df['Dest State'].apply(get_region)
            if 'Status' in df.columns: df['Status_Group'] = df['Status'].apply(lambda s: STATUS_MAP.get(str(s).strip(), "Unknown"))
            df['Origin_State'] = df.get('Pickup State', 'Missing').fillna('Missing')
            df['Station'] = df.get('Dest Station', 'Unknown').fillna('Unknown')
            return df
        except Exception as e:
            print(f"Error: {e}")
            return pd.DataFrame()

    def merge_dataframe(self, new_df):
        if self.df.empty: self.df = new_df
        else:
            combined = pd.concat([self.df, new_df])
            self.df = combined.drop_duplicates(subset='PL No', keep='last').reset_index(drop=True)
        return True

# --- [4] UI이미지 설정 로직 ---
class LogisticsApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("KW Logistics Partner System")
        self.geometry("1600x950")
        self.configure(bg="#F1F5F9")
        
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.data_manager = DataManager()
        self.current_export_tables = {}
        self.current_export_figures = {}
        self.client_export_fig = None
        self.last_update_date = "Never"
        self.client_others_info = ""
        self.current_page_id = None
        
        try:
            img_data = base64.b64decode(TRANSPARENT_ICON_B64)
            self.invisible_icon = ImageTk.PhotoImage(Image.open(io.BytesIO(img_data)))
        except: self.invisible_icon = None
        
        self.setup_window_icon()

        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TFrame", background="#F1F5F9")
        style.configure("Card.TFrame", background="#FFFFFF", relief="flat", borderwidth=0)
        style.configure("Sidebar.TFrame", background="#1E293B")
        style.configure("TNotebook", background="#F1F5F9", borderwidth=0, padding=0) # Removing Padding
        style.configure("TNotebook.Tab", font=('AppleGothic', 10), padding=[15, 5], background="#E2E8F0", foreground="#64748B")
        style.map("TNotebook.Tab", 
                  background=[("selected", "#3B82F6"), ("active", "#60A5FA")], 
                  foreground=[("selected", "white"), ("active", "white")],
                  font=[("selected", ('AppleGothic', 11, 'bold'))],
                  padding=[("selected", [20, 8])])

        style.configure("TButton", font=('Arial', 10), padding=5)

        # 메인 창
        self.sidebar = ttk.Frame(self, style="Sidebar.TFrame", width=260)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        self.main_container = ttk.Frame(self, style="TFrame")
        self.main_container.pack(side="right", fill="both", expand=True)

        self.setup_sidebar()
        
        # 상단 빈 공간 줄이기
        self.header = tk.Frame(self.main_container, bg="#F1F5F9", height=40)
        self.header.pack(side="top", fill="x", padx=20, pady=(10, 0))
        
        header_right = tk.Frame(self.header, bg="#F1F5F9")
        header_right.pack(side="right")
        self.lbl_update = tk.Label(header_right, text=f"Last updated: {self.last_update_date}", font=("Arial", 9), bg="#F1F5F9", fg="#64748B")
        self.lbl_update.pack(side="left", padx=15)
        tk.Button(header_right, text="+ Add Data File", font=("Arial", 10, "bold"), bg="#3B82F6", fg="white", 
                  relief="flat", padx=15, pady=6, cursor="hand2", command=self.run_single_import).pack(side="left")

        # 콘텐츠 영역
        self.content_area = ttk.Frame(self.main_container, style="TFrame")
        self.content_area.pack(side="bottom", fill="both", expand=True, padx=20, pady=5)
        
        self.pages = {}
        for page_id in ["client", "internal"]:
            f = ScrollableFrame(self.content_area)
            f.grid(row=0, column=0, sticky="nsew")
            self.pages[page_id] = f
        self.content_area.grid_rowconfigure(0, weight=1)
        self.content_area.grid_columnconfigure(0, weight=1)

        self.create_client_page(self.pages["client"].scrollable_frame)
        self.create_internal_page(self.pages["internal"].scrollable_frame)

        self.show_page("client")
    # 창에 뜨는 아이콘 변경
    def setup_window_icon(self):
        icon_path = "logo.png"
        if os.path.exists(icon_path):
            try:
                # 맥 호환 아이콘 설정 방식
                img = ImageTk.PhotoImage(Image.open(icon_path))
                self.wm_iconphoto(True, img)
            except Exception as e:
                print(f"Icon Error: {e}")
    # 사이드바 설정
    def setup_sidebar(self):
        tk.Label(self.sidebar, text="KW Partner", font=("Arial Black", 18), bg="#1E293B", fg="#3B82F6", pady=30).pack(side="top")
        menu_frame = tk.Frame(self.sidebar, bg="#1E293B")
        menu_frame.pack(side="top", fill="x")
        self.menu_btns = {}
        menus = [("client", "👥  Client View"), ("internal", "📊  Internal Analysis")]
        for mid, mtext in menus:
            btn = tk.Button(menu_frame, text=mtext, font=("AppleGothic", 11, "bold"), bg="#1E293B", fg="#94A3B8",
                            activebackground="#334155", activeforeground="white", relief="flat", bd=0, padx=25, pady=15, 
                            anchor="w", cursor="hand2", command=lambda m=mid: self.show_page(m))
            btn.pack(fill="x")
            self.menu_btns[mid] = btn

        self.sidebar_bottom = tk.Frame(self.sidebar, bg="#1E293B")
        self.sidebar_bottom.pack(side="bottom", fill="x", pady=20)
        self.btn_sidebar_export = tk.Button(self.sidebar_bottom, text="💾 Export Report", font=("AppleGothic", 10, "bold"), 
                                            bg="#3B82F6", fg="white", relief="flat", padx=20, pady=10, 
                                            cursor="hand2", command=self.run_export_thread)
        self.btn_sidebar_save = tk.Button(self.sidebar_bottom, text="📷 Save Graph", font=("AppleGothic", 10, "bold"), 
                                          bg="#3B82F6", fg="white", relief="flat", padx=20, pady=10, 
                                          cursor="hand2", command=self.save_client_image)

    def show_page(self, page_id):
        self.pages[page_id].tkraise()
        for mid, btn in self.menu_btns.items():
            if mid == page_id: btn.configure(bg="#3B82F6", fg="white")
            else: btn.configure(bg="#1E293B", fg="#94A3B8")
        self.btn_sidebar_export.pack_forget()
        self.btn_sidebar_save.pack_forget()
        if page_id == "internal": self.btn_sidebar_export.pack(fill="x", padx=20)
        elif page_id == "client": self.btn_sidebar_save.pack(fill="x", padx=20)

    # --- Client 창 설정 로직 ---
    def save_client_image(self):
        if not self.client_export_fig: return messagebox.showwarning("Empty", "No graph to save.")
        path = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG Image", "*.png")])
        if path: self.client_export_fig.savefig(path, dpi=150, bbox_inches='tight'); messagebox.showinfo("Done", "Saved.")

    def create_client_page(self, parent):
        main_f = ttk.Frame(parent)
        main_f.pack(fill="both", expand=True)
        
        filter_card = tk.Frame(main_f, bg="white", highlightthickness=1, highlightbackground="#E2E8F0", pady=10, padx=20)
        filter_card.pack(fill="x", pady=(0, 15))
        
        vars = {'Year': tk.StringVar(), 'Quarter': tk.StringVar(), 'BillTo': tk.StringVar(), 'SubType': tk.StringVar(), 'PLType': tk.StringVar()}
        f_grid = tk.Frame(filter_card, bg="white"); f_grid.pack(fill="x")
        def add_filter(label, var, col, width=15):
            tk.Label(f_grid, text=label, font=("AppleGothic", 9, "bold"), bg="white", fg="#64748B").grid(row=0, column=col*2, padx=10, sticky='e')
            cb = ttk.Combobox(f_grid, textvariable=var, width=width); cb.grid(row=0, column=col*2+1, padx=5, sticky='w'); return cb
        cb_year = add_filter("Year:", vars['Year'], 0, 8)
        cb_q = add_filter("Q:", vars['Quarter'], 1, 5); cb_q['values'] = ["All", "1", "2", "3", "4"]; cb_q.current(0)
        cb_bill = add_filter("Bill To:", vars['BillTo'], 2, 15)
        cb_sub = add_filter("Sub Type:", vars['SubType'], 3, 15)
        cb_pl = add_filter("PL Type:", vars['PLType'], 4, 15)

        content_f = tk.Frame(main_f, bg="#F1F5F9"); content_f.pack(fill="both", expand=True)
        map_card = tk.Frame(content_f, bg="white", highlightthickness=1, highlightbackground="#E2E8F0")
        map_card.pack(side="left", fill="both", expand=True, padx=(0, 10))
        
        if MAP_AVAILABLE:
            map_w = tkintermapview.TkinterMapView(map_card, corner_radius=0)
            map_w.pack(fill="both", expand=True, padx=2, pady=2); map_w.set_position(39.8283, -98.5795); map_w.set_zoom(4)
        else: map_w = None

        rep_card = tk.Frame(content_f, bg="white", highlightthickness=1, highlightbackground="#E2E8F0", width=550)
        rep_card.pack(side="right", fill="y")
        report_frame = tk.Frame(rep_card, bg="white"); report_frame.pack(fill="both", expand=True, padx=10, pady=10)

        def load_c(e):
            df = self.data_manager.df
            if df.empty: return
            if 'Year' in df: cb_year['values'] = ["All"] + sorted([str(int(y)) for y in df['Year'].dropna().unique()], reverse=True)
            if 'Bill To' in df: cb_bill['values'] = ["All"] + sorted(list(df['Bill To'].dropna().unique()))
            if 'Sub Type' in df: cb_sub['values'] = ["All"] + sorted(list(df['Sub Type'].dropna().unique()))
            if 'PL Type' in df: cb_pl['values'] = ["All"] + sorted(list(df['PL Type'].dropna().unique()))
        filter_card.bind("<Enter>", load_c)
        # 버블차트 그리기
        def draw_bubble_map(widget, region_data):
            widget.delete_all_marker(); widget.delete_all_path(); widget.delete_all_polygon()
            if region_data.empty: return
            max_vol = region_data.max()
            for region, vol in region_data.items():
                if region in REGION_CENTERS:
                    lat, lon = REGION_CENTERS[region]
                    radius = 1.5 + (vol / max_vol) * 3.5
                    circle = []
                    for a in range(0, 361, 20):
                        r = math.radians(a)
                        circle.append((lat + radius * math.sin(r), lon + (radius * math.cos(r)) / math.cos(math.radians(lat))))
                    widget.set_polygon(circle, fill_color=REGION_COLORS.get(region, "#3B82F6"), outline_color="white", border_width=1)
                    if self.invisible_icon: widget.set_marker(lat, lon, text=f"{region}\n{vol:,}", icon=self.invisible_icon)
                    else: widget.set_marker(lat, lon, text=f"{region}\n{vol:,}")
        # 클라이언트 버블차트에 없는 것들 나타내기
        def show_others_popup():
            if not self.client_others_info: messagebox.showinfo("Info", "No data."); return
            pop = tk.Toplevel(self); pop.title("Details"); pop.geometry("600x400")
            txt = tk.Text(pop, font=("Arial", 14), wrap="word", padx=15, pady=15)
            vsb = ttk.Scrollbar(pop, orient="vertical", command=txt.yview); txt.configure(yscrollcommand=vsb.set)
            vsb.pack(side="right", fill="y"); txt.pack(side="left", fill="both", expand=True)
            txt.insert("1.0", self.client_others_info); txt.config(state="disabled")
        # 클라이언트 내보내기 기능 누를 시
        def run_client_report_thread():
            pw = ProgressWindow(self, title="Generating Report")
            def task():
                try:
                    pw.update_progress_safe(10, 100, "Loading Data...")
                    time.sleep(0.1) # UI Refresh
                    df = self.data_manager.df
                    if df.empty: self.after(0, lambda: messagebox.showinfo("Info", "No Data.")); return
                    pw.update_progress_safe(30, 100, "Filtering...")
                    cond = pd.Series([True] * len(df))
                    if vars['Year'].get() and vars['Year'].get() != "All": cond &= (df['Year'] == int(vars['Year'].get()))
                    if vars['Quarter'].get() and vars['Quarter'].get() != "All": cond &= (df['Quarter'] == int(vars['Quarter'].get()))
                    if vars['BillTo'].get() and vars['BillTo'].get() != "All": cond &= (df['Bill To'] == vars['BillTo'].get())
                    if vars['SubType'].get() and vars['SubType'].get() != "All": cond &= (df['Sub Type'] == vars['SubType'].get())
                    if vars['PLType'].get() and vars['PLType'].get() != "All": cond &= (df['PL Type'] == vars['PLType'].get())
                    filtered = df[cond].copy()
                    
                    pw.update_progress_safe(60, 100, "Calculating...")
                    time.sleep(0.1)
                    valid_lt = filtered[(filtered['Status'] == 'Delivered') & (filtered['Lead_Time_Days'] >= 0) & (filtered['Lead_Time_Days'] < 21)] if 'Lead_Time_Days' in filtered else pd.DataFrame()
                    valid_lt = valid_lt[~valid_lt['Region'].isin(['Other', 'Unknown'])]
                    other_data = filtered[(filtered['Region'].isin(['Other', 'Unknown'])) & (filtered['Status'] == 'Delivered')]
                    if not other_data.empty:
                        details = []
                        grouped = other_data.groupby(other_data['Dest State'].fillna('Unknown (Empty)'))
                        for state, group in grouped:
                            pl_list = group['PL No'].astype(str).tolist()
                            pl_str = ", ".join(pl_list[:50])
                            if len(pl_list) > 50: pl_str += f" ... (+{len(pl_list)-50} more)"
                            details.append(f"[{state}] Count: {len(group)}\nPL List: {pl_str}\n" + "-"*40)
                        self.client_others_info = "\n".join(details)
                    else: self.client_others_info = "No Data"

                    def update_ui():
                        for w in report_frame.winfo_children(): w.destroy()
                        plt.close('all')
                        avg_lt = valid_lt['Lead_Time_Days'].mean() if not valid_lt.empty else 0
                        fig, ax = plt.subplots(figsize=(5, 5.0)) 
                        fig.text(0.5, 0.96, f"Total: {len(filtered):,} | Avg LT: {avg_lt:.1f} Days", ha='center', fontsize=12, fontweight='bold', color='#004080')
                        
                        if not valid_lt.empty:
                            grp = valid_lt.groupby('Region')['Lead_Time_Days'].mean().sort_values()
                            colors = [REGION_COLORS.get(r, '#66b3ff') for r in grp.index]
                            
                            grp.plot(kind='barh', ax=ax, color=colors)
                            
                            # [핵심 수정] X축 범위를 최댓값의 1.3배로 늘려 레이블 공간 확보
                            ax.set_xlim(right=grp.max() * 1.3)
                            
                            ax.set_title("Avg Lead Time"); ax.set_xlabel("Days"); ax.set_ylabel("Region")
                            
                            # 레이블: 검정색, 폰트 10, 패딩 3
                            for c in ax.containers:
                                ax.bar_label(c, fmt='%.1f', padding=3, color='black', fontweight='bold', fontsize=10)
                            
                        self.client_export_fig = fig 
                        FigureCanvasTkAgg(fig, master=report_frame).draw(); FigureCanvasTkAgg(fig, master=report_frame).get_tk_widget().pack(fill='both', expand=True, padx=5, pady=5)
                        if map_w:
                            cnt = filtered['Region'].value_counts()
                            for x in ['Other','Unknown']: 
                                if x in cnt: cnt = cnt.drop(x)
                            draw_bubble_map(map_w, cnt)

                    pw.update_progress_safe(90, 100, "Rendering..."); self.after(0, update_ui)
                except Exception as e: self.after(0, lambda: messagebox.showerror("Error", str(e)))
                finally: self.after(0, pw.close)
            threading.Thread(target=task, daemon=True).start()

        tk.Button(filter_card, text="ℹ️ Others Detail", font=("AppleGothic", 10, "bold"), bg="#F8FAFC", fg="#64748B", 
                  relief="flat", padx=15, pady=8, command=show_others_popup).pack(side='left', padx=10)
        tk.Button(filter_card, text="🔍 Generate Insights", font=("AppleGothic", 10, "bold"), bg="#3B82F6", fg="white", 
                  relief="flat", padx=15, pady=8, command=run_client_report_thread).pack(side='right', padx=10)

    # 상세 분석창
    def create_internal_page(self, parent):
        # [Fix: Remove container padding]
        container = ttk.Frame(parent)
        container.pack(fill="both", expand=True, pady=0) # Removed vertical padding
        inner_note = ttk.Notebook(container)
        inner_note.pack(fill="both", expand=True, pady=0)
        self.create_analysis_tab_content(inner_note, "GPC-A Milkrun", lambda df: df[df['Sub Type'] == 'GPC-A Milkrun'])
        self.create_analysis_tab_content(inner_note, "Encompass MR", lambda df: df[df['Sub Type'] == 'Encompass MR'])
        self.create_analysis_tab_content(inner_note, "AmazonSend", lambda df: df[df['PL Type'] == 'AmazonSend'])
        self.create_analysis_tab_content(inner_note, "General", lambda df: df[~( (df['Sub Type'] == 'GPC-A Milkrun') | (df['Sub Type'] == 'Encompass MR') | (df['PL Type'] == 'AmazonSend') )])

    def create_analysis_tab_content(self, parent_nb, title, filter_func):
        page = ttk.Frame(parent_nb, style="TFrame") 
        parent_nb.add(page, text=f"  {title}  ")
        
        # [Fix: Moved up by reducing pady and setting to 0 where possible]
        filter_frame = tk.Frame(page, bg="white", highlightthickness=1, highlightbackground="#E2E8F0", pady=2)
        filter_frame.pack(fill="x", pady=(0, 2)) # Minimized gap between tab and filter
        
        vars = {'Year': tk.StringVar(), 'Quarter': tk.StringVar(), 'BillTo': tk.StringVar(), 'SubType': tk.StringVar(), 'PLType': tk.StringVar()}
        f_grid = tk.Frame(filter_frame, bg="white"); f_grid.pack(fill="x", padx=20)
        def add_filter(label, var, col):
            tk.Label(f_grid, text=label, font=("AppleGothic", 9, "bold"), bg="white", fg="#64748B").grid(row=0, column=col*2, padx=5, sticky='e')
            cb = ttk.Combobox(f_grid, textvariable=var, width=10); cb.grid(row=0, column=col*2+1, padx=2, sticky='w'); return cb
        cb_year = add_filter("Year:", vars['Year'], 0)
        cb_q = add_filter("Q:", vars['Quarter'], 1); cb_q['values'] = ["All", "1", "2", "3", "4"]; cb_q.current(0)
        cb_bill = add_filter("Bill To:", vars['BillTo'], 2)
        cb_sub = add_filter("Sub Type:", vars['SubType'], 3)
        cb_pl = add_filter("PL Type:", vars['PLType'], 4)
        tk.Button(filter_frame, text="🔍 Run Analysis", font=("AppleGothic", 10, "bold"), bg="#3B82F6", fg="white", 
                  relief="flat", padx=20, pady=5, command=lambda: self.run_internal_analysis(title, filter_func, vars, summary_card_frame, viz_container, stats_frame)).pack(side="right", padx=10)

        content_frame = ttk.Frame(page); content_frame.pack(fill='both', expand=True)
        stats_frame = tk.Frame(content_frame, bg="#F1F5F9"); stats_frame.pack(fill='x', pady=2)
        summary_card_frame = tk.Frame(content_frame, bg="#F1F5F9"); summary_card_frame.pack(fill='x', pady=(0, 5))
        viz_container = tk.Frame(content_frame, bg="white", highlightthickness=1, highlightbackground="#E2E8F0")
        # 그래프 보여주는 버튼
        def toggle_visualization():
            if viz_container.winfo_ismapped(): viz_container.pack_forget(); btn_viz.config(text="📊 Show Graphs")
            else: viz_container.pack(fill='both', expand=True, pady=5); btn_viz.config(text="🔼 Hide Graphs")

        action_bar = tk.Frame(page, bg="#F1F5F9"); action_bar.pack(side="bottom", fill="x", pady=5)
        btn_viz = tk.Button(action_bar, text="📊 Show Graphs", font=("AppleGothic", 10, "bold"), bg="#F8FAFC", fg="#3B82F6", 
                            relief="flat", highlightthickness=1, highlightbackground="#3B82F6", padx=15, pady=8, command=toggle_visualization)
        btn_viz.pack(side='left')

        # 배송 오래걸린 것 보여주는 버튼
        def show_delayed_details_internal(data):
            if data.empty: return
            pop = tk.Toplevel(self); pop.title("Details"); pop.geometry("1100x600")
            tree = ttk.Treeview(pop, columns=("PL", "BillTo", "State", "Days", "Log", "File"), show='headings')
            vsb = ttk.Scrollbar(pop, orient="vertical", command=tree.yview); tree.configure(yscroll=vsb.set); vsb.pack(side='right', fill='y'); tree.pack(fill='both', expand=True)
            for c in ["PL", "BillTo", "State", "Days", "Log", "File"]: tree.heading(c, text=c)
            for _, r in data.sort_values('Lead_Time_Days', ascending=False).iterrows():
                tree.insert("", "end", values=(r.get('PL No'), r.get('Bill To'), r.get('Dest State'), int(r['Lead_Time_Days']), r.get('Last Log'), r.get('Source_File')))

        # 배송에 문제생긴 것 보여주는 버튼
        def show_issues_details(data):
            if data.empty: return
            pop = tk.Toplevel(self); pop.title("Issues"); pop.geometry("1100x600")
            tree = ttk.Treeview(pop, columns=("PL", "BillTo", "Status", "Log", "File"), show='headings')
            vsb = ttk.Scrollbar(pop, orient="vertical", command=tree.yview); tree.configure(yscroll=vsb.set); vsb.pack(side='right', fill='y'); tree.pack(fill='both', expand=True)
            for c in ["PL", "BillTo", "Status", "Log", "File"]: tree.heading(c, text=c)
            for _, r in data.sort_values('PL No', ascending=False).iterrows():
                tree.insert("", "end", values=(r.get('PL No'), r.get('Bill To'), r.get('Status'), r.get('Last Log'), r.get('Source_File')))

        def load_v(e):
            df = self.data_manager.df
            if df.empty: return
            t = filter_func(df)
            if not t.empty:
                cb_year['values'] = ["All"] + sorted([str(int(y)) for y in t['Year'].dropna().unique()], reverse=True)
                cb_bill['values'] = ["All"] + sorted(list(t['Bill To'].dropna().unique()))
                cb_sub['values'] = ["All"] + sorted(list(t['Sub Type'].dropna().unique()))
                cb_pl['values'] = ["All"] + sorted(list(t['PL Type'].dropna().unique()))
        filter_frame.bind("<Enter>", load_v)

        # 상세분석 진행 버튼 누르기
        def run_internal_analysis(title, filter_func, vars, summary_p, viz_p, stats_p):
            pw = ProgressWindow(self, title="Analyzing...")
            def task():
                try:
                    self.current_export_tables = {}; self.current_export_figures = {}
                    pw.update_progress_safe(10, 100, "Reading Data...")
                    time.sleep(0.1)
                    df = self.data_manager.df
                    if df.empty: return
                    base = filter_func(df)
                    
                    pw.update_progress_safe(30, 100, "Applying Filters...")
                    time.sleep(0.1)
                    cond = pd.Series([True] * len(base), index=base.index)
                    if vars['Year'].get() and vars['Year'].get() != "All": cond &= (base['Year'] == int(vars['Year'].get()))
                    if vars['Quarter'].get() and vars['Quarter'].get() != "All": cond &= (base['Quarter'] == int(vars['Quarter'].get()))
                    if vars['BillTo'].get() and vars['BillTo'].get() != "All": cond &= (base['Bill To'] == vars['BillTo'].get())
                    if vars['SubType'].get() and vars['SubType'].get() != "All": cond &= (base['Sub Type'] == vars['SubType'].get())
                    if vars['PLType'].get() and vars['PLType'].get() != "All": cond &= (base['PL Type'] == vars['PLType'].get())
                    
                    filtered = base[cond].copy()
                    
                    pw.update_progress_safe(50, 100, "Processing Statistics...")
                    viz_df = filtered.copy()
                    if 'Lead_Time_Days' in viz_df.columns: viz_df.loc[viz_df['Lead_Time_Days'] >= 21, 'Lead_Time_Days'] = np.nan
                    
                    # [데이터 추출]
                    long_lt = filtered[filtered['Lead_Time_Days'] >= 21].copy() if 'Lead_Time_Days' in filtered else pd.DataFrame()
                    issues = filtered[filtered['Status_Group'] == 'Exception'].copy() if 'Status_Group' in filtered else pd.DataFrame()
                    
                    # [핵심 수정] 정렬 로직 (Last Log 기준 내림차순)
                    # 'Last Log' 컬럼이 없으면 기존 방식(Lead Time 또는 PL No)으로 정렬
                    if 'Last Log' in filtered.columns:
                        if not long_lt.empty: 
                            long_lt = long_lt.sort_values(by='Last Log', ascending=False, na_position='last')
                        if not issues.empty: 
                            issues = issues.sort_values(by='Last Log', ascending=False, na_position='last')
                    else:
                        if not long_lt.empty: 
                            long_lt = long_lt.sort_values(by='Lead_Time_Days', ascending=False)
                        if not issues.empty: 
                            issues = issues.sort_values(by='PL No', ascending=False)

                    total = len(filtered); completed = len(filtered[filtered['Status_Group']=='Completed'])
                    active = len(filtered[filtered['Status_Group']=='In-Transit']); issue_cnt = len(issues)
                    avg_val = viz_df['Lead_Time_Days'].mean() if 'Lead_Time_Days' in viz_df else 0
                    
                    # Export 테이블 저장 (이미 정렬된 데이터프레임 사용)
                    summary_df = pd.DataFrame({'Metric':['Total','Active','Completed','Issues','LongDelay'], 'Value':[total, active, completed, issue_cnt, len(long_lt)]})
                    self.current_export_tables[f'Summary_{title}'] = summary_df
                    
                    if not issues.empty: self.current_export_tables[f'Issues_{title}'] = issues
                    if not long_lt.empty: self.current_export_tables[f'Delays_{title}'] = long_lt
                    
                    pw.update_progress_safe(75, 100, "Generating Graphs...")
                    time.sleep(0.1)
                    
                    def update_ui():
                        btn_viz.config(state='normal')
                        for w in stats_p.winfo_children(): w.destroy()
                        tk.Label(stats_p, text=f"Total: {len(filtered):,} | Avg LT: {avg_val:.1f} Days", font=("Arial", 14, "bold"), fg="#004080", bg="#F1F5F9").pack(side='left', padx=10)
                        
                        # 버튼에 연결되는 함수에도 정렬된 데이터(long_lt, issues)가 들어감
                        if not long_lt.empty: tk.Button(stats_p, text=f"🚨 Delays ({len(long_lt)})", bg="#ffcccc", command=lambda: show_delayed_details_internal(long_lt)).pack(side='right', padx=5)
                        if not issues.empty: tk.Button(stats_p, text=f"⚠️ Issues ({len(issues)})", bg="#fff3cd", command=lambda: show_issues_details(issues)).pack(side='right', padx=5)

                        for w in summary_p.winfo_children(): w.destroy()
                        metrics = [("Total Orders", total, "black"), ("Active (In-Transit)", active, "#007acc"), ("Completed", completed, "#28a745"), ("Issues (Exception)", issue_cnt, "#dc3545")]
                        for m_t, m_v, m_c in metrics:
                            card = tk.Frame(summary_p, bg="white", highlightthickness=1, highlightbackground="#ccc")
                            card.pack(side="left", fill="both", expand=True, padx=5)
                            tk.Label(card, text=m_t, font=("Arial", 10, "bold"), bg="white", fg="#555").pack(pady=(10, 5))
                            tk.Label(card, text=f"{m_v:,}", font=("Arial", 20, "bold"), bg="white", fg=m_c).pack(pady=(0, 10))
                        self.update_graphs(viz_p, viz_df, filtered, title)
                    self.after(0, update_ui)
                    pw.update_progress_safe(100, 100, "Done!")
                    time.sleep(0.2)
                except Exception as e: self.after(0, lambda: messagebox.showerror("Error", str(e)))
                finally: self.after(0, pw.close)
            threading.Thread(target=task, daemon=True).start()
        
        self.run_internal_analysis = run_internal_analysis

    def update_graphs(self, parent, viz_data, raw_data, title):
        for w in parent.winfo_children(): w.destroy()
        plt.close('all')
        nb = ttk.Notebook(parent); nb.pack(fill="both", expand=True)
        
        # 1. Trend (MoM) 그래프
        t_tab = ttk.Frame(nb); nb.add(t_tab, text=" 📈 Trend (MoM) ")
        if 'Month_Num' in viz_data.columns:
            f_top = ttk.Frame(t_tab); f_top.pack(fill='both', expand=True, pady=5)
            f_bot = ttk.Frame(t_tab); f_bot.pack(fill='x', expand=False, pady=5, padx=10)
            fig, ax = plt.subplots(figsize=(8, 2.5), facecolor='none')
            ax.set_facecolor('none')
            monthly = viz_data.groupby(['Year', 'Month_Num']).size().reset_index(name='PL No')
            pivot = monthly.pivot(index='Month_Num', columns='Year', values='PL No').fillna(0)
            pivot.plot(kind='bar', ax=ax, rot=0, width=0.8, colormap='viridis', alpha=0.8)
            
            # [수정] Y축 범위 확장 (상단 레이블 공간 확보)
            if not monthly.empty: ax.set_ylim(top=monthly['PL No'].max() * 1.25)
            
            ax.set_title("Monthly Order Trends", fontsize=12, fontweight='bold', pad=10)
            ax.set_xlabel("Month"); ax.set_ylabel("Volume")
            
            for c in ax.containers: ax.bar_label(c, fmt='%d', padding=2, fontsize=9, color='black')
            
            self.current_export_figures[f'Trend_Chart_{title}'] = fig
            FigureCanvasTkAgg(fig, master=f_top).draw(); FigureCanvasTkAgg(fig, master=f_top).get_tk_widget().pack(fill="both", expand=True)
            
            cols = ("Year", "Month", "Volume", "Prev Month", "Diff", "Growth(%)")
            tree = ttk.Treeview(f_bot, columns=cols, show='headings', height=6)
            for c in cols: tree.heading(c, text=c); tree.column(c, width=100, anchor='center')
            tree.pack(fill='x', expand=True)
            monthly = monthly.sort_values(['Year', 'Month_Num']); prev = None
            for _, r in monthly.iterrows():
                curr = int(r['PL No']); diff = curr - prev if prev else 0; pct = (diff/prev*100) if prev else 0.0
                tree.insert("", "end", values=(r['Year'], r['Month_Num'], curr, prev if prev else "-", f"{diff:+}", f"{pct:.1f}%" if prev else "-"))
                prev = curr

        # 2. Trend (WoW) - Year Filter & Fixed 1-52 Axis
        w_tab = ttk.Frame(nb); nb.add(w_tab, text=" 📉 Trend (WoW) ")
        
        # [컨트롤 패널] - Station 탭과 동일한 스타일
        ctrl_w = tk.Frame(w_tab, bg="#F1F5F9", pady=5)
        ctrl_w.pack(fill='x', padx=10)
        
        tk.Label(ctrl_w, text="Year:", font=("Arial", 10, "bold"), bg="#F1F5F9", fg="#333").pack(side='left', padx=(0, 5))
        
        # 연도 선택
        avail_years_w = sorted(list(raw_data['Year'].dropna().unique().astype(int)), reverse=True)
        year_var_w = tk.StringVar(value=str(avail_years_w[0]) if avail_years_w else "")
        cb_year_w = ttk.Combobox(ctrl_w, textvariable=year_var_w, values=avail_years_w, width=6, state="readonly")
        cb_year_w.pack(side='left')
        
        # 업데이트 버튼
        tk.Button(ctrl_w, text="🔄 Update", font=("AppleGothic", 9, "bold"), bg="#3B82F6", fg="white", 
                  relief="flat", padx=10, command=lambda: redraw_wow()).pack(side='left', padx=15)

        # 그래프 및 테이블 영역
        f_w_graph = ttk.Frame(w_tab); f_w_graph.pack(fill='both', expand=True, pady=5)
        f_w_table = ttk.Frame(w_tab); f_w_table.pack(fill='both', expand=True, pady=5)

        def redraw_wow():
            # 기존 위젯 삭제
            for w in f_w_graph.winfo_children(): w.destroy()
            for w in f_w_table.winfo_children(): w.destroy()
            
            target_year = int(year_var_w.get()) if year_var_w.get() else None
            if not target_year: return

            # [1] 데이터 필터링 (선택된 연도)
            df_year = raw_data[raw_data['Year'] == target_year].copy()
            
            if df_year.empty:
                tk.Label(f_w_graph, text=f"No Data for {target_year}", bg="#F1F5F9").pack(pady=20)
                return

            # [2] 피벗 생성 (1~52주 고정)
            # 행: Week_Num, 열: Bill To
            pivot = df_year.pivot_table(index='Week_Num', columns='Bill To', values='PL No', aggfunc='count', fill_value=0)
            
            # 1부터 52까지 강제 Reindex (데이터 없어도 X축 유지)
            full_weeks = range(1, 53)
            pivot = pivot.reindex(full_weeks, fill_value=0)

            # [3] Top 10 선별 (해당 연도 총합 기준)
            if not pivot.empty:
                top_10_cols = pivot.sum().sort_values(ascending=False).head(10).index
                pivot_top10 = pivot[top_10_cols]
            else:
                pivot_top10 = pd.DataFrame()
            
            # [4] 그래프 그리기
            fig_w, ax_w = plt.subplots(figsize=(10, 3.5)) 
            
            if not pivot_top10.empty:
                # Line Chart
                pivot_top10.plot(kind='line', ax=ax_w, marker='o', markersize=4, linewidth=1.5)
                
                # X축 설정 (1~52 고정, 분기별 눈금: 1, 13, 26, 39, 52)
                ax_w.set_xlim(0, 53) # 양옆 여백
                ticks = [1, 13, 26, 39, 52]
                ax_w.set_xticks(ticks)
                ax_w.set_xticklabels([str(t) for t in ticks])
                
                # 분기선 표시
                for t in ticks:
                    ax_w.axvline(x=t, color='#e0e0e0', linestyle='--', linewidth=1)

                ax_w.set_title(f"Weekly Order Trends (Top 10) - {target_year}", fontsize=12, fontweight='bold', pad=10)
                ax_w.set_xlabel("Week")
                ax_w.set_ylabel("Volume")
                ax_w.legend(title="Bill To", bbox_to_anchor=(1.04, 1), loc='upper left', fontsize='small', frameon=True)
                ax_w.grid(True, axis='y', linestyle=':', alpha=0.4)
            
            plt.tight_layout()
            
            # Export 저장
            self.current_export_figures[f'Weekly_Trend_{title}'] = fig_w
            self.current_export_tables[f'Weekly_Data_{title}'] = pivot_top10
            
            canvas_w = FigureCanvasTkAgg(fig_w, master=f_w_graph); canvas_w.draw(); canvas_w.get_tk_widget().pack(fill='both', expand=True)
            
            # [5] 하단 테이블 (Top 10 대상, Transpose하여 보기 좋게)
            if not pivot_top10.empty:
                pivot_display = pivot_top10.T 
                
                # 테이블 컬럼: Bill To + 1~52
                cols = ['Bill To'] + [str(i) for i in full_weeks]
                
                tree = ttk.Treeview(f_w_table, columns=cols, show='headings', height=8)
                
                # 스크롤바
                vsb = ttk.Scrollbar(f_w_table, orient="vertical", command=tree.yview)
                hsb = ttk.Scrollbar(f_w_table, orient="horizontal", command=tree.xview)
                tree.configure(yscroll=vsb.set, xscroll=hsb.set)
                
                vsb.pack(side='right', fill='y')
                hsb.pack(side='bottom', fill='x')
                tree.pack(fill='both', expand=True)
                
                # 헤더 설정 (Bill To)
                tree.heading('Bill To', text='Bill To')
                tree.column('Bill To', width=120, anchor='w') 
                
                # [핵심 수정] 나머지 주차 컬럼 (데이터 길이에 따른 자동 너비 조절)
                for c in cols[1:]:
                    # 해당 주차(Week)의 데이터들을 가져옴
                    col_data = pivot_display[int(c)]
                    
                    # 가장 긴 글자수 찾기 (헤더 vs 데이터 중 긴 것)
                    max_len = len(str(c)) 
                    for val in col_data:
                        v_len = len(f"{val:,}") # 콤마 포함한 길이 계산
                        if v_len > max_len: max_len = v_len
                    
                    # 너비 계산: (글자수 * 8픽셀) + 15픽셀 여유분
                    # 최소 너비는 40으로 설정하여 너무 좁아지는 것 방지
                    calc_width = (max_len * 9) + 15
                    final_width = max(40, calc_width)

                    tree.heading(c, text=c)
                    tree.column(c, width=final_width, anchor='center')
                
                # 데이터 삽입
                for bill_to, row in pivot_display.iterrows():
                    # 값에 천단위 콤마 추가해서 보기 좋게 변경 (옵션)
                    # 원치 않으시면 f"{x:,}" 대신 그냥 x 사용
                    vals = [bill_to] + [f"{x:,}" for x in row.tolist()]
                    tree.insert("", "end", values=vals)
        if avail_years_w:
            redraw_wow()   

        # 3. Station (Weekly/Monthly + Specific Period Filter)
        s_tab = ttk.Frame(nb); nb.add(s_tab, text=" 🏢 Station Analysis ")
        
        # [컨트롤 패널]
        ctrl_frame = tk.Frame(s_tab, bg="#F1F5F9", pady=5)
        ctrl_frame.pack(fill='x', padx=10)
        
        # 1. View 모드 (Weekly/Monthly)
        view_var = tk.StringVar(value="Weekly")
        tk.Label(ctrl_frame, text="View:", font=("Arial", 10, "bold"), bg="#F1F5F9", fg="#333").pack(side='left', padx=(0, 5))
        
        # 라디오 버튼 클릭 시 Period 목록 갱신 함수
        def update_period_options():
            mode = view_var.get()
            if mode == "Weekly":
                vals = ["All"] + [str(i) for i in range(1, 53)]
            else:
                vals = ["All"] + [str(i) for i in range(1, 13)]
            cb_period['values'] = vals
            cb_period.current(0) # All로 리셋
            
        rb_w = ttk.Radiobutton(ctrl_frame, text="Weekly", variable=view_var, value="Weekly", command=update_period_options)
        rb_w.pack(side='left', padx=5)
        rb_m = ttk.Radiobutton(ctrl_frame, text="Monthly", variable=view_var, value="Monthly", command=update_period_options)
        rb_m.pack(side='left', padx=5)
        
        # 2. Year 선택
        tk.Label(ctrl_frame, text="|  Year:", font=("Arial", 10, "bold"), bg="#F1F5F9", fg="#333").pack(side='left', padx=(15, 5))
        avail_years = sorted(list(raw_data['Year'].dropna().unique().astype(int)), reverse=True)
        year_var = tk.StringVar(value=str(avail_years[0]) if avail_years else "")
        cb_year = ttk.Combobox(ctrl_frame, textvariable=year_var, values=avail_years, width=6, state="readonly")
        cb_year.pack(side='left')
        
        # 3. [NEW] Period 선택 (특정 주차/월)
        tk.Label(ctrl_frame, text="|  Period:", font=("Arial", 10, "bold"), bg="#F1F5F9", fg="#333").pack(side='left', padx=(15, 5))
        period_var = tk.StringVar(value="All")
        cb_period = ttk.Combobox(ctrl_frame, textvariable=period_var, width=5, state="readonly")
        cb_period.pack(side='left')
        
        # 초기 Period 목록 설정
        update_period_options()

        # 4. 업데이트 버튼
        tk.Button(ctrl_frame, text="🔄 Update", font=("AppleGothic", 9, "bold"), bg="#3B82F6", fg="white", 
                  relief="flat", padx=10, command=lambda: redraw_station()).pack(side='left', padx=15)

        station_scroll_frame = ScrollableFrame(s_tab)
        station_scroll_frame.pack(fill='both', expand=True, pady=5)
        scroll_f = station_scroll_frame.scrollable_frame
        
        def redraw_station():
            for w in scroll_f.winfo_children(): w.destroy()
            
            # [1] 데이터 필터링 (Year)
            target_year = int(year_var.get()) if year_var.get() else None
            if not target_year: return
            
            station_df = raw_data[(raw_data['Status'] == 'Delivered') & (raw_data['Year'] == target_year)].copy()
            
            # [2] 데이터 필터링 (Period - 특정 주차 선택 시)
            sel_period = period_var.get()
            is_specific_period = (sel_period != "All")
            
            if is_specific_period:
                sel_val = int(sel_period)
                if view_var.get() == "Weekly":
                    station_df = station_df[station_df['Week_Num'] == sel_val]
                else:
                    station_df = station_df[station_df['Month_Num'] == sel_val]
            
            if station_df.empty:
                msg = f"No Data for {target_year}" + (f" - {view_var.get()} {sel_period}" if is_specific_period else "")
                tk.Label(scroll_f, text=msg, bg="#F1F5F9").pack(pady=20)
                return

            # [3] KPI 계산 (필터된 데이터 기준)
            if 'Origin_State' not in station_df.columns: station_df['Origin_State'] = 'Missing'
            if 'Station' not in station_df.columns: station_df['Station'] = 'Unknown'
            
            origin_counts = station_df['Origin_State'].value_counts()
            top_origin_state = origin_counts.idxmax() if not origin_counts.empty else "None"
            
            total_del = len(station_df)
            xyz_df = station_df[station_df['Station'].str.contains('XYZ|Unknown', case=False, na=False)]
            xyz_rate = (len(xyz_df) / total_del * 100) if total_del > 0 else 0
            top_st = station_df['Station'].value_counts().idxmax() if not station_df.empty else "-"
            
            # KPI 카드
            card_frame = ttk.Frame(scroll_f); card_frame.pack(fill='x', padx=10, pady=5)
            metrics = [("📦 Total Delivered", f"{total_del:,}", "black"), 
                       (f"🚚 Top Origin ({top_origin_state})", f"{origin_counts.max():,}" if not origin_counts.empty else "0", "#2a9d8f"),
                       ("🏆 Top Dest Station", top_st, "#0055aa"), 
                       ("⚠️ XYZ Rate", f"{xyz_rate:.1f}%", "#e63946")]
            for i, (t, v, c) in enumerate(metrics):
                f = tk.Frame(card_frame, bg="white", highlightbackground="#ccc", highlightthickness=1); f.grid(row=0, column=i, sticky="ew", padx=5)
                tk.Label(f, text=t, font=("Arial", 9, "bold"), fg="#555", bg="white").pack(anchor='w', padx=10, pady=(10,0))
                tk.Label(f, text=str(v), font=("Arial", 16, "bold"), fg=c, bg="white").pack(anchor='e', padx=10, pady=(0,10))
            card_frame.columnconfigure((0,1,2,3), weight=1)

            # [그래프 1] Top Origin: Dest Station Volume
            target_df = station_df[station_df['Origin_State'] == top_origin_state].copy()
            fig1, ax1 = plt.subplots(figsize=(10, 4.0))
            
            if not target_df.empty:
                mode = view_var.get()
                idx_col = 'Week_Num' if mode == "Weekly" else 'Month_Num'
                
                pivot = target_df.pivot_table(index=idx_col, columns='Station', values='PL No', aggfunc='count', fill_value=0)
                
                # [중요] All 선택 시에만 1-52주 고정축 사용 / 특정 주차는 해당 데이터만 표시
                if not is_specific_period:
                    if mode == "Weekly":
                        pivot = pivot.reindex(range(1, 53), fill_value=0)
                    else:
                        pivot = pivot.reindex(range(1, 13), fill_value=0)

                special_cols = [c for c in pivot.columns if 'XYZ' in str(c).upper() or 'UNKNOWN' in str(c).upper()]
                normal_cols = [c for c in pivot.columns if c not in special_cols]
                pivot = pivot[list(pivot[normal_cols].sum().sort_values(ascending=False).index) + special_cols]
                
                import itertools
                base_colors = ['#4E79A7', '#F28E2B', '#E15759', '#76B7B2', '#59A14F', '#EDC948', '#B07AA1', '#FF9DA7', '#9C755F', '#BAB0AC']
                color_cycle = itertools.cycle(base_colors)
                
                pivot.plot(kind='bar', stacked=True, ax=ax1, 
                           color=[('#D3D3D3' if c in special_cols else next(color_cycle)) for c in pivot.columns], 
                           width=0.8 if not is_specific_period else 0.4, # 하나일 땐 막대 얇게
                           edgecolor='#333', linewidth=0.5)
                
                ax1.legend(title="Dest Station", bbox_to_anchor=(1.05, 1), loc='upper left')
                
                # [요청사항 반영] 특정 주차 선택 시에만 레이블 표시
                if is_specific_period:
                    for c in ax1.containers:
                        labels = [f"{int(v.get_height())}" if v.get_height() > 0 else '' for v in c]
                        ax1.bar_label(c, labels=labels, label_type='center', color='black', fontsize=10, fontweight='bold')
                
                ax1.set_ylim(top=pivot.sum(axis=1).max() * 1.25)

                # X축 설정
                if not is_specific_period:
                    if mode == "Weekly":
                        ax1.set_xlim(-1, 52)
                        ticks = [1, 13, 26, 39, 52]
                        ax1.set_xticks([t-1 for t in ticks])
                        ax1.set_xticklabels([str(t) for t in ticks], rotation=0)
                        ax1.set_xlabel("Week (Quarterly)")
                    else:
                        ax1.set_xticks(range(12))
                        ax1.set_xticklabels([str(i) for i in range(1, 13)], rotation=0)
                        ax1.set_xlabel("Month (1-12)")
                else:
                    # 특정 주차 하나만 보일 때
                    ax1.set_xlabel(f"Selected {mode}: {sel_period}")
                    ax1.set_xticklabels([str(sel_period)], rotation=0)

                title_suffix = f"- {mode} {sel_period}" if is_specific_period else f"({target_year})"
                ax1.set_title(f"1. {top_origin_state} Origin: Dest Station Volume {title_suffix}", fontsize=12, fontweight='bold', pad=10)
            
            self.current_export_figures[f'Station_TopOrigin_{title}'] = fig1
            plt.tight_layout()
            FigureCanvasTkAgg(fig1, master=scroll_f).draw(); FigureCanvasTkAgg(fig1, master=scroll_f).get_tk_widget().pack(fill='both', expand=True, pady=5)

            ttk.Separator(scroll_f, orient='horizontal').pack(fill='x', pady=10)
            non_top = origin_counts[origin_counts.index != top_origin_state]
            
            # [그래프 2 & 3]
            fig2, (ax2a, ax2b) = plt.subplots(1, 2, figsize=(10, 3.5))
            if not non_top.empty:
                # [그래프 2]
                top_others = non_top.head(7).sort_values()
                top_others.plot(kind='barh', ax=ax2a, color=[('#E15759' if l == 'Missing' else '#59A14F') for l in top_others.index], alpha=0.8)
                ax2a.set_xlim(right=top_others.max() * 1.3)
                ax2a.set_title(f"Other Origins (Excl. {top_origin_state})"); ax2a.set_xlabel("Volume")
                for c in ax2a.containers: ax2a.bar_label(c, padding=3, color='black', fontsize=10)

                # [그래프 3]
                other_df = station_df[station_df['Origin_State'].isin(non_top.head(5).index)].copy()
                other_df['Dest_Type'] = other_df['Station'].apply(lambda x: 'XYZ' if 'XYZ' in str(x).upper() else 'Normal')
                p_org = other_df.pivot_table(index='Origin_State', columns='Dest_Type', values='PL No', aggfunc='count', fill_value=0)
                color_map = {'Normal': '#59A14F', 'XYZ': '#D3D3D3'}
                p_org.plot(kind='bar', stacked=True, ax=ax2b, rot=0, color=[color_map.get(c, '#333') for c in p_org.columns])
                ax2b.set_ylim(top=p_org.sum(axis=1).max() * 1.25)
                ax2b.set_title("Dest Type by Origin")
                
                # 특정 주차일 경우 여기도 레이블 표시 (옵션)
                if is_specific_period:
                    for c in ax2b.containers:
                        labels = [f"{int(v.get_height())}" if v.get_height() > 0 else '' for v in c]
                        ax2b.bar_label(c, labels=labels, label_type='center', color='black', fontsize=9)
            
            self.current_export_figures[f'Station_Others_{title}'] = fig2
            plt.tight_layout()
            FigureCanvasTkAgg(fig2, master=scroll_f).draw(); FigureCanvasTkAgg(fig2, master=scroll_f).get_tk_widget().pack(fill='both', expand=True, pady=5)

        # 4. Geography
        g_tab = ttk.Frame(nb); nb.add(g_tab, text=" 🗺️ Geography ")
        if 'Dest State' in viz_data.columns:
            fig_geo, (ax_g1, ax_g2) = plt.subplots(1, 2, figsize=(10, 3.5))
            top_states = viz_data['Dest State'].value_counts().head(10)
            if not top_states.empty:
                top_states.sort_values().plot.barh(ax=ax_g1, color='#58D68D'); ax_g1.set_title("Top 10 States by Volume")
                
                # [수정] X축 범위 확장
                ax_g1.set_xlim(right=top_states.max() * 1.3)
                
                for c in ax_g1.containers: ax_g1.bar_label(c, padding=3, color='black', fontsize=10)
                
                state_lt = viz_data[viz_data['Dest State'].isin(top_states.index)].groupby('Dest State')['Lead_Time_Days'].mean().reindex(top_states.index).sort_values()
                state_lt.plot.barh(ax=ax_g2, color='#F1948A'); ax_g2.set_title("Avg LT by Top States")
                
                # [수정] X축 범위 확장
                ax_g2.set_xlim(right=state_lt.max() * 1.3)
                
                for c in ax_g2.containers: ax_g2.bar_label(c, fmt='%.1f', padding=3, color='black', fontsize=10)
                
                self.current_export_figures[f'Geo_TopStates_{title}'] = fig_geo
                FigureCanvasTkAgg(fig_geo, master=g_tab).draw(); FigureCanvasTkAgg(fig_geo, master=g_tab).get_tk_widget().pack(fill='both', expand=True)

        # 5. Service
        v_tab = ttk.Frame(nb); nb.add(v_tab, text=" 📦 Service ")
        if 'Sub Type' in viz_data.columns:
            fig_svc, (ax_s1, ax_s2) = plt.subplots(1, 2, figsize=(10, 3.5))
            viz_data['Sub Type'].value_counts().plot.pie(ax=ax_s1, autopct='%1.1f%%', cmap='Pastel1'); ax_s1.set_title("Service Type Share")
            
            grp_s = viz_data.groupby('Sub Type')['Lead_Time_Days'].mean().sort_values()
            grp_s.plot.bar(ax=ax_s2, color='skyblue', rot=0); ax_s2.set_title("Avg Lead Time by Service Type")
            
            # [수정] Y축 범위 확장
            ax_s2.set_ylim(top=grp_s.max() * 1.25)
            
            for c in ax_s2.containers: ax_s2.bar_label(c, fmt='%.1f', padding=2, color='black', fontsize=10)
            
            self.current_export_figures[f'Service_{title}'] = fig_svc
            FigureCanvasTkAgg(fig_svc, master=v_tab).draw(); FigureCanvasTkAgg(fig_svc, master=v_tab).get_tk_widget().pack(fill='both', expand=True)

        # 6. Timing
        tm_tab = ttk.Frame(nb); nb.add(tm_tab, text=" 📅 Timing ")
        if 'Day_of_Week' in viz_data.columns:
            days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
            fig_time, ax_t = plt.subplots(figsize=(8, 3.5))
            
            day_counts = viz_data['Day_of_Week'].value_counts().reindex(days).fillna(0)
            day_counts.plot.bar(ax=ax_t, color='purple', alpha=0.6, rot=0)
            ax_t.set_title("Weekly Order Pattern")
            ax_t.set_ylim(top=day_counts.max() * 1.25)
            for c in ax_t.containers: ax_t.bar_label(c, fmt='%d', padding=2, color='black', fontsize=10)
            
            # [추가됨] Export용 저장
            self.current_export_figures[f'Timing_{title}'] = fig_time
            
            FigureCanvasTkAgg(fig_time, master=tm_tab).draw(); FigureCanvasTkAgg(fig_time, master=tm_tab).get_tk_widget().pack(fill='both', expand=True)

        # 7. In-Transit
        tr_tab = ttk.Frame(nb); nb.add(tr_tab, text=" 🚚 In-Transit ")
        it_df = raw_data[raw_data['Status_Group'] == 'In-Transit'].copy()
        if not it_df.empty:
            it_df['Days'] = (datetime.now() - it_df['Calc_Pickup_Date']).dt.days.fillna(0).astype(int)
            fig_it, ax_it = plt.subplots(figsize=(8, 3.0))
            
            transit_counts = pd.cut(it_df['Days'], bins=[0, 3, 7, 14, 30, 100], labels=['0-3d', '4-7d', '8-14d', '15-30d', '30d+']).value_counts().sort_index()
            transit_counts.plot.bar(ax=ax_it, color='#F39C12', rot=0)
            ax_it.set_title("In-Transit Status")
            if not transit_counts.empty: ax_it.set_ylim(top=transit_counts.max() * 1.25)
            for c in ax_it.containers: ax_it.bar_label(c, fmt='%d', padding=2, color='black', fontsize=10)
            
            # [추가됨] Export용 저장 (그래프 + 엑셀 시트)
            self.current_export_figures[f'InTransit_Graph_{title}'] = fig_it
            self.current_export_tables[f'InTransit_List_{title}'] = it_df[['PL No', 'Bill To', 'Calc_Pickup_Date', 'Days', 'Dest State', 'Station']].sort_values('Days', ascending=False)

            FigureCanvasTkAgg(fig_it, master=tr_tab).draw(); FigureCanvasTkAgg(fig_it, master=tr_tab).get_tk_widget().pack(fill="x")
            
            tree = ttk.Treeview(tr_tab, columns=("PL", "BillTo", "Date", "Days"), show='headings', height=8)
            for c in ["PL", "BillTo", "Date", "Days"]: tree.heading(c, text=c)
            tree.pack(fill='both', expand=True)
            for _, r in it_df.sort_values('Days', ascending=False).iterrows(): tree.insert("", "end", values=(r['PL No'], r['Bill To'], str(r['Calc_Pickup_Date'])[:10], int(r['Days'])))
            
    def run_single_import(self):
        f = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if f: self.process_import_thread([f], "Single Import")

    def run_bulk_import(self):
        d = filedialog.askdirectory()
        if d:
            files = glob.glob(os.path.join(d, "*.xls*"))
            if files: self.process_import_thread(files, "Bulk Import")

    def process_import_thread(self, file_list, mode):
        pw = ProgressWindow(self, title=mode)
        def task():
            try:
                self.data_manager.backup_data()
                batch = pd.DataFrame()
                for i, path in enumerate(file_list):
                    pw.update_progress_safe(i, len(file_list), f"Reading {os.path.basename(path)}...")
                    batch = pd.concat([batch, self.data_manager.process_file(path)])
                    time.sleep(0.01) # Keep UI responsive
                if not batch.empty:
                    self.data_manager.merge_dataframe(batch); self.data_manager.save_data()
                    self.last_update_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    self.after(0, lambda: self.lbl_update.config(text=f"Last updated: {self.last_update_date}"))
                    self.after(0, lambda: messagebox.showinfo("Success", f"Updated {len(batch)} records."))
            except Exception as e: self.after(0, lambda: messagebox.showerror("Error", str(e)))
            finally: self.after(0, pw.close)
        threading.Thread(target=task, daemon=True).start()

    def run_export_thread(self):
        # 1. 예외 처리
        if not EXCEL_AVAILABLE: return messagebox.showerror("Error", "openpyxl required.")
        if not self.current_export_tables: return messagebox.showwarning("Warning", "Run analysis first.")
        
        # 2. 경로 선택
        d = filedialog.askdirectory()
        if not d: return
        
        # 3. 로딩창 생성
        pw = ProgressWindow(self, title="Exporting Data")
        
        def task():
            try:
                # 총 작업량 계산 (표 개수 + 그래프 개수)
                tables = self.current_export_tables
                figures = self.current_export_figures
                total_steps = len(tables) + len(figures)
                current_step = 0
                
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                excel_path = os.path.join(d, f"Report_{ts}.xlsx")
                
                # [단계 1] 엑셀 파일 생성 및 시트 저장
                pw.update_progress_safe(0, total_steps, "Initializing Excel...")
                with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                    for name, df in tables.items():
                        current_step += 1
                        # 로그 업데이트 (너무 긴 이름은 잘라서 표시)
                        safe_name = name[:20] + "..." if len(name) > 20 else name
                        pw.update_progress_safe(current_step, total_steps, f"Saving Sheet: {safe_name}")
                        
                        df.to_excel(writer, sheet_name=name[:30]) # 엑셀 시트 이름 제한(30자)
                        time.sleep(0.05) # UI가 갱신될 틈을 줌
                
                # [단계 2] 그래프 이미지 저장
                for name, fig in figures.items():
                    current_step += 1
                    safe_name = name[:20] + "..." if len(name) > 20 else name
                    pw.update_progress_safe(current_step, total_steps, f"Saving Image: {safe_name}")
                    
                    fig.savefig(os.path.join(d, f"{name}_{ts}.png"), dpi=100)
                    time.sleep(0.05)
                
                # [단계 3] 완료
                pw.update_progress_safe(total_steps, total_steps, "Export Complete!")
                time.sleep(0.5) # 완료 메시지 잠시 보여줌
                
                self.after(0, lambda: messagebox.showinfo("Success", "Export Done."))
                
            except Exception as e: 
                self.after(0, lambda: messagebox.showerror("Error", str(e)))
            finally: 
                self.after(0, pw.close)
        
        # 스레드 시작
        threading.Thread(target=task, daemon=True).start()

    def on_closing(self):
        self.quit(); self.destroy(); os._exit(0)

if __name__ == "__main__":
    app = LogisticsApp()
    app.mainloop()