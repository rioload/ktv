import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import ttkbootstrap as ttk  
from ttkbootstrap.constants import *
import subprocess
import threading
from pathlib import Path
import urllib.request
import zipfile
import shutil
import multiprocessing
import re
import webbrowser
import ssl

# 防止 EXE 遞迴啟動
if __name__ == "__main__":
    multiprocessing.freeze_support()
    if len(sys.argv) > 1 and not any(arg.startswith('--multiprocessing') for arg in sys.argv):
        sys.exit(0)

class AudioSeparatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("KTV AI Studio - 人工智慧伴唱帶生成器")
        
        # 🚀 呼叫置中邏輯 (設定寬高 1050x750)
        self.center_window(1050, 750)
        
        # 取得執行路徑 (EXE 所在目錄)
        if getattr(sys, 'frozen', False):
            self.app_dir = Path(sys.executable).parent
        else:
            self.app_dir = Path(__file__).parent
            
        migrations = {
            "bin": "engine_ffmpeg",
            "python_env": "runtime_python",
            "packages": "ai_libraries",
            "models": "ai_models"
        }
        for old_name, new_name in migrations.items():
            old_p = self.app_dir / old_name
            new_p = self.app_dir / new_name
            if old_p.exists() and not new_p.exists():
                try: old_p.rename(new_p)
                except: pass

        self.bin_dir = self.app_dir / "engine_ffmpeg"      
        self.py_dir = self.app_dir / "runtime_python"      
        self.lib_dir = self.app_dir / "ai_libraries"       
        self.models_dir = self.app_dir / "ai_models"       
        
        for d in [self.bin_dir, self.py_dir, self.lib_dir, self.models_dir]:
            d.mkdir(exist_ok=True)
        
        self.local_python = self.py_dir / "python.exe"
        
        os.environ["PATH"] = f"{self.bin_dir}{os.pathsep}{self.py_dir}{os.pathsep}{os.environ['PATH']}"
        if hasattr(os, 'add_dll_directory'):
            try: os.add_dll_directory(str(self.bin_dir))
            except: pass
        
        os.environ["PYTHONPATH"] = str(self.lib_dir)
        
        self.subp_flags = subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
        self.file_list = []
        self.is_processing = False
        
        self.setup_ui()
        self.root.after(500, lambda: self.check_components(prompt=False))

    def center_window(self, width, height):
        """將視窗置中於螢幕"""
        # 取得螢幕的寬度與高度
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        # 計算置中的 X 與 Y 座標
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        
        # 套用尺寸與座標
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def setup_ui(self):
        # ==========================================
        # 🚀 全新佈局：側邊導覽列 (Sidebar)
        # ==========================================
        self.sidebar = ttk.Frame(self.root, bootstyle="dark", padding=15)
        self.sidebar.pack(side=tk.LEFT, fill=tk.Y)

        # 側邊欄 Logo 區 (加上空格解決斜體被切邊問題)
        ttk.Label(self.sidebar, text="KTV \nAI Studio ", font=("Impact", 28, "italic"), bootstyle="inverse-dark", justify="center").pack(pady=(20, 40))

        # 側邊欄導覽按鈕 
        self.nav_var = tk.StringVar(value="page1")
        nav_btns = [
            ("📺 製作 YT 伴唱", "page1", "danger"),
            ("📥 批次下載 YT", "page2", "info"),
            ("🎬 本地影片轉檔", "page3", "primary"),
            ("🎵 提取本地音檔", "page4", "success")
        ]
        
        for text, val, style in nav_btns:
            ttk.Radiobutton(
                self.sidebar, text=text, variable=self.nav_var, value=val, 
                command=self.switch_page, bootstyle=f"{style}-toolbutton", width=18
            ).pack(pady=8, ipady=8)

        # 側邊欄底部工具
        ttk.Frame(self.sidebar).pack(fill=tk.Y, expand=True) # 彈性空間推到底部
        self.update_btn = ttk.Button(self.sidebar, text="⚙️ 系統初始化/修復", command=self.check_components, bootstyle="outline-warning")
        self.update_btn.pack(fill=tk.X, pady=10, ipady=5)

        # ==========================================
        # 🚀 全新佈局：主內容區 (Main Content)
        # ==========================================
        self.main_area = ttk.Frame(self.root, padding=25)
        self.main_area.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # 頂部動態標題
        self.header_label = ttk.Label(self.main_area, text="📺 YouTube 批次轉 MKV 伴唱帶", font=("Microsoft JhengHei", 20, "bold"), bootstyle="danger")
        self.header_label.pack(anchor=tk.W, pady=(0, 20))

        # --- 卡片疊加區 (容納四個頁面) ---
        self.content_container = ttk.Frame(self.main_area)
        self.content_container.pack(fill=tk.BOTH, expand=True)
        self.content_container.grid_rowconfigure(0, weight=1)
        self.content_container.grid_columnconfigure(0, weight=1)

        self.pages = {}

        # 📝 Page 1: YouTube 轉 MKV
        p1 = ttk.Frame(self.content_container)
        p1.grid(row=0, column=0, sticky="nsew")
        ttk.Label(p1, text="請貼上 YouTube 網址 (可貼多筆，一行一網址):", font=("Arial", 11)).pack(anchor=tk.W, pady=(0, 5))
        self.yt_text = scrolledtext.ScrolledText(p1, height=5, font=("Consolas", 11))
        self.yt_text.pack(fill=tk.BOTH, expand=True, pady=5)
        self.yt_text.bind("<Button-1>", self.quick_paste_url)
        ttk.Label(p1, text="💡 提示: 點擊文字框內可自動貼上剪貼簿內容", bootstyle="secondary").pack(anchor=tk.W)
        self.pages["page1"] = p1

        # 📝 Page 2: YouTube 純下載
        p2 = ttk.Frame(self.content_container)
        p2.grid(row=0, column=0, sticky="nsew")
        ttk.Label(p2, text="請貼上 YouTube 網址 (可貼多筆，一行一網址):", font=("Arial", 11)).pack(anchor=tk.W, pady=(0, 5))
        self.yt_dl_text = scrolledtext.ScrolledText(p2, height=5, font=("Consolas", 11))
        self.yt_dl_text.pack(fill=tk.BOTH, expand=True, pady=5)
        self.yt_dl_text.bind("<Button-1>", self.quick_paste_url)
        
        dl_opt_frame = ttk.Frame(p2)
        dl_opt_frame.pack(fill=tk.X, pady=10)
        self.dl_type_var = tk.StringVar(value="both")
        ttk.Label(dl_opt_frame, text="下載模式:").pack(side=tk.LEFT, padx=(0, 10))
        ttk.Radiobutton(dl_opt_frame, text="無損 MP3 + MP4 影像", variable=self.dl_type_var, value="both", bootstyle="info").pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(dl_opt_frame, text="僅提取無損 MP3", variable=self.dl_type_var, value="mp3", bootstyle="info").pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(dl_opt_frame, text="僅下載 MP4 影像", variable=self.dl_type_var, value="mp4", bootstyle="info").pack(side=tk.LEFT, padx=10)
        self.pages["page2"] = p2

        # 📝 Page 3: 本地影片
        p3 = ttk.Frame(self.content_container)
        p3.grid(row=0, column=0, sticky="nsew")
        
        v_btn_frame = ttk.Frame(p3)
        v_btn_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, 10))
        self.v_list = [] 
        ttk.Button(v_btn_frame, text="➕ 加入單一影片", command=self.browse_local_video, bootstyle="primary").pack(side=tk.LEFT, padx=5)
        ttk.Button(v_btn_frame, text="📁 加入整個資料夾", command=self.browse_local_v_folder, bootstyle="info").pack(side=tk.LEFT, padx=5)
        ttk.Button(v_btn_frame, text="🗑️ 移除選取", command=self.remove_selected_v, bootstyle="outline-danger").pack(side=tk.RIGHT, padx=5)
        ttk.Button(v_btn_frame, text="🧹 清除全部", command=self.clear_v_list, bootstyle="outline-secondary").pack(side=tk.RIGHT, padx=5)
        
        self.v_listbox = tk.Listbox(p3, selectmode=tk.EXTENDED, bg="#1a1a1a", fg="#00ffcc", font=("Consolas", 10), relief="flat", highlightthickness=1, highlightcolor="#00ffcc")
        self.v_listbox.pack(fill=tk.BOTH, expand=True)
        self.pages["page3"] = p3

        # 📝 Page 4: 本地音檔
        p4 = ttk.Frame(self.content_container)
        p4.grid(row=0, column=0, sticky="nsew")
        
        f_btn_frame = ttk.Frame(p4)
        f_btn_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, 10))
        self.file_list = []
        ttk.Button(f_btn_frame, text="➕ 加入音檔", command=self.browse_file, bootstyle="success").pack(side=tk.LEFT, padx=5)
        ttk.Button(f_btn_frame, text="🗑️ 移除選取", command=self.remove_selected_file, bootstyle="outline-danger").pack(side=tk.RIGHT, padx=5)
        ttk.Button(f_btn_frame, text="🧹 清除全部", command=self.clear_files, bootstyle="outline-secondary").pack(side=tk.RIGHT, padx=5)

        self.file_listbox = tk.Listbox(p4, selectmode=tk.EXTENDED, bg="#1a1a1a", fg="#00ffcc", font=("Consolas", 10), relief="flat", highlightthickness=1, highlightcolor="#00ffcc")
        self.file_listbox.pack(fill=tk.BOTH, expand=True)
        self.pages["page4"] = p4

        # ==========================================
        # 🚀 全新佈局：底部通用控制中心
        # ==========================================
        control_panel = ttk.Frame(self.main_area)
        control_panel.pack(fill=tk.X, pady=(20, 0))

        ttk.Separator(control_panel, bootstyle="secondary").pack(fill=tk.X, pady=(0, 15))

        # 控制列 1: 目錄與硬體
        row1 = ttk.Frame(control_panel)
        row1.pack(fill=tk.X, pady=5)
        
        self.output_dir_var = tk.StringVar(value=str(self.app_dir / "output"))
        ttk.Label(row1, text="📂 輸出:").pack(side=tk.LEFT)
        ttk.Entry(row1, textvariable=self.output_dir_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10)
        ttk.Button(row1, text="瀏覽", command=self.browse_output_dir, bootstyle="secondary").pack(side=tk.LEFT, padx=(0, 20))
        
        self.device_var = tk.StringVar(value="cpu")
        ttk.Radiobutton(row1, text="CPU", variable=self.device_var, value="cpu", bootstyle="primary").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(row1, text="GPU", variable=self.device_var, value="gpu", bootstyle="primary").pack(side=tk.LEFT, padx=5)
        ttk.Button(row1, text="🔍 檢測 GPU", command=self.check_gpu_env, bootstyle="outline-warning").pack(side=tk.LEFT, padx=10)

        # 控制列 2: AI 模型與開關
        row2 = ttk.Frame(control_panel)
        row2.pack(fill=tk.X, pady=10)
        
        ttk.Label(row2, text="🧠 模型:").pack(side=tk.LEFT)
        self.model_var = tk.StringVar(value="UVR-MDX-NET-Inst_HQ_3.onnx")
        model_opts = ["UVR-MDX-NET-Inst_HQ_3.onnx (MDX - 伴奏優化)", "UVR-MDX-NET-Inst_HQ_4.onnx (MDX - 高品質綜合)", "Kim_Vocal_2.onnx (MDX - 極致人聲)", "htdemucs.yaml (Demucs - 4音軌)", "htdemucs_ft.yaml (Demucs - 流行樂)", "htdemucs_6s.yaml (Demucs - 6音軌)"]
        ttk.Combobox(row2, textvariable=self.model_var, values=model_opts, state="readonly", width=40).pack(side=tk.LEFT, padx=10)
        
        ttk.Label(row2, text="格式:").pack(side=tk.LEFT, padx=(10, 5))
        self.output_format_var = tk.StringVar(value="mp3")
        for fmt in ["mp3", "wav", "flac"]:
            ttk.Radiobutton(row2, text=fmt.upper(), variable=self.output_format_var, value=fmt, bootstyle="info").pack(side=tk.LEFT, padx=3)
        
        self.denoise_var = tk.BooleanVar(value=True)
        self.overlap_var = tk.DoubleVar(value=0.5)
        ttk.Checkbutton(row2, text="AI 深度去噪", variable=self.denoise_var, bootstyle="success-round-toggle").pack(side=tk.RIGHT)

        # 執行大按鈕與進度區
        exec_frame = ttk.Frame(control_panel)
        exec_frame.pack(fill=tk.X, pady=(15, 5))
        
        self.start_btn = ttk.Button(exec_frame, text="啟動核心任務", command=self.on_start_click, bootstyle="danger", width=20)
        self.start_btn.pack(side=tk.LEFT, ipady=8)
        
        prog_frame = ttk.Frame(exec_frame)
        prog_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=20)
        
        status_lbl_frame = ttk.Frame(prog_frame)
        status_lbl_frame.pack(fill=tk.X)
        self.status_var = tk.StringVar(value="系統就緒")
        self.status_label = ttk.Label(status_lbl_frame, textvariable=self.status_var, bootstyle="success")
        self.status_label.pack(side=tk.LEFT)
        self.progress_text = ttk.Label(status_lbl_frame, text="0%")
        self.progress_text.pack(side=tk.RIGHT)
        
        self.progress_bar = ttk.Progressbar(prog_frame, orient=tk.HORIZONTAL, mode='determinate', bootstyle="danger-striped")
        self.progress_bar.pack(fill=tk.X, pady=5)

        # 日誌區
        self.log_area = scrolledtext.ScrolledText(self.main_area, height=7, font=("Consolas", 9), bg="#0d0d0d", fg="#a6a6a6")
        self.log_area.pack(fill=tk.BOTH, expand=True)

        self.root.update_idletasks()
        self.switch_page() 
        self.show_welcome_message()

    def switch_page(self):
        page = self.nav_var.get()
        self.pages[page].tkraise()
        
        configs = {
            "page1": {"title": "📺 YouTube 批次轉 MKV 伴唱帶", "btn": "開始批次製作", "style": "danger"},
            "page2": {"title": "📥 YouTube 批次無損下載", "btn": "開始批次下載", "style": "info"},
            "page3": {"title": "🎬 本地影片批次轉 KTV", "btn": "開始影片處理", "style": "primary"},
            "page4": {"title": "🎵 本地音檔批量分離", "btn": "開始音訊分離", "style": "success"}
        }
        
        self.header_label.config(text=configs[page]["title"], bootstyle=configs[page]["style"])
        self.start_btn.config(text=configs[page]["btn"], bootstyle=configs[page]["style"])
        self.progress_bar.config(bootstyle=f"{configs[page]['style']}-striped")

    def on_start_click(self):
        page = self.nav_var.get()
        if page == "page1": self.start_yt_process()
        elif page == "page2": self.start_pure_download()
        elif page == "page3": self.start_local_v_process()
        elif page == "page4": self.start_separation()

    def get_current_yt_urls(self):
        page = self.nav_var.get()
        text = self.yt_text.get("1.0", tk.END) if page == "page1" else (self.yt_dl_text.get("1.0", tk.END) if page == "page2" else "")
        return [u.strip() for u in text.replace(',', '\n').split('\n') if "http" in u.strip()]

    def quick_paste_url(self, event):
        try:
            clipboard = self.root.clipboard_get().strip()
            if "http" in clipboard:
                widget = event.widget
                current_val = widget.get("1.0", tk.END).strip()
                clip_urls = [u.strip() for u in clipboard.replace(',', '\n').split('\n') if "http" in u.strip()]
                added_count = 0
                for u in clip_urls:
                    if u not in current_val:
                        widget.insert(tk.END, ("\n" if widget.get("1.0", tk.END).strip() else "") + u)
                        current_val = widget.get("1.0", tk.END).strip()
                        added_count += 1
                if added_count > 0: self.log(f"📋 已從剪貼簿自動貼上 {added_count} 個網址")
        except: pass

    # --- 本地影片與音檔功能 ---
    def browse_local_video(self):
        file_paths = filedialog.askopenfilenames(filetypes=[("影片檔案", "*.mp4 *.mkv *.avi *.mov *.wmv *.webm"), ("所有檔案", "*.*")])
        if file_paths:
            for fp in file_paths:
                fp_abs = str(Path(fp).absolute())
                if fp_abs not in self.v_list:
                    self.v_list.append(fp_abs)
                    self.v_listbox.insert(tk.END, os.path.basename(fp_abs))

    def browse_local_v_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            found = False
            for ext in [".mp4", ".mkv", ".avi", ".mov", ".wmv", ".webm"]:
                for fp in Path(folder_path).glob(f"*{ext}"):
                    fp_abs = str(fp.absolute())
                    if fp_abs not in self.v_list:
                        self.v_list.append(fp_abs)
                        self.v_listbox.insert(tk.END, fp.name)
                        found = True
            if not found: self.log(f"⚠️ 在資料夾中找不到常見影片檔案。")

    def remove_selected_v(self):
        for index in reversed(self.v_listbox.curselection()):
            self.v_list.pop(index)
            self.v_listbox.delete(index)

    def clear_v_list(self):
        self.v_list.clear(); self.v_listbox.delete(0, tk.END)

    def browse_file(self):
        filenames = filedialog.askopenfilenames(filetypes=[("Audio", "*.mp3 *.wav *.flac *.m4a"), ("All", "*.*")])
        if filenames:
            for f in filenames:
                f_abs = str(Path(f).absolute())
                if f_abs not in self.file_list:
                    self.file_list.append(f_abs)
                    self.file_listbox.insert(tk.END, os.path.basename(f_abs))

    def remove_selected_file(self):
        for index in reversed(self.file_listbox.curselection()):
            self.file_list.pop(index)
            self.file_listbox.delete(index)

    def clear_files(self):
        self.file_list.clear(); self.file_listbox.delete(0, tk.END)

    def browse_output_dir(self):
        directory = filedialog.askdirectory()
        if directory: self.output_dir_var.set(directory)

    # --- YouTube 批次製作 MKV ---
    def start_yt_process(self):
        urls = self.get_current_yt_urls()
        if not urls: return messagebox.showwarning("警告", "請輸入至少一個 YouTube 網址！")
        if self.is_processing: return

        if self.device_var.get() == "gpu" and not self._quick_check_gpu():
            if messagebox.askyesno("環境未就緒", "偵測到您的 GPU 環境尚未配置完成，是否現在進行一鍵修復？"):
                return self.check_gpu_env()
            else: self.device_var.set("cpu")

        self.is_processing = True
        self.start_btn.config(state=tk.DISABLED)
        self.log_area.delete(1.0, tk.END)
        self.update_status(f"準備批次處理 {len(urls)} 部影片...", "warning")
        threading.Thread(target=self.yt_batch_process, args=(urls,), daemon=True).start()

    def yt_batch_process(self, urls):
        total, success_count = len(urls), 0
        output_dir = self.output_dir_var.get()
        Path(output_dir).mkdir(parents=True, exist_ok=True)
        
        for i, url in enumerate(urls):
            if not self.is_processing: break
            self.log(f"\n========== 🎬 開始處理 ({i+1}/{total}): {url} ==========")
            self.update_progress(int(i / total * 100), f"正在處理第 {i+1}/{total} 個")
            
            video_file, audio_file = self.download_youtube(url, output_dir)
            if not video_file or not audio_file:
                self.log(f"❌ ({i+1}/{total}) 下載失敗跳過。"); continue

            self.log("  > 正在分離人聲與伴奏...")
            if self.run_audio_separator(audio_file, output_dir):
                self.log("  > 正在合成 MKV...")
                voc_file, inst_file = self.consolidate_stems(audio_file, video_file, output_dir)
                if voc_file and inst_file:
                    mkv_file = Path(output_dir) / f"{Path(video_file).stem}_KTV.mkv"
                    if self.synthesize_mkv(video_file, voc_file, inst_file, str(mkv_file)):
                        self.log(f"✅ ({i+1}/{total}) 成功製作: {mkv_file.name}")
                        success_count += 1
                    else: self.log(f"❌ ({i+1}/{total}) MKV 合成失敗。")
                else: self.log(f"❌ ({i+1}/{total}) 找不到分離後的檔案。")
            else: self.log(f"❌ ({i+1}/{total}) 音訊分離失敗。")
                
        self.update_progress(100, "處理完成")
        self.log(f"\n✨ YouTube 批次任務結束！成功: {success_count}/{total}")
        self.root.after(0, lambda: messagebox.showinfo("完成", f"批次處理完成！\n成功: {success_count}/{total} 個"))
        self.root.after(0, lambda: os.startfile(output_dir))
        self.finish_processing()

    # --- YouTube 批次純下載 ---
    def start_pure_download(self):
        urls = self.get_current_yt_urls()
        if not urls: return messagebox.showwarning("警告", "請輸入至少一個 YouTube 網址！")
        if self.is_processing: return

        self.is_processing = True
        self.start_btn.config(state=tk.DISABLED)
        self.log_area.delete(1.0, tk.END)
        self.update_status(f"準備下載 {len(urls)} 個檔案...", "warning")
        threading.Thread(target=self.pure_download_batch_process, args=(urls,), daemon=True).start()

    def pure_download_batch_process(self, urls):
        try:
            total, output_dir, dl_type = len(urls), self.output_dir_var.get(), self.dl_type_var.get()
            for i, url in enumerate(urls):
                if not self.is_processing: break
                self.log(f"\n========== 🚀 開始下載 ({i+1}/{total}): {url} ==========")
                self.update_progress(int(i / total * 100), f"正在下載第 {i+1}/{total} 個")
                
                if dl_type in ["both", "mp4"]: self.download_youtube(url, output_dir, mode="mp4")
                if dl_type in ["both", "mp3"]: self.download_youtube(url, output_dir, mode="mp3")
                    
            self.update_progress(100, "全部下載完成")
            self.log(f"\n✅ 批次下載任務結束！")
            self.root.after(0, lambda: messagebox.showinfo("完成", "YouTube 批次下載完成！"))
            self.root.after(0, lambda: os.startfile(output_dir))
        except Exception as e: self.log(f"❌ 批次下載出錯: {str(e)}")
        finally: self.finish_processing()

    # --- 本地影片批次處理 ---
    def start_local_v_process(self):
        if not self.v_list: return messagebox.showwarning("警告", "請先加入影片檔案！")
        if self.is_processing: return

        self.is_processing = True
        self.start_btn.config(state=tk.DISABLED)
        self.log_area.delete(1.0, tk.END)
        self.update_status("正在處理本地影片...", "warning")
        threading.Thread(target=self.local_v_batch_process, daemon=True).start()

    def local_v_batch_process(self):
        try:
            output_dir, total = self.output_dir_var.get(), len(self.v_list)
            if not os.path.exists(output_dir): os.makedirs(output_dir)
            
            for i, video_path in enumerate(self.v_list):
                if not self.is_processing: break
                video_stem = Path(video_path).stem
                self.log(f"\n--- 正在處理 ({i+1}/{total}): {os.path.basename(video_path)} ---")
                
                self.v_listbox.selection_clear(0, tk.END); self.v_listbox.selection_set(i); self.v_listbox.see(i)
                temp_audio = Path(output_dir) / f"{video_stem}_temp_audio.mp3"
                pb, ps = int((i / total) * 100), int(100 / total)
                
                self.update_progress(pb + int(ps * 0.1), f"擷取音訊 ({i+1}/{total})")
                subprocess.run([str(self.bin_dir / "ffmpeg.exe"), "-y", "-i", video_path, "-vn", "-acodec", "libmp3lame", "-ab", "320k", str(temp_audio)], check=True, creationflags=self.subp_flags)
                
                self.update_progress(pb + int(ps * 0.3), f"AI 分離 ({i+1}/{total})")
                if self.run_audio_separator(str(temp_audio), output_dir):
                    voc_file, inst_file = self.consolidate_stems(str(temp_audio), video_path, output_dir)
                    if voc_file and inst_file:
                        self.update_progress(pb + int(ps * 0.8), f"合成 MKV ({i+1}/{total})")
                        mkv_file = Path(output_dir) / f"{video_stem}_KTV.mkv"
                        if self.synthesize_mkv(video_path, voc_file, inst_file, str(mkv_file)): self.log(f"✅ 生成 MKV: {mkv_file.name}")
                        else: self.log(f"❌ {video_stem} MKV 合成失敗。")
                    else: self.log(f"❌ {video_stem} 找不到分離檔案。")
                else: self.log(f"❌ {video_stem} 音訊分離失敗。")

            self.update_progress(100, "批次處理完成")
            self.log("\n✨ 影片批次處理結束！")
            self.root.after(0, lambda: messagebox.showinfo("完成", f"已完成 {total} 個影片！"))
            self.root.after(0, lambda: os.startfile(output_dir))
        except Exception as e: self.log(f"❌ 處理出錯: {str(e)}")
        finally: self.finish_processing()

    # --- 本地音檔批次處理 ---
    def start_separation(self):
        if not self.file_list: return messagebox.showwarning("警告", "請先加入音檔！")
        if self.is_processing: return

        if self.device_var.get() == "gpu" and not self._quick_check_gpu():
            if messagebox.askyesno("環境未就緒", "偵測到 GPU 環境未完成配置，是否進行修復？"): return self.check_gpu_env()
            else: self.device_var.set("cpu")

        self.is_processing = True
        self.start_btn.config(state=tk.DISABLED)
        self.log_area.delete(1.0, tk.END)
        self.update_status("正在分離音檔...", "warning")
        threading.Thread(target=self.batch_process, daemon=True).start()

    def batch_process(self):
        total, output_dir = len(self.file_list), self.output_dir_var.get()
        for i, input_file in enumerate(self.file_list):
            if not os.path.exists(input_file): continue
            self.log(f"\n--- 正在分離 ({i+1}/{total}): {os.path.basename(input_file)} ---")
            self.update_progress(int(i / total * 100), f"正在處理 {i+1}/{total}")
            
            if self.run_audio_separator(input_file, output_dir): self.log(f"✅ 完成: {os.path.basename(input_file)}")
            else: self.log(f"❌ 失敗: {os.path.basename(input_file)}")
        
        self.update_progress(100, "全部完成")
        self.root.after(0, lambda: messagebox.showinfo("成功", f"完成 {total} 個檔案。"))
        self.root.after(0, lambda: os.startfile(output_dir))
        self.finish_processing()

    # --- 底層共用函數 ---
    def show_welcome_message(self):
        self.log_area.insert(tk.END, "==================================================\n 🚀 KTV Studio 智慧影音處理中樞已啟動\n==================================================\n▶ 網址支援多筆貼上，系統將自動批次下載。\n▶ 初次使用請務必點擊左下方【⚙️ 系統初始化】下載引擎。\n==================================================\n")

    def log(self, message): self.root.after(0, lambda: self._safe_log(message))
    def _safe_log(self, message): self.log_area.insert(tk.END, message + "\n"); self.log_area.see(tk.END)

    def update_progress(self, percent, text=None): self.root.after(0, lambda: self._safe_update_progress(percent, text))
    def _safe_update_progress(self, percent, text):
        self.progress_bar['value'] = percent
        if text: self.status_var.set(f"{text} ({percent}%)")
        else: self.progress_text.config(text=f"{percent}%")

    def update_status(self, text, color="info"): self.root.after(0, lambda: self._safe_update_status(text, color))
    def _safe_update_status(self, text, color):
        self.status_var.set(text)
        self.status_label.config(bootstyle=color)

    def finish_processing(self):
        self.is_processing = False
        self.root.after(0, lambda: self.start_btn.config(state=tk.NORMAL))
        self.update_status("準備就緒", "success")
        self.update_progress(0, "0%")

    def consolidate_stems(self, input_audio, reference_video, output_dir):
        fmt, audio_stem, video_stem, out_path = self.output_format_var.get(), Path(input_audio).stem, Path(reference_video).stem, Path(output_dir)
        voc_final, inst_final = out_path / f"{video_stem}_人聲.{fmt}", out_path / f"{video_stem}_伴奏.{fmt}"
        
        for f in out_path.iterdir():
            if f.name.startswith(audio_stem) and "(Vocals)" in f.name and f.suffix == f".{fmt}":
                if voc_final.exists(): os.remove(voc_final)
                f.rename(voc_final); break
        
        found_inst = False
        for f in out_path.iterdir():
            if f.name.startswith(audio_stem) and any(x in f.name for x in ["(Instrumental)", "(No Vocals)"]) and f.suffix == f".{fmt}":
                if inst_final.exists(): os.remove(inst_final)
                f.rename(inst_final); found_inst = True; break
        
        if not found_inst:
            stems_to_merge = [f for f in out_path.iterdir() if f.name.startswith(audio_stem) and any(tag in f.name for tag in ["(Bass)", "(Drums)", "(Other)", "_Bass", "_Drums", "_Other"]) and f.suffix == f".{fmt}"]
            if stems_to_merge:
                inputs = []
                for f in stems_to_merge: inputs.extend(["-i", str(f)])
                filter_str = "".join([f"[{i}:a]" for i in range(len(stems_to_merge))]) + f"amix=inputs={len(stems_to_merge)}:duration=first[out]"
                try:
                    subprocess.run([str(self.bin_dir / "ffmpeg.exe"), "-y"] + inputs + ["-filter_complex", filter_str, "-map", "[out]", "-b:a", "320k", str(inst_final)], check=True, creationflags=self.subp_flags)
                    found_inst = True
                except: pass

        for f in out_path.iterdir():
            if f.name.startswith(audio_stem):
                try: f.unlink()
                except: pass
        if Path(input_audio).exists():
            try: os.remove(input_audio)
            except: pass
            
        return (str(voc_final) if voc_final.exists() else None, str(inst_final) if inst_final.exists() else None)

    def download_youtube(self, url, output_dir, mode="both"):
        video_id_match = re.search(r"(?:v=|\/)([0-9A-Za-z_-]{11})", url)
        video_id = video_id_match.group(1) if video_id_match else "temp_id"
        ytdlp_exe = self.py_dir / "Scripts" / "yt-dlp.exe"
        ytdlp_cmd_base = [str(ytdlp_exe)] if ytdlp_exe.exists() else [str(self.local_python), "-m", "yt_dlp"]
        common_opts = ["--no-playlist", "--ffmpeg-location", str(self.bin_dir), "--encoding", "utf-8", "--progress"]

        def run_ytdlp_with_logging(cmd):
            process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, bufsize=1, universal_newlines=True, creationflags=self.subp_flags, errors='replace')
            last_percent = -1
            for line in process.stdout:
                line = line.strip()
                if not line: continue
                if "[download]" in line and "%" in line:
                    match = re.search(r"(\d+\.\d+)%", line)
                    if match and int(float(match.group(1))) > last_percent:
                        self.log(f"    {line}"); last_percent = int(float(match.group(1)))
                elif any(x in line for x in ["[ffmpeg]", "Merging", "Extracting"]): self.log(f"    {line}")
            process.wait()
            return process.returncode == 0

        def find_file(pattern):
            files = list(Path(output_dir).glob(pattern))
            if files: files.sort(key=lambda x: os.path.getmtime(x), reverse=True); return str(files[0])
            return None

        video_file, audio_file = None, None
        if mode in ["both", "mp4"]:
            self.log("  > 下載 MP4 影片...")
            mp4_template = os.path.join(output_dir, f"%(title)s_#{video_id}#.%(ext)s")
            if run_ytdlp_with_logging(ytdlp_cmd_base + common_opts + ["-f", "bestvideo[ext=mp4][height<=1080]+bestaudio[ext=m4a]/best[ext=mp4]/best", "-o", mp4_template, url]): video_file = find_file(f"*#{video_id}#.*")

        if mode in ["both", "mp3"]:
            self.log("  > 下載 MP3 音訊...")
            mp3_template = os.path.join(output_dir, f"%(title)s_audio_#{video_id}#.%(ext)s")
            if run_ytdlp_with_logging(ytdlp_cmd_base + common_opts + ["-x", "--audio-format", "mp3", "--audio-quality", "320K", "-o", mp3_template, url]): audio_file = find_file(f"*_audio_#{video_id}#.mp3")
            if mode == "both" and video_file:
                try:
                    clean_v = str(Path(video_file).parent / Path(video_file).name.replace(f"_#{video_id}#", ""))
                    if os.path.exists(clean_v): os.remove(clean_v)
                    os.rename(video_file, clean_v); video_file = clean_v
                except: pass

        return (video_file, audio_file) if mode == "both" else (video_file if mode == "mp4" else audio_file)

    def synthesize_mkv(self, video_file, vocal_file, instrumental_file, output_file):
        if not (os.path.exists(video_file) and os.path.exists(vocal_file) and os.path.exists(instrumental_file)): return False
        try:
            subprocess.run([str(self.bin_dir / "ffmpeg.exe"), "-y", "-i", str(video_file), "-i", str(vocal_file), "-i", str(instrumental_file), "-filter_complex", "[1:a][2:a]amix=inputs=2:duration=first[mix]", "-map", "0:v", "-map", "[mix]", "-map", "2:a", "-c:v", "copy", "-c:a", "aac", "-b:a", "320k", "-metadata:s:a:0", "title=導唱 (人聲+伴奏)", "-metadata:s:a:1", "title=伴唱 (純伴奏)", str(output_file)], check=True, creationflags=self.subp_flags)
            return True
        except: return False

    def run_audio_separator(self, input_file, output_dir):
        self.fix_python_pth()
        fmt, device = self.output_format_var.get(), "cuda" if self.device_var.get() == "gpu" else "cpu"
        
        env = os.environ.copy()
        env["PYTHONPATH"] = str(self.lib_dir)
        env["PATH"] = os.pathsep.join([str(self.lib_dir)] + [str(p) for p in self.lib_dir.rglob("bin")] + [str(p) for p in self.lib_dir.rglob("lib")]) + os.pathsep + env.get("PATH", "")

        script = f"import sys, os\nsys.path.insert(0, r'{self.lib_dir}')\nif hasattr(os, 'add_dll_directory'):\n    lib_path = r'{self.lib_dir}'\n    for root, dirs, files in os.walk(lib_path):\n        if 'bin' in dirs or 'lib' in dirs:\n            for d in ['bin', 'lib']:\n                p = os.path.join(root, d)\n                if os.path.isdir(p):\n                    try: os.add_dll_directory(p)\n                    except: pass\nfrom audio_separator.utils.cli import main\nmain()"
        selected_model = self.model_var.get().split(" ")[0]
        is_demucs = selected_model.endswith(".yaml")
        command = [str(self.local_python), "-c", script, input_file, "-m", selected_model, "--model_file_dir", str(self.models_dir), "--output_dir", output_dir, "--output_format", fmt, "--output_bitrate", "320k", "--normalization", "0.9"]

        if is_demucs: command.extend(["--vr_batch_size", "1", "--vr_window_size", "512", "--vr_aggression", "5"])
        else:
            command.extend(["--mdx_overlap", str(self.overlap_var.get()), "--mdx_segment_size", "256", "--mdx_hop_length", "1024"])
            if self.denoise_var.get(): command.append("--mdx_enable_denoise")
        
        if device == "cuda":
            command.append("--use_autocast")
            if not is_demucs: command.extend(["--mdx_batch_size", "4"])
        else:
            if not is_demucs: command.extend(["--mdx_batch_size", "1"])
            
        try:
            process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, bufsize=1, universal_newlines=True, creationflags=self.subp_flags, env=env)
            error_detected = False
            for line in process.stdout:
                line = line.strip()
                if line:
                    self.log(f"    {line}")
                    if "no kernel image" in line: error_detected = True
            process.wait()
            if process.returncode == 0 and not error_detected:
                input_stem, out_path = Path(input_file).stem, Path(output_dir)
                if any(out_path.glob(f"{input_stem}_(Vocals)*.{fmt}")) or any(out_path.glob(f"{input_stem}_(Instrumental)*.{fmt}")): return True
            return False
        except: return False

    # --- 環境與依賴部署 ---
    def check_gpu_env(self):
        self.log("\n--- [開始 GPU 環境深度檢測] ---")
        if not self.local_python.exists():
            if messagebox.askyesno("初始化環境", "偵測到環境尚未初始化，是否現在進行？"): self.check_components(prompt=True)
            return

        env = os.environ.copy()
        env["PYTHONPATH"], env["PYTHONIOENCODING"] = str(self.lib_dir), "utf-8"
        env["PATH"] = os.pathsep.join([str(self.lib_dir)] + [str(p) for p in self.lib_dir.rglob("bin")] + [str(p) for p in self.lib_dir.rglob("lib")]) + os.pathsep + env.get("PATH", "")

        check_script = f"import sys, os\nsys.path.insert(0, r'{self.lib_dir}')\nif hasattr(os, 'add_dll_directory'):\n    lib_path = r'{self.lib_dir}'\n    for root, dirs, files in os.walk(lib_path):\n        if 'bin' in dirs or 'lib' in dirs:\n            for d in ['bin', 'lib']:\n                p = os.path.join(root, d)\n                if os.path.isdir(p):\n                    try: os.add_dll_directory(p)\n                    except: pass\ntry:\n    import onnxruntime as ort; print(f'[OK] ONNX Runtime: {{ort.__version__}}')\n    print(f'[OK] Providers: {{ort.get_available_providers()}}')\nexcept: print('[ERROR] ONNX Runtime failed')\ntry:\n    import torch; print(f'[OK] PyTorch: {{torch.__version__}}')\n    if torch.cuda.is_available():\n        torch.zeros(1).cuda()\n        print(f'[OK] GPU Ready: {{torch.cuda.get_device_name(0)}}')\n    else: print('[INFO] CPU Only')\nexcept: print('[ERROR] PyTorch failed')"
        try:
            res = subprocess.run([str(self.local_python), "-c", check_script], capture_output=True, text=True, env=env, encoding='utf-8', errors='replace', creationflags=self.subp_flags)
            self.log(res.stdout.strip())
            if "GPU Ready" not in res.stdout:
                if messagebox.askyesno("配置 CUDA", "偵測到 GPU 未就緒，是否執行一鍵修復？"): self._start_async_setup()
        except Exception as e: self.log(f"[ERROR] 檢測失敗: {str(e)}")

    def _start_async_setup(self):
        if not self.is_processing:
            self.is_processing = True
            self.update_status("正在執行一鍵修復...", "warning")
            threading.Thread(target=self._async_setup_environment, daemon=True).start()

    def check_components(self, prompt=True):
        if self.is_processing: return
        if prompt:
            if not messagebox.askyesno("確認下載", "這將下載 AI 核心 (約 800MB+)，是否開始？"): return
        else:
            if not self.local_python.exists() or not (self.bin_dir / "ffmpeg.exe").exists(): return
            if not (self.py_dir / "Scripts" / "yt-dlp.exe").exists(): threading.Thread(target=self._install_ytdlp_silent, daemon=True).start()
            if self._quick_check_gpu(): self.device_var.set("gpu")
            return
        self._start_async_setup()

    def _install_ytdlp_silent(self):
        try: subprocess.run([str(self.local_python), "-m", "pip", "install", "yt-dlp"], creationflags=self.subp_flags, check=True)
        except: pass

    def _async_setup_environment(self):
        self.log("--- 開始自動化環境部署 ---")
        if not self.local_python.exists():
            if not self.download_portable_python(): return self.finish_processing()
        else: self.fix_python_pth()

        if not (self.bin_dir / "ffmpeg.exe").exists(): self.download_ffmpeg()
        
        has_torch = (self.lib_dir / "torch").exists()
        has_sep = (self.lib_dir / "audio_separator").exists()
        packages_ok = False
        
        if has_torch and has_sep:
            try:
                check_cmd = f"import sys, os; sys.path.insert(0, r'{self.lib_dir}'); import torch; import onnxruntime as ort; ok = 'CUDAExecutionProvider' in ort.get_available_providers() and torch.cuda.is_available(); print('OK' if ok else 'FAIL')"
                res = subprocess.run([str(self.local_python), "-c", check_cmd], capture_output=True, text=True, creationflags=self.subp_flags)
                if "OK" in res.stdout: packages_ok = True
            except: pass
            
        if not packages_ok:
            if not self.install_packages_locally(): return self.finish_processing()
            
        if not (self.py_dir / "Scripts" / "yt-dlp.exe").exists(): self._install_ytdlp_silent()
        
        self.update_progress(100, "全部就緒")
        self.finish_processing()

    def fix_python_pth(self):
        try:
            pth_file = list(self.py_dir.glob("*._pth"))[0]
            with open(pth_file, "r") as f: lines = [l.strip() for l in f.readlines() if l.strip()]
            for item in ["python310.zip", ".", "Lib/site-packages", "..\\ai_libraries", "import site"]:
                if item not in lines:
                    if f"#{item}" in lines: lines[lines.index(f"#{item}")] = item
                    else: lines.append(item)
            with open(pth_file, "w") as f: f.write("\n".join(lines) + "\n")
        except: pass

    def download_portable_python(self):
        self.log("🚀 下載內建 Python 核心...")
        try:
            urllib.request.install_opener(urllib.request.build_opener(urllib.request.HTTPSHandler(context=ssl._create_unverified_context())))
            urllib.request.urlretrieve("https://www.python.org/ftp/python/3.10.11/python-3.10.11-embed-amd64.zip", self.py_dir / "py.zip")
            with zipfile.ZipFile(self.py_dir / "py.zip", 'r') as z: z.extractall(self.py_dir)
            self.fix_python_pth(); os.remove(self.py_dir / "py.zip")
            return True
        except: return False

    def download_ffmpeg(self):
        self.log("🚀 下載 FFmpeg...")
        try:
            urllib.request.install_opener(urllib.request.build_opener(urllib.request.HTTPSHandler(context=ssl._create_unverified_context())))
            urllib.request.urlretrieve("https://github.com/BtbN/FFmpeg-Builds/releases/download/latest/ffmpeg-master-latest-win64-gpl-shared.zip", self.bin_dir / "ffmpeg.zip")
            with zipfile.ZipFile(self.bin_dir / "ffmpeg.zip", 'r') as z:
                for f in z.namelist():
                    if "/bin/" in f.replace('\\', '/') and (f.endswith(".exe") or f.endswith(".dll")):
                        with z.open(f) as s, open(self.bin_dir / os.path.basename(f), "wb") as t: shutil.copyfileobj(s, t)
            os.remove(self.bin_dir / "ffmpeg.zip")
            return True
        except: return False

    def install_packages_locally(self):
        self.log("📦 安裝 AI 加速組件...")
        try:
            pip_script = self.py_dir / "get-pip.py"
            urllib.request.urlretrieve("https://bootstrap.pypa.io/get-pip.py", pip_script)
            subprocess.run([str(self.local_python), str(pip_script)], creationflags=self.subp_flags, check=True)
            
            steps = [
                ["--index-url", "https://pypi.tuna.tsinghua.edu.cn/simple", "setuptools", "wheel", "pip"], 
                ["--index-url", "https://pypi.tuna.tsinghua.edu.cn/simple", "nvidia-cuda-runtime-cu12", "nvidia-cudnn-cu12", "nvidia-cublas-cu12"],
                ["--index-url", "https://pypi.tuna.tsinghua.edu.cn/simple", "--extra-index-url", "https://download.pytorch.org/whl/cu124", "torch==2.5.1+cu124", "torchaudio==2.5.1+cu124", "onnxruntime-gpu", "audio-separator[gpu]"]
            ]
            for step in steps:
                subprocess.run([str(self.local_python), "-m", "pip", "install", "--target", str(self.lib_dir), "--upgrade"] + step, creationflags=self.subp_flags, check=True)
            if pip_script.exists(): os.remove(pip_script)
            return True
        except: return False

    def _quick_check_gpu(self):
        if not self.local_python.exists(): return False
        try: return "READY" in subprocess.run([str(self.local_python), "-c", f"import sys; sys.path.insert(0, r'{self.lib_dir}'); import torch, onnxruntime as ort; print('READY' if torch.cuda.is_available() and 'CUDAExecutionProvider' in ort.get_available_providers() else 'NO')"], capture_output=True, text=True, creationflags=self.subp_flags).stdout
        except: return False

if __name__ == "__main__":
    # 🚀 使用 cyborg 科技感深色主題
    root = ttk.Window(themename="cyborg") 
    app = AudioSeparatorApp(root)
    root.mainloop()