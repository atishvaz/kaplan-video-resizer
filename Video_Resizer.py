import customtkinter as ctk
import tkinter.filedialog as fd
import threading
import cv2
import pytesseract
import os
import re
import subprocess
import openpyxl
import time
import shutil 
from PIL import Image, ImageDraw, ImageOps, ImageTk

# Point to your Mac's Tesseract installation
pytesseract.pytesseract.tesseract_cmd = r'/opt/homebrew/bin/tesseract'

# ==========================================
# UI THEME SETTINGS
# ==========================================
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class JobTile(ctk.CTkFrame):
    def __init__(self, master, video_name, **kwargs):
        super().__init__(master, fg_color="#2b2b2b", corner_radius=8, **kwargs)
        
        self.grid_columnconfigure((0, 1, 2, 3), weight=1)
        self.start_time = None
        self.is_running = False
        
        self.lbl_title = ctk.CTkLabel(self, text=video_name, font=ctk.CTkFont(size=14, weight="bold"), anchor="w")
        self.lbl_title.grid(row=0, column=0, columnspan=4, padx=15, pady=(10, 0), sticky="w")
        
        self.lbl_status = ctk.CTkLabel(self, text="Waiting in queue...", text_color="#a3a3a3", font=ctk.CTkFont(size=12), anchor="w")
        self.lbl_status.grid(row=1, column=0, columnspan=4, padx=15, pady=0, sticky="w")
        
        self.progressbar = ctk.CTkProgressBar(self, height=8, corner_radius=4)
        self.progressbar.grid(row=2, column=0, columnspan=4, padx=15, pady=(5, 10), sticky="ew")
        self.progressbar.set(0)

        self.lbl_duration = ctk.CTkLabel(self, text="Video Duration: --:--", font=ctk.CTkFont(size=11), text_color="#a3a3a3")
        self.lbl_duration.grid(row=3, column=0, padx=15, pady=(0, 10), sticky="w")

        self.lbl_elapsed = ctk.CTkLabel(self, text="Elapsed: 00:00", font=ctk.CTkFont(size=11), text_color="#a3a3a3")
        self.lbl_elapsed.grid(row=3, column=1, padx=5, pady=(0, 10), sticky="w")

        self.lbl_stage = ctk.CTkLabel(self, text="Stage: Queued", font=ctk.CTkFont(size=11, weight="bold"), text_color="#3B8ED0")
        self.lbl_stage.grid(row=3, column=2, padx=5, pady=(0, 10), sticky="w")

        self.lbl_resized = ctk.CTkLabel(self, text="Resized: Pending", font=ctk.CTkFont(size=11), text_color="#a3a3a3")
        self.lbl_resized.grid(row=3, column=3, padx=15, pady=(0, 10), sticky="e")

    def set_duration(self, seconds):
        mins, secs = divmod(int(seconds), 60)
        self.lbl_duration.configure(text=f"Video Duration: {mins:02d}:{secs:02d}")

    def start_timer(self):
        self.start_time = time.time()
        self.is_running = True
        self.update_clock()

    def stop_timer(self):
        self.is_running = False

    def update_clock(self):
        if not self.is_running: return
        elapsed = int(time.time() - self.start_time)
        mins, secs = divmod(elapsed, 60)
        self.lbl_elapsed.configure(text=f"Elapsed: {mins:02d}:{secs:02d}")
        self.after(1000, self.update_clock)

    def update_tile(self, status_text, progress_val=None, color="#ffffff", stage=None, resized_status=None):
        self.lbl_status.configure(text=status_text, text_color=color)
        if progress_val is not None: self.progressbar.set(progress_val)
        if stage is not None: self.lbl_stage.configure(text=f"Stage: {stage}")
        if resized_status is not None: self.lbl_resized.configure(text=f"Resized: {resized_status}")


class VideoResizerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Video Resizer")
        self.geometry("1050x700")
        
        # --- NEW: App Icon Setup ---
        icon_path = "app_icon.png"
        if os.path.exists(icon_path):
            icon_img = ImageTk.PhotoImage(Image.open(icon_path))
            self.iconphoto(True, icon_img)
            
        self.excel_path = None
        self.video_folder = None
        self.bg_path = None
        self.output_video_folder = None
        self.col_indices = {} 
        
        self.user_choice = None
        self.wait_event = threading.Event()
        self.cancel_flag = False 

        # ==========================================
        # GLOBAL SETTINGS DICTIONARY
        # ==========================================
        self.cfg = {
            "shrink": 0.87,
            "bottom_pct": 0.12,
            "corner_pct": 0.15,
            "crf": 28,
            "preset": "slower",
            "time_buffer_start": 0.5,
            "max_bridge_gap": 15.0
        }

        # --- GRID LAYOUT ---
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # --- SIDEBAR ---
        self.sidebar_frame = ctk.CTkFrame(self, width=330, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(8, weight=1) 

        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="Video Resizer", font=ctk.CTkFont(size=26, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(30, 20))

        btn_params = {"width": 270, "height": 45, "font": ctk.CTkFont(size=13, weight="bold"), "corner_radius": 8}

        self.btn_select_excel = ctk.CTkButton(self.sidebar_frame, text="1. Select Existing Excel", command=self.select_excel, **btn_params)
        self.btn_select_excel.grid(row=1, column=0, padx=20, pady=10)

        self.btn_select_folder = ctk.CTkButton(self.sidebar_frame, text="2. Select Video Folder", command=self.select_video_folder, **btn_params)
        self.btn_select_folder.grid(row=2, column=0, padx=20, pady=10)

        self.btn_select_bg = ctk.CTkButton(self.sidebar_frame, text="3. Select Background Image", command=self.select_background, **btn_params)
        self.btn_select_bg.grid(row=3, column=0, padx=20, pady=10)

        self.lbl_limit = ctk.CTkLabel(self.sidebar_frame, text="Batch Limit (Pause after X videos):", font=ctk.CTkFont(weight="bold"))
        self.lbl_limit.grid(row=4, column=0, padx=20, pady=(20, 0), sticky="w")
        
        self.entry_limit = ctk.CTkEntry(self.sidebar_frame, placeholder_text="e.g., 5 or 10", width=270)
        self.entry_limit.insert(0, "5") 
        self.entry_limit.grid(row=5, column=0, padx=20, pady=(5, 10))

        self.compress_var = ctk.IntVar(value=1)
        self.switch_compress = ctk.CTkSwitch(self.sidebar_frame, text="Enable Compression", variable=self.compress_var, font=ctk.CTkFont(size=12, weight="bold"))
        self.switch_compress.grid(row=6, column=0, padx=20, pady=(15, 10), sticky="w")

        self.btn_settings = ctk.CTkButton(self.sidebar_frame, text="⚙️ SETTINGS & PREVIEW", fg_color="#333333", hover_color="#444444", command=self.open_settings_window, **btn_params)
        self.btn_settings.grid(row=8, column=0, padx=20, pady=(10, 10))

        self.btn_run = ctk.CTkButton(self.sidebar_frame, text="▶ START BATCH", fg_color="#28a745", hover_color="#218838", text_color="white", **btn_params)
        self.btn_run.configure(command=self.start_pipeline_thread)
        self.btn_run.grid(row=9, column=0, padx=20, pady=(10, 30))

        # --- MAIN AREA ---
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")
        self.main_frame.grid_rowconfigure(1, weight=1) 
        self.main_frame.grid_columnconfigure(0, weight=1)

        self.queue_label = ctk.CTkLabel(self.main_frame, text="Batch Queue (Sequential)", font=ctk.CTkFont(size=20, weight="bold"))
        self.queue_label.grid(row=0, column=0, sticky="w", pady=(0, 10))

        self.queue_frame = ctk.CTkScrollableFrame(self.main_frame, fg_color="#1e1e1e", corner_radius=10)
        self.queue_frame.grid(row=1, column=0, sticky="nsew")

    # ==========================================
    # SETTINGS & LIVE PREVIEW WINDOW
    # ==========================================
    def open_settings_window(self):
        settings_win = ctk.CTkToplevel(self)
        settings_win.title("Settings & Live Preview")
        settings_win.geometry("900x550")
        settings_win.attributes('-topmost', True)
        settings_win.grab_set() 

        settings_win.grid_columnconfigure(0, weight=1)
        settings_win.grid_columnconfigure(1, weight=2)
        settings_win.grid_rowconfigure(0, weight=1)

        ctrl_frame = ctk.CTkFrame(settings_win, fg_color="transparent")
        ctrl_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        
        ctk.CTkLabel(ctrl_frame, text="Layout Controls", font=ctk.CTkFont(size=15, weight="bold")).pack(anchor="w", pady=(0, 15))

        self.lbl_shrink_val = ctk.CTkLabel(ctrl_frame, text=f"Shrink Factor: {int(self.cfg['shrink']*100)}%")
        self.lbl_shrink_val.pack(anchor="w")
        self.slider_shrink = ctk.CTkSlider(ctrl_frame, from_=0.5, to=1.0, command=self.update_live_preview)
        self.slider_shrink.set(self.cfg['shrink'])
        self.slider_shrink.pack(fill="x", pady=(0, 20))

        self.lbl_bot_val = ctk.CTkLabel(ctrl_frame, text=f"Scan Area (Bottom): {int(self.cfg['bottom_pct']*100)}%")
        self.lbl_bot_val.pack(anchor="w")
        self.slider_bot = ctk.CTkSlider(ctrl_frame, from_=0.05, to=0.30, command=self.update_live_preview)
        self.slider_bot.set(self.cfg['bottom_pct'])
        self.slider_bot.pack(fill="x", pady=(0, 20))

        self.lbl_corn_val = ctk.CTkLabel(ctrl_frame, text=f"Corner Ignore Mask: {int(self.cfg['corner_pct']*100)}%")
        self.lbl_corn_val.pack(anchor="w")
        self.slider_corn = ctk.CTkSlider(ctrl_frame, from_=0.0, to=0.40, command=self.update_live_preview)
        self.slider_corn.set(self.cfg['corner_pct'])
        self.slider_corn.pack(fill="x", pady=(0, 20))

        ctk.CTkLabel(ctrl_frame, text="Compression Settings", font=ctk.CTkFont(size=15, weight="bold")).pack(anchor="w", pady=(15, 15))
        
        self.lbl_crf_val = ctk.CTkLabel(ctrl_frame, text=f"CRF (Lower = Better Qty): {self.cfg['crf']}")
        self.lbl_crf_val.pack(anchor="w")
        self.slider_crf = ctk.CTkSlider(ctrl_frame, from_=18, to=35, number_of_steps=17, command=self.update_crf_label)
        self.slider_crf.set(self.cfg['crf'])
        self.slider_crf.pack(fill="x", pady=(0, 20))

        ctk.CTkLabel(ctrl_frame, text="Encoding Preset:").pack(anchor="w")
        self.combo_preset = ctk.CTkComboBox(ctrl_frame, values=["ultrafast", "superfast", "veryfast", "faster", "fast", "medium", "slow", "slower", "veryslow"])
        self.combo_preset.set(self.cfg['preset'])
        self.combo_preset.pack(fill="x", pady=(0, 20))

        btn_save = ctk.CTkButton(ctrl_frame, text="Save Settings", fg_color="#28a745", hover_color="#218838", command=lambda: self.save_settings(settings_win))
        btn_save.pack(fill="x", side="bottom", pady=10)

        preview_frame = ctk.CTkFrame(settings_win, fg_color="#1e1e1e", corner_radius=10)
        preview_frame.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")
        
        ctk.CTkLabel(preview_frame, text="Live Canvas Preview (16:9)", font=ctk.CTkFont(weight="bold")).pack(pady=(10, 5))
        
        legend_frame = ctk.CTkFrame(preview_frame, fg_color="transparent")
        legend_frame.pack(pady=(0, 10))
        ctk.CTkLabel(legend_frame, text="■ Shrink Layout", text_color="#a855f7", font=ctk.CTkFont(weight="bold")).pack(side="left", padx=10)
        ctk.CTkLabel(legend_frame, text="■ OCR Watch Zone", text_color="#3B8ED0", font=ctk.CTkFont(weight="bold")).pack(side="left", padx=10)

        self.preview_lbl = ctk.CTkLabel(preview_frame, text="")
        self.preview_lbl.pack(expand=True, padx=20, pady=20)
        
        self.update_live_preview()

    def update_crf_label(self, val):
        self.lbl_crf_val.configure(text=f"CRF (Lower = Better Qty): {int(val)}")

    def update_live_preview(self, *args):
        shrink = self.slider_shrink.get()
        bot = self.slider_bot.get()
        corn = self.slider_corn.get()

        self.lbl_shrink_val.configure(text=f"Shrink Factor: {int(shrink*100)}%")
        self.lbl_bot_val.configure(text=f"Scan Area (Bottom): {int(bot*100)}%")
        self.lbl_corn_val.configure(text=f"Corner Ignore Mask: {int(corn*100)}%")

        w, h = 640, 360 
        
        if self.bg_path and os.path.exists(self.bg_path):
            try:
                bg_img = Image.open(self.bg_path).convert("RGBA")
                img = ImageOps.fit(bg_img, (w, h), Image.Resampling.LANCZOS)
            except Exception:
                img = Image.new('RGBA', (w, h), color='#2b2b2b')
        else:
            img = Image.new('RGBA', (w, h), color='#2b2b2b')

        draw = ImageDraw.Draw(img, "RGBA")

        # Draw Shrink Box
        sw, sh = w * shrink, h * shrink
        sx0, sy0 = (w - sw) / 2, 0 
        sx1, sy1 = sx0 + sw, sy0 + sh
        draw.rectangle([sx0, sy0, sx1, sy1], outline='#a855f7', width=4)
        draw.text((sx0 + 10, sy0 + 10), "Video Boundaries", fill="#a855f7")

        # Draw OCR Watch Zone
        wy0 = h * (1 - bot)
        wx0 = w * corn
        wx1 = w * (1 - corn)
        draw.rectangle([wx0, wy0, wx1, h], fill=(59, 142, 208, 60), outline='#3B8ED0', width=3)
        draw.text((wx0 + 10, wy0 + 10), "OCR Watch Area", fill="#3B8ED0")

        ctk_img = ctk.CTkImage(light_image=img, dark_image=img, size=(w, h))
        self.preview_lbl.configure(image=ctk_img)
        self.preview_lbl.image = ctk_img

    def save_settings(self, win):
        self.cfg["shrink"] = self.slider_shrink.get()
        self.cfg["bottom_pct"] = self.slider_bot.get()
        self.cfg["corner_pct"] = self.slider_corn.get()
        self.cfg["crf"] = int(self.slider_crf.get())
        self.cfg["preset"] = self.combo_preset.get()
        win.destroy()

    # ==========================================
    # FILE SELECTION & MAIN LOGIC
    # ==========================================
    def select_excel(self):
        self.excel_path = fd.askopenfilename(title='Select Excel Tracking File', filetypes=[('Excel files', '*.xlsx')])
        if self.excel_path: self.btn_select_excel.configure(text="1. Excel Linked ✓", fg_color="#3a5a40")

    def select_video_folder(self):
        self.video_folder = fd.askdirectory(title='Select Folder Containing All Videos')
        if self.video_folder:
            self.output_video_folder = os.path.join(self.video_folder, "Final_Render_Folder")
            if not os.path.exists(self.output_video_folder): os.makedirs(self.output_video_folder)
            self.btn_select_folder.configure(text="2. Folder Linked ✓", fg_color="#3a5a40")

    def select_background(self):
        self.bg_path = fd.askopenfilename(title='Select Background Image', filetypes=[('Image files', '*.jpg *.jpeg *.png')])
        if self.bg_path: self.btn_select_bg.configure(text="3. Background Linked ✓", fg_color="#3a5a40")

    def cancel_batch(self):
        self.cancel_flag = True
        self.btn_run.configure(state="disabled", text="CANCELLING...")

    def start_pipeline_thread(self):
        if not self.excel_path or not self.video_folder or not self.bg_path: return
        try:
            batch_limit = int(self.entry_limit.get())
            if batch_limit <= 0: raise ValueError
        except ValueError: return

        self.cancel_flag = False
        
        self.btn_run.configure(text="⏹ CANCEL BATCH", fg_color="#ff4c4c", hover_color="#cc0000", command=self.cancel_batch)
        self.btn_select_excel.configure(state="disabled")
        self.btn_select_folder.configure(state="disabled")
        self.btn_select_bg.configure(state="disabled")
        self.entry_limit.configure(state="disabled")
        self.switch_compress.configure(state="disabled")
        self.btn_settings.configure(state="disabled")
        
        threading.Thread(target=self.run_master_controller, args=(batch_limit,), daemon=True).start()

    def ask_to_continue(self):
        self.user_choice = None
        self.wait_event.clear()
        self.after(0, self._create_popup)
        self.wait_event.wait() 
        return self.user_choice

    def _create_popup(self):
        popup = ctk.CTkToplevel(self)
        popup.title("Batch Complete")
        popup.geometry("350x180")
        popup.attributes('-topmost', True)
        popup.grab_set() 
        
        lbl = ctk.CTkLabel(popup, text="Current batch finished safely.\nWould you like to process the next batch?", font=ctk.CTkFont(size=14))
        lbl.pack(pady=30)
        
        btn_frame = ctk.CTkFrame(popup, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20)
        
        def on_yes():
            self.user_choice = True
            self.wait_event.set()
            popup.destroy()
            
        def on_no():
            self.user_choice = False
            self.wait_event.set()
            popup.destroy()

        ctk.CTkButton(btn_frame, text="Continue", fg_color="#28a745", hover_color="#218838", command=on_yes, width=120).pack(side="left", padx=10)
        ctk.CTkButton(btn_frame, text="Stop Here", fg_color="#ff4c4c", hover_color="#cc0000", command=on_no, width=120).pack(side="right", padx=10)
        popup.protocol("WM_DELETE_WINDOW", on_no)

    def setup_excel_headers(self, sheet, wb):
        # NEW: Added 'Original Resolution' to the dynamic headers
        headers_to_add = ["QA Status", "Original Resolution", "Audio Status", "Overlap Timestamps", "Obscured Text"]
        max_col = 1
        while sheet.cell(row=1, column=max_col).value is not None: max_col += 1

        existing_headers = {}
        for col in range(1, max_col):
            val = sheet.cell(row=1, column=col).value
            if val in headers_to_add: existing_headers[val] = col

        next_blank = max_col
        needs_save = False
        for header in headers_to_add:
            if header in existing_headers:
                self.col_indices[header] = existing_headers[header]
            else:
                sheet.cell(row=1, column=next_blank, value=header)
                self.col_indices[header] = next_blank
                next_blank += 1
                needs_save = True
        if needs_save: wb.save(self.excel_path)

    def run_master_controller(self, batch_limit):
        try:
            wb = openpyxl.load_workbook(self.excel_path)
            sheet = wb.active
            self.setup_excel_headers(sheet, wb)
            
            all_pending_jobs = []
            for row in range(2, sheet.max_row + 1):
                raw_title = sheet.cell(row=row, column=1).value
                if not raw_title: continue
                
                # Convert cell to string safely to prevent errors on empty cells
                current_status = str(sheet.cell(row=row, column=self.col_indices["QA Status"]).value or "")
                
                # If the word "Fixed" is ANYWHERE in the status, or if it is "Clear", skip it
                if "Fixed" in current_status or current_status == "Clear": 
                    continue 
                    
                all_pending_jobs.append({"row": row, "title": raw_title})

            if not all_pending_jobs: return

            for i in range(0, len(all_pending_jobs), batch_limit):
                if self.cancel_flag: break

                current_batch = all_pending_jobs[i : i + batch_limit]
                for widget in self.queue_frame.winfo_children(): widget.destroy()
                
                job_tiles = {}
                for job in current_batch:
                    tile = JobTile(self.queue_frame, video_name=job['title'])
                    tile.pack(fill="x", pady=5)
                    job_tiles[job['row']] = tile

                for job in current_batch:
                    if self.cancel_flag: break
                    self.process_single_video(job, job_tiles[job['row']])
                
                if self.cancel_flag: break

                if i + batch_limit < len(all_pending_jobs):
                    if not self.ask_to_continue(): break

        except Exception as e:
            print(f"Master Error: {str(e)}")
        finally:
            self.btn_run.configure(state="normal", text="▶ START BATCH", fg_color="#28a745", hover_color="#218838", command=self.start_pipeline_thread)
            self.btn_select_excel.configure(state="normal", fg_color=["#3B8ED0", "#1F6AA5"])
            self.btn_select_folder.configure(state="normal", fg_color=["#3B8ED0", "#1F6AA5"])
            self.btn_select_bg.configure(state="normal", fg_color=["#3B8ED0", "#1F6AA5"])
            self.entry_limit.configure(state="normal")
            self.switch_compress.configure(state="normal")
            self.btn_settings.configure(state="normal")

    def process_single_video(self, job, tile):
        if self.cancel_flag: 
            tile.update_tile("Cancelled", 1.0, color="#ff4c4c", stage="Aborted", resized_status="N/A")
            return

        row = job['row']
        raw_title = job['title']
        
        tile.start_timer()
        tile.update_tile("Locating file...", 0.05, "#ffcc00", stage="Initialization")
        video_path = self.find_video_file(raw_title)

        if not video_path:
            self.save_to_excel(row, qa_status="File Not Found", resolution="N/A")
            tile.stop_timer()
            tile.update_tile("File Not Found", 1.0, color="#ff4c4c", stage="Halted", resized_status="Error")
            return

        # Fetch duration and mathematically calculate exact resolution
        cap = cv2.VideoCapture(video_path)
        fps = cap.get(cv2.CAP_PROP_FPS) or 30
        frames = cap.get(cv2.CAP_PROP_FRAME_COUNT)
        vid_w = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
        vid_h = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
        resolution_str = f"{vid_w}x{vid_h}"
        
        if fps > 0 and frames > 0:
            tile.set_duration(frames / fps)
        cap.release()

        video_filename = os.path.basename(video_path)
        tile.update_tile("Checking Audio...", 0.1, "#ffcc00", stage="Audio Check")
        audio_status = self.check_audio(video_path)

        intervals = self.scan_video(video_path, tile)
        
        if getattr(self, 'cancel_flag', False):
            tile.stop_timer()
            tile.update_tile("Cancelled during scan", 1.0, color="#ff4c4c", stage="Aborted", resized_status="N/A")
            return

        if intervals is None:
            self.save_to_excel(row, qa_status="Error Reading Video", resolution=resolution_str)
            tile.stop_timer()
            tile.update_tile("Read Error", 1.0, color="#ff4c4c", stage="Halted", resized_status="Error")
            return
        elif len(intervals) == 0:
            tile.update_tile("Copying Clean Video...", 0.9, "#3B8ED0", stage="File Transfer")
            shutil.copy2(video_path, os.path.join(self.output_video_folder, video_filename))
            self.save_to_excel(row, qa_status="Clear", resolution=resolution_str, audio_status=audio_status, timestamps="None", text_data="None")
            tile.stop_timer()
            tile.update_tile("Clear & Copied", 1.0, color="#28a745", stage="Complete", resized_status="No (Original Copied)")
            return
            
        all_timestamps = ", ".join([f"{self.format_timestamp(i['start'])}-{self.format_timestamp(i['end'])}" for i in intervals])
        all_texts = " | ".join([text for i in intervals for text in i['texts']])

        success = self.render_and_compress(video_path, intervals, tile)
        
        if success:
            is_compressed = self.compress_var.get() == 1
            success_msg = "Fixed & Compressed" if is_compressed else "Fixed without compression"
            self.save_to_excel(row, qa_status=success_msg, resolution=resolution_str, audio_status=audio_status, timestamps=all_timestamps, text_data=all_texts)
            tile.stop_timer()
            tile.update_tile(success_msg, 1.0, color="#28a745", stage="Complete", resized_status="Yes")
        else:
            self.save_to_excel(row, qa_status="Render Error", resolution=resolution_str, audio_status=audio_status)
            tile.stop_timer()
            tile.update_tile("Render Failed", 1.0, color="#ff4c4c", stage="Halted", resized_status="Error")

    def save_to_excel(self, row, qa_status, resolution="N/A", audio_status="N/A", timestamps=None, text_data=None):
        try:
            wb = openpyxl.load_workbook(self.excel_path)
            sheet = wb.active
            sheet.cell(row=row, column=self.col_indices["QA Status"], value=qa_status)
            sheet.cell(row=row, column=self.col_indices["Original Resolution"], value=resolution)
            sheet.cell(row=row, column=self.col_indices["Audio Status"], value=audio_status)
            if timestamps is not None: sheet.cell(row=row, column=self.col_indices["Overlap Timestamps"], value=timestamps)
            if text_data is not None: sheet.cell(row=row, column=self.col_indices["Obscured Text"], value=text_data)
            wb.save(self.excel_path)
        except Exception as e: print(f"Excel Save Error: {e}")

    def normalize_string(self, s):
        if not s: return ""
        return re.sub(r'[^a-zA-Z0-9]', '', str(s)).lower()

    def find_video_file(self, excel_title):
        normalized_title = self.normalize_string(excel_title).replace('mp4', '')
        for filename in os.listdir(self.video_folder):
            if filename.lower().endswith(('.mp4', '.mov', '.avi')):
                normalized_filename = self.normalize_string(filename).replace('mp4', '').replace('mov', '').replace('avi', '')
                if normalized_title == normalized_filename or normalized_title in normalized_filename: return os.path.join(self.video_folder, filename)
        return None

    def check_audio(self, video_path):
        try:
            cmd = ["ffprobe", "-v", "error", "-show_entries", "stream=codec_type", "-of", "csv=p=0", video_path]
            return "Has Audio" if "audio" in subprocess.check_output(cmd, text=True).strip() else "No Audio"
        except Exception: return "Audio Check Error"

    def format_timestamp(self, seconds):
        mins, secs = divmod(int(seconds), 60)
        return f"{mins:02d}:{secs:02d}"

    def scan_video(self, video_path, tile):
        cap = cv2.VideoCapture(video_path)
        if not cap.isOpened(): return None

        fps = cap.get(cv2.CAP_PROP_FPS) or 30
        total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
        
        bot_pct = self.cfg["bottom_pct"]
        corn_pct = self.cfg["corner_pct"]
        frame_interval = int(fps * self.cfg.get("CHECK_INTERVAL", 1.0))
        
        frame_count, current_start = 0, 0
        raw_intervals, current_texts = [], []
        is_overlapping = False

        while cap.isOpened():
            if getattr(self, 'cancel_flag', False):
                cap.release()
                return None
                
            ret, frame = cap.read()
            if not ret: break

            if frame_count % frame_interval == 0:
                if total_frames > 0:
                    prog = 0.1 + (0.5 * (frame_count / total_frames))
                    tile.update_tile(f"Scanning ({self.format_timestamp(frame_count/fps)})...", prog, "#3B8ED0", stage="OCR Analysis")

                h, w, _ = frame.shape
                
                subtitle_zone = frame[int(h * (1 - bot_pct)):h, 0:w]
                gray_zone = cv2.cvtColor(subtitle_zone, cv2.COLOR_BGR2GRAY)
                
                mask_w = int(w * corn_pct)
                gray_zone[:, :mask_w] = 255           
                gray_zone[:, w - mask_w:] = 255   
                
                # 1. OPTICAL FIX: Melt away fabric/skin wrinkles so OCR can't see them
                blurred_zone = cv2.medianBlur(gray_zone, 3)
                
                _, thresh_zone = cv2.threshold(blurred_zone, 150, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
                detected_text = pytesseract.image_to_string(thresh_zone).strip()
                
                text_without_kaplan = re.sub(r'(?i)kaplan', '', detected_text)
                
                # 2. LOGIC FIX: Must contain at least one actual 3-letter continuous word
                valid_words = re.findall(r'[a-zA-Z]{3,}', text_without_kaplan)
                text_present = len(valid_words) > 0 

                if text_present:
                    clean_text = detected_text.replace('\n', ' ').strip()
                    if clean_text not in current_texts: current_texts.append(clean_text)

                if text_present and not is_overlapping:
                    is_overlapping = True
                    current_start = frame_count / fps
                elif not text_present and is_overlapping:
                    is_overlapping = False
                    raw_intervals.append({'start': current_start, 'end': frame_count / fps, 'texts': current_texts})
                    current_texts = [] 

            frame_count += 1

        if is_overlapping: raw_intervals.append({'start': current_start, 'end': (frame_count / fps), 'texts': current_texts})
        cap.release()
        
        if not raw_intervals: return []

        raw_intervals.sort(key=lambda x: x['start'])
        smoothed_intervals = []
        
        current_interval = raw_intervals[0]
        current_interval['start'] = max(0, current_interval['start'] - self.cfg["time_buffer_start"])

        for next_interval in raw_intervals[1:]:
            next_start_buffered = max(0, next_interval['start'] - self.cfg["time_buffer_start"])
            gap = next_start_buffered - current_interval['end']
            
            if gap <= self.cfg["max_bridge_gap"]:
                current_interval['end'] = next_interval['end']
                current_interval['texts'] = list(set(current_interval['texts'] + next_interval['texts']))
            else:
                smoothed_intervals.append(current_interval)
                current_interval = next_interval
                current_interval['start'] = next_start_buffered

        smoothed_intervals.append(current_interval)
        return smoothed_intervals

    def render_and_compress(self, input_path, intervals, tile):
        video_filename = os.path.basename(input_path)
        base_name, ext = os.path.splitext(video_filename)
        
        rendered_temp = os.path.join(self.output_video_folder, f"TEMP_{video_filename}")
        final_output = os.path.join(self.output_video_folder, f"{base_name}_resized{ext}")

        cap = cv2.VideoCapture(input_path)
        vid_w = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
        vid_h = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
        cap.release()

        enable_expr = "+".join([f"between(t,{i['start']},{i['end']})" for i in intervals])
        shrink = self.cfg["shrink"]
        
        filter_complex = (
            f"[1:v]scale=1920:1080:force_original_aspect_ratio=increase,crop=1920:1080[bg];"
            f"[0:v]scale=1920:1080:force_original_aspect_ratio=decrease,pad=1920:1080:(ow-iw)/2:(oh-ih)/2:color=black[orig_standard];"
            f"[orig_standard]split=2[orig][to_shrink];"
            f"[to_shrink]scale=iw*{shrink}:ih*{shrink}[shrunk];"
            f"[bg][shrunk]overlay=(W-w)/2:0[padded];"
            f"[orig][padded]overlay=enable='{enable_expr}'[v]"
        )

        try:
            do_compression = self.compress_var.get() == 1

            if do_compression:
                tile.update_tile("Building layout...", 0.65, "#a855f7", stage="Hardware Render")
                subprocess.run([
                    "ffmpeg", "-y", "-i", input_path, "-loop", "1", "-t", "1", "-i", self.bg_path,
                    "-filter_complex", filter_complex, "-map", "[v]", "-map", "0:a?",
                    "-c:v", "h264_videotoolbox", "-b:v", "4M", "-c:a", "copy", rendered_temp
                ], check=True, capture_output=True)

                if self.cancel_flag: return False 

                tile.update_tile("Crunching file size...", 0.85, "#a855f7", stage="Software Compression")
                subprocess.run([
                    "ffmpeg", "-y", "-i", rendered_temp, "-c:v", "libx264", 
                    "-crf", str(self.cfg["crf"]), 
                    "-preset", self.cfg["preset"], 
                    "-c:a", "copy", final_output
                ], check=True, capture_output=True)

                if os.path.exists(rendered_temp): os.remove(rendered_temp)
            else:
                tile.update_tile("Fast Hardware Render...", 0.75, "#a855f7", stage="Fast Render")
                subprocess.run([
                    "ffmpeg", "-y", "-i", input_path, "-loop", "1", "-t", "1", "-i", self.bg_path,
                    "-filter_complex", filter_complex, "-map", "[v]", "-map", "0:a?",
                    "-c:v", "h264_videotoolbox", "-b:v", "4M", "-c:a", "copy", final_output
                ], check=True, capture_output=True)

            return True
        except subprocess.CalledProcessError: 
            return False

if __name__ == "__main__":
    app = VideoResizerApp()
    app.mainloop()