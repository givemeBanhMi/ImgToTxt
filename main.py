# pyrefly: ignore [missing-import]
import customtkinter as ctk
import tkinter.ttk as ttk
from tkinter import filedialog, messagebox
import threading
import os
import glob
import time
import concurrent.futures, queue
from config_manager import load_config, save_config
from ocr_service import extract_passport_data
from excel_service import ExcelMatcher, save_results_to_excel

# All columns for the Treeview
TREE_COLUMNS = (
    "STT", "Tên File Ảnh", "Số Passport", "Họ Tên", "Ngày Sinh",
    "Giới Tính", "Quốc Tịch", "Nơi Sinh", "Ngày Cấp",
    "Ngày Hết Hạn", "Cơ Quan Cấp", "Dòng MRZ", "Trạng Thái"
)

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("HỆ THỐNG OCR HỘ CHIẾU & ĐỐI CHIẾU DỮ LIỆU TỰ ĐỘNG")
        self.geometry("1280x750")
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self.config_data = load_config()
        self.use_offline = self.config_data.get("use_offline", False)
        self.output_folder = os.path.join(os.getcwd(), "Outputs")
        self.output_excel_path = ""
        self.results_data = []

        self.create_widgets()
        self.load_initial_data()
        # Initialize pause control and timer for future use
        self.pause_event = threading.Event()
        self.pause_event.set()  # start as unpaused
        self.is_paused = False
        self.is_processing = False
        self.start_time = None  # will be set when processing starts

    def toggle_pause(self):
        if self.is_paused:
            # Resume
            self.is_paused = False
            self.pause_event.set()
            self.pause_btn.configure(text="Tạm dừng")
        else:
            # Pause
            self.is_paused = True
            self.pause_event.clear()
            self.pause_btn.configure(text="Tiếp tục")

    def update_timer_label(self):
        if self.start_time and self.is_processing:
            elapsed = int(time.time() - self.start_time)
            self.timer_label.configure(text=f"Thời gian: {elapsed}s")
            self.after(1000, self.update_timer_label)

    def create_widgets(self):
        # ── Header ──
        self.header_frame = ctk.CTkFrame(self, fg_color="#1a4c82", corner_radius=0)
        self.header_frame.pack(fill="x")
        ctk.CTkLabel(
            self.header_frame,
            text="HỆ THỐNG OCR HỘ CHIẾU & ĐỐI CHIẾU DỮ LIỆU TỰ ĐỘNG",
            font=ctk.CTkFont(size=20, weight="bold"), text_color="white"
        ).pack(pady=10)
        ctk.CTkLabel(
            self.header_frame,
            text="Phiên bản Doanh nghiệp - Tích hợp 9router / OpenAI",
            font=ctk.CTkFont(size=12, slant="italic"), text_color="white"
        ).pack(pady=(0, 10))

        # ── Main Container ──
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # ── Config Section ──
        cfg = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        cfg.pack(fill="x", pady=10)

        labels = [
            "API Key:",
            "API Base URL (9router):",
            "Model Name:",
            "Thư mục Hình ảnh:",
            "File Excel SecureScan:",
            "Số luồng (1‑10):",
            "Chế độ OCR:"
        ]

        for i, lbl in enumerate(labels):
            ctk.CTkLabel(cfg, text=lbl, width=170, anchor="w").grid(row=i, column=0, padx=5, pady=5)

        # API Key
        self.api_key_entry = ctk.CTkEntry(cfg, show="*", width=420, placeholder_text="Dán API Key của bạn vào đây...")
        self.api_key_entry.grid(row=0, column=1, padx=5, pady=5)
        ctk.CTkButton(cfg, text="Hiện/Ẩn", width=80, command=self.toggle_key_visibility).grid(row=0, column=2, padx=5, pady=5)

        # API Base URL
        self.api_url_entry = ctk.CTkEntry(cfg, width=420, placeholder_text="VD: https://api.9router.com/v1")
        self.api_url_entry.grid(row=1, column=1, padx=5, pady=5)

        # Model Name
        self.model_name_entry = ctk.CTkEntry(cfg, width=420, placeholder_text="VD: google/gemini-1.5-flash")
        self.model_name_entry.grid(row=2, column=1, padx=5, pady=5)

        # Image Folder
        self.img_folder_entry = ctk.CTkEntry(cfg, width=420, placeholder_text="Dán đường dẫn thư mục ảnh hoặc bấm Chọn...")
        self.img_folder_entry.grid(row=3, column=1, padx=5, pady=5)
        ctk.CTkButton(cfg, text="Chọn...", width=80, command=self.browse_img_folder).grid(row=3, column=2, padx=5, pady=5)

        # Excel File
        self.excel_file_entry = ctk.CTkEntry(cfg, width=420, placeholder_text="Không bắt buộc - chọn file Excel nếu muốn đối chiếu")
        self.excel_file_entry.grid(row=4, column=1, padx=5, pady=5)
        ctk.CTkButton(cfg, text="Chọn...", width=80, command=self.browse_excel_file).grid(row=4, column=2, padx=5, pady=5)

        # Thread count entry (default 4)
        self.thread_entry = ctk.CTkEntry(cfg, width=80, placeholder_text="4")
        self.thread_entry.grid(row=5, column=1, padx=5, pady=5, sticky="w")
        ctk.CTkLabel(cfg, text="Chỉ dùng 1 API của 9router", anchor="w").grid(row=5, column=1, padx=(95, 5), pady=5, sticky="w")

        # OCR mode switch
        self.offline_switch = ctk.CTkSwitch(
            cfg,
            text="Offline OCR (EasyOCR)",
            command=self.toggle_offline_mode
        )
        self.offline_switch.grid(row=6, column=1, padx=5, pady=5, sticky="w")
        self.ocr_mode_label = ctk.CTkLabel(cfg, text="", anchor="w")
        self.ocr_mode_label.grid(row=6, column=1, padx=(200, 5), pady=5, sticky="w")


        # ── Progress Section ──
        prog = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        prog.pack(fill="x", pady=10)

        # Timer label (elapsed time)
        self.timer_label = ctk.CTkLabel(prog, text="Thời gian: 0s", anchor="w")
        self.timer_label.pack(fill="x", padx=5, pady=(0,5))

        self.status_label = ctk.CTkLabel(prog, text="Sẵn sàng.", anchor="w")
        self.status_label.pack(fill="x", padx=5)

        self.progress_bar = ctk.CTkProgressBar(prog)
        self.progress_bar.pack(fill="x", padx=5, pady=5)
        self.progress_bar.set(0)

        # ── Buttons Section ──
        btns = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        btns.pack(fill="x", pady=10)

        self.start_btn = ctk.CTkButton(btns, text="Bắt đầu Nhận dạng & Đối chiếu",
                                       fg_color="green", hover_color="darkgreen",
                                       command=self.start_processing)
        self.start_btn.pack(side="left", padx=5)

        # Pause/Resume button (initially Paused = False)
        self.pause_btn = ctk.CTkButton(btns, text="Tạm dừng",
                                       fg_color="orange", hover_color="darkorange",
                                       command=self.toggle_pause)
        self.pause_btn.pack(side="left", padx=5)

        self.open_excel_btn = ctk.CTkButton(btns, text="Mở File Excel Kết quả",
                                            state="disabled", command=self.open_results_excel)
        self.open_excel_btn.pack(side="left", padx=5)

        ctk.CTkButton(btns, text="Mở Thư mục Outputs",
                       command=self.open_outputs_folder).pack(side="left", padx=5)

        ctk.CTkButton(btns, text="Thoát", fg_color="gray", hover_color="darkgray",
                       command=self.destroy).pack(side="right", padx=5)

        # ── Treeview (Results Table) ──
        tree_container = ctk.CTkFrame(self.main_frame)
        tree_container.pack(fill="both", expand=True, pady=10)

        self.tree = ttk.Treeview(tree_container, columns=TREE_COLUMNS, show="headings")

        col_widths = {
            "STT": 40,   "Tên File Ảnh": 180, "Số Passport": 110,
            "Họ Tên": 150, "Ngày Sinh": 90, "Giới Tính": 70,
            "Quốc Tịch": 90, "Nơi Sinh": 120, "Ngày Cấp": 90,
            "Ngày Hết Hạn": 100, "Cơ Quan Cấp": 120, "Dòng MRZ": 180,
            "Trạng Thái": 180
        }
        for col in TREE_COLUMNS:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor="center", width=col_widths.get(col, 100))

        # Horizontal scrollbar
        xscroll = ttk.Scrollbar(tree_container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(xscrollcommand=xscroll.set)
        xscroll.pack(side="bottom", fill="x")

        # Vertical scrollbar
        yscroll = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=yscroll.set)
        yscroll.pack(side="right", fill="y")

        self.tree.pack(fill="both", expand=True)

    # ─── Helpers ───────────────────────────────────────────────

    def load_initial_data(self):
        self.api_key_entry.insert(0, self.config_data.get("api_key", ""))
        self.api_url_entry.insert(0, self.config_data.get("api_base_url", ""))
        self.model_name_entry.insert(0, self.config_data.get("model_name", ""))
        self.img_folder_entry.insert(0, self.config_data.get("image_folder", ""))
        self.excel_file_entry.insert(0, self.config_data.get("excel_file", ""))
        self.thread_entry.insert(0, str(self.config_data.get("thread_count", 4)))
        if self.use_offline:
            self.offline_switch.select()
        else:
            self.offline_switch.deselect()
        self.update_ocr_mode_label()

    def toggle_offline_mode(self):
        self.use_offline = bool(self.offline_switch.get())
        self.update_ocr_mode_label()
        self.save_current_config()

    def update_ocr_mode_label(self):
        if self.use_offline:
            self.ocr_mode_label.configure(text="Đang dùng Offline", text_color="green")
        else:
            self.ocr_mode_label.configure(text="Đang dùng Online API", text_color="#1a4c82")


    def save_current_config(self):
        save_config(
            api_key=self.api_key_entry.get().strip(),
            api_base_url=self.api_url_entry.get().strip(),
            model_name=self.model_name_entry.get().strip(),
            image_folder=self.img_folder_entry.get().strip(),
            excel_file=self.excel_file_entry.get().strip(),
            thread_count=int(self.thread_entry.get().strip() or 4),
            use_offline=self.use_offline
        )


    def toggle_key_visibility(self):
        if self.api_key_entry.cget("show") == "*":
            self.api_key_entry.configure(show="")
        else:
            self.api_key_entry.configure(show="*")

    def browse_img_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.img_folder_entry.delete(0, "end")
            self.img_folder_entry.insert(0, folder)
            self.save_current_config()

    def browse_excel_file(self):
        f = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if f:
            self.excel_file_entry.delete(0, "end")
            self.excel_file_entry.insert(0, f)
            self.save_current_config()

    def update_status(self, text, progress_val=None):
        self.status_label.configure(text=text)
        if progress_val is not None:
            self.progress_bar.set(progress_val)
        self.update_idletasks()

    # ─── Processing ────────────────────────────────────────────

    def start_processing(self):
        api_key    = self.api_key_entry.get().strip()
        api_url    = self.api_url_entry.get().strip()
        model_name = self.model_name_entry.get().strip()
        img_folder = self.img_folder_entry.get().strip()
        excel_path = self.excel_file_entry.get().strip()

        # Validate required fields
        if self.use_offline:
            if not img_folder:
                messagebox.showwarning("Thiếu thông tin",
                                       "Vui lòng chọn Thư mục ảnh.")
                return
        elif not api_key or not model_name or not img_folder:
            messagebox.showwarning("Thiếu thông tin",
                                   "Vui lòng nhập đầy đủ API Key, Model Name và Thư mục ảnh.")
            return
        if not os.path.exists(img_folder):
            messagebox.showerror("Lỗi", "Thư mục hình ảnh không tồn tại.")
            return

        try:
            thread_count = int(self.thread_entry.get().strip() or 4)
        except ValueError:
            thread_count = 4
        thread_count = max(1, min(10, thread_count))
        self.thread_entry.delete(0, "end")
        self.thread_entry.insert(0, str(thread_count))

        self.save_current_config()

        for item in self.tree.get_children():
            self.tree.delete(item)
        self.results_data.clear()
        self.open_excel_btn.configure(state="disabled")
        self.start_btn.configure(state="disabled")

        self.pause_event = threading.Event()
        self.pause_event.set()
        self.is_paused = False
        self.is_processing = True
        self.pause_btn.configure(text="Tạm dừng")
        self.start_time = time.time()
        self.update_timer_label()

        worker_thread = threading.Thread(
            target=self.process_files,
            args=(api_key, api_url, model_name, img_folder, excel_path, thread_count),
            daemon=True
        )
        worker_thread.start()


    def process_files(self, api_key, api_url, model_name, img_folder, excel_path, thread_count):
        # Load Excel data if provided
        matcher = None
        if excel_path and os.path.exists(excel_path):
            self.after(0, self.update_status, "Đang tải dữ liệu Excel...", None)
            matcher = ExcelMatcher(excel_path)
        else:
            self.after(0, self.update_status, "Không dùng Excel, chỉ OCR ảnh.", None)

        image_files = []
        for ext in ('*.png', '*.jpg', '*.jpeg', '*.PNG', '*.JPG', '*.JPEG'):
            image_files.extend(glob.glob(os.path.join(img_folder, ext)))
        image_files = list(dict.fromkeys(image_files))

        total = len(image_files)
        if total == 0:
            self.is_processing = False
            self.after(0, self.update_status, "Hoàn thành! Không tìm thấy ảnh nào.", 1.0)
            self.after(0, lambda: self.start_btn.configure(state="normal"))
            return

        results_lock = threading.Lock()
        completed_count = 0
        rate_limit_warned = False

        def process_one(idx, img_path):
            self.pause_event.wait()
            file_name = os.path.basename(img_path)
            self.after(0, self.update_status, f"Đang xử lý: {file_name}...", None)

            retries = 2
            result = None
            while retries > 0:
                self.pause_event.wait()
                result = extract_passport_data(img_path, api_key, api_url, model_name, use_offline=self.use_offline)
                status = result.get("trang_thai", "")
                if "Rate Limit" in status or "Quá tải" in status or "429" in status:
                    time.sleep(2)
                    retries -= 1
                else:
                    break

            if result is None:
                result = {"trang_thai": "OCR Thất bại"}

            # --- Double Check Logic ---
            double_check_status = result.get("trang_thai", "OCR Thất bại")
            if double_check_status == "Thành công":
                self.after(0, self.update_status, f"Đang Double Check: {file_name}...", None)
                retries2 = 2
                result2 = None
                while retries2 > 0:
                    self.pause_event.wait()
                    result2 = extract_passport_data(img_path, api_key, api_url, model_name, use_offline=self.use_offline)
                    status2 = result2.get("trang_thai", "")
                    if "Rate Limit" in status2 or "Quá tải" in status2 or "429" in status2:
                        time.sleep(2)
                        retries2 -= 1
                    else:
                        break
                
                if result2 and result2.get("trang_thai") == "Thành công":
                    # So sánh số passport và ngày cấp
                    if result.get("so_passport") == result2.get("so_passport") and result.get("ngay_cap") == result2.get("ngay_cap"):
                        double_check_status = "Verified"
                    else:
                        double_check_status = "Double Check Không khớp"
                else:
                    double_check_status = "Lỗi khi Double Check"

            # --- Tích hợp Excel Matcher (nếu có) ---
            passport_num = result.get("so_passport", "")
            if passport_num and matcher:
                is_match = matcher.match_passport(passport_num)
                match_status = "Khớp Excel" if is_match else "Không có trong Excel"
                final_status = f"{double_check_status} - {match_status}"
            else:
                final_status = double_check_status

            return {
                "stt": idx + 1,
                "ten_file": file_name,
                "so_passport": passport_num or "LỖI",
                "ho_ten": result.get("ho_ten", ""),
                "ngay_sinh": result.get("ngay_sinh", ""),
                "gioi_tinh": result.get("gioi_tinh", ""),
                "quoc_tich": result.get("quoc_tich", ""),
                "noi_sinh": result.get("noi_sinh", ""),
                "ngay_cap": result.get("ngay_cap", ""),
                "ngay_het_han": result.get("ngay_het_han", ""),
                "co_quan_cap": result.get("co_quan_cap", ""),
                "dong_mrz": result.get("dong_mrz", ""),
                "trang_thai": final_status
            }

        with concurrent.futures.ThreadPoolExecutor(max_workers=thread_count) as executor:
            future_map = {
                executor.submit(process_one, idx, img_path): img_path
                for idx, img_path in enumerate(image_files)
            }
            for future in concurrent.futures.as_completed(future_map):
                try:
                    row = future.result()
                except Exception as exc:
                    row = {
                        "stt": len(self.results_data) + 1,
                        "ten_file": os.path.basename(future_map[future]),
                        "so_passport": "LỖI",
                        "ho_ten": "",
                        "ngay_sinh": "",
                        "gioi_tinh": "",
                        "quoc_tich": "",
                        "noi_sinh": "",
                        "ngay_cap": "",
                        "ngay_het_han": "",
                        "co_quan_cap": "",
                        "dong_mrz": "",
                        "trang_thai": f"Lỗi OCR: {exc}"
                    }

                with results_lock:
                    completed_count += 1
                    self.results_data.append(row)
                    progress = completed_count / total

                if "Rate Limit" in row["trang_thai"] and not rate_limit_warned:
                    rate_limit_warned = True
                    self.after(0, lambda: messagebox.showwarning(
                        "Cảnh báo Rate Limit",
                        "API đang bị rate limit/quá tải. Hãy giảm số luồng hoặc kiểm tra gói API."
                    ))

                self.after(0, self.add_tree_row, row)
                self.after(0, self.update_status, f"Đã xử lý {completed_count}/{total} ảnh...", progress)

        self.results_data.sort(key=lambda item: item["stt"])
        output_file, success = save_results_to_excel(self.results_data, self.output_folder)
        if success:
            self.output_excel_path = output_file
            self.after(0, lambda: self.open_excel_btn.configure(state="normal"))
            self.after(0, self.update_status, f"Hoàn thành! Đã lưu kết quả tại {output_file}", 1.0)
        else:
            self.after(0, self.update_status, f"Hoàn thành xử lý ảnh, nhưng lỗi lưu Excel: {output_file}", 1.0)

        self.is_processing = False
        self.after(0, lambda: self.start_btn.configure(state="normal"))

    def add_tree_row(self, row):
        self.tree.insert("", "end", values=(
            row["stt"], row["ten_file"], row["so_passport"],
            row["ho_ten"], row["ngay_sinh"], row["gioi_tinh"],
            row["quoc_tich"], row["noi_sinh"], row["ngay_cap"],
            row["ngay_het_han"], row["co_quan_cap"], row["dong_mrz"],
            row["trang_thai"]
        ))

    def open_results_excel(self):
        if self.output_excel_path and os.path.exists(self.output_excel_path):
            try:
                os.startfile(self.output_excel_path)
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể mở file: {e}")

    def open_outputs_folder(self):
        if not os.path.exists(self.output_folder):
            os.makedirs(self.output_folder)
        try:
            os.startfile(self.output_folder)
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể mở thư mục: {e}")


if __name__ == "__main__":
    app = App()
    app.mainloop()
