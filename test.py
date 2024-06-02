import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from datetime import datetime, timedelta
import time
import tkinter as tk
import requests
import subprocess
from tkinter import filedialog, messagebox
import threading
import os
import sys
from urllib.parse import urljoin

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("PHẦN MỀM TỰ ĐỘNG HÓA THAO TÁC HỆ THỐNG MỘT CỬA ĐIỆN TỬ")
        self.root.geometry("700x450")

        self.file_path = ""
        self.stop_flag = False

        # Thiết lập font chữ tổng thể
        self.default_font = ("Helvetica", 12)
        self.button_font = ("Helvetica", 12, "bold")
        self.title_font = ("Helvetica", 16, "bold")
        self.signature_font = ("Helvetica",9)

        # Tạo nút "Check Update"
        self.check_update_button = tk.Button(root, text="Check Update", command=self.check_for_update)
        self.check_update_button.pack(anchor='se', padx=5, pady=5)
        
        # Tiêu đề
        self.title_label = tk.Label(root, text="PHẦN MỀM TỰ ĐỘNG HÓA THAO TÁC\nHỆ THỐNG MỘT CỬA ĐIỆN TỬ", fg="red", font=self.title_font)
        self.title_label.pack(pady=10)
        

        # Khung chứa phần chọn file
        self.file_frame = tk.Frame(root)
        self.file_frame.pack(pady=10, padx=10, fill=tk.X)

        self.file_entry = tk.Entry(self.file_frame, width=50, font=self.default_font)
        self.file_entry.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

        self.open_button = tk.Button(self.file_frame, text="Open", command=self.open_file, font=self.default_font, fg="blue")
        self.open_button.pack(side=tk.LEFT, padx=5)
        
        

        # Khung chứa các checkbutton
        self.checkbox_frame = tk.Frame(root)
        self.checkbox_frame.pack(pady=10, padx=10, anchor='w')

        self.checkbox_value3 = tk.IntVar()
        self.checkbox_value4 = tk.IntVar()
        self.checkbox_value5 = tk.IntVar()
        self.checkbox_value6 = tk.IntVar()

        self.checkbox1 = tk.Checkbutton(self.checkbox_frame, text="Đồng bộ Thủ Tục Chung", variable=self.checkbox_value3, font=self.default_font)
        self.checkbox1.pack(anchor='w')

        self.checkbox2 = tk.Checkbutton(self.checkbox_frame, text="Đồng bộ Dịch Vụ Công", variable=self.checkbox_value4, font=self.default_font)
        self.checkbox2.pack(anchor='w')

        self.checkbox4 = tk.Checkbutton(self.checkbox_frame, text="Đồng bộ Lĩnh Vực", variable=self.checkbox_value6, font=self.default_font)
        self.checkbox4.pack(anchor='w')

        self.checkbox3 = tk.Checkbutton(self.checkbox_frame, text="Cấu hình Ngày Nghỉ Lễ (cho năm sau)", variable=self.checkbox_value5, font=self.default_font)
        self.checkbox3.pack(anchor='w')

        # Các nút điều khiển
        self.control_frame = tk.Frame(root)
        self.control_frame.pack(pady=10)

        self.start_button = tk.Button(self.control_frame, text="START", command=self.start_thread, font=self.button_font, fg="green")
        self.start_button.grid(row=0, column=0, padx=5)

        self.stop_button = tk.Button(self.control_frame, text="STOP", command=self.stop_automation, state=tk.DISABLED, font=self.button_font, fg="red")
        self.stop_button.grid(row=0, column=1, padx=5)

        # Button to download sample file
        self.download_button = tk.Button(root, text="Tải file mẫu", command=self.download_sample_file, font=self.button_font, fg="blue")
        self.download_button.pack(pady=5)

        # Chữ ký
        self.signature_label = tk.Label(root, text="Phòng Ứng dụng CNTT - Trung tâm Công nghệ thông tin và Truyền thông",fg="purple", font=self.signature_font, anchor='e')
        self.signature_label.pack(side=tk.BOTTOM, padx=10, pady=10, anchor='se')
        

        
    
    def check_for_update(self):
        # URL của tệp cập nhật trên GitHub
        repo_owner = "duy97ct"
        repo_name = "motcua_auto"
        file_path = "Motcua_auto.exe"  # Đường dẫn tệp trên GitHub

        # Đường dẫn lưu tệp trên hệ thống (cùng thư mục với file .exe)
        if getattr(sys, 'frozen', False):
            application_path = os.path.dirname(sys.executable)
        else:
            application_path = os.path.dirname(os.path.abspath(__file__))
            
        save_path = os.path.join(application_path, "Motcua_auto.exe")
        
        # URL trực tiếp tới tệp .exe trên GitHub
        url = f"https://github.com/{repo_owner}/{repo_name}/raw/main/{file_path}"
        
        response = requests.get(url, stream=True)
        
        if response.status_code == 200:
            if messagebox.askyesno("Cập nhật", "Có bản cập nhật mới. Bạn có muốn tải xuống và cài đặt không?"):
                with open(save_path, 'wb') as file:
                    for chunk in response.iter_content(chunk_size=8192):
                        file.write(chunk)
                messagebox.showinfo("Thông báo", "Cập nhật thành công. Vui lòng khởi động lại ứng dụng.")
                self.root.destroy()
            else:
                messagebox.showinfo("Thông báo", "Đã hủy cập nhật.")
        else:
            messagebox.showerror("Lỗi", "Không thể tải tệp cập nhật. Vui lòng thử lại sau.")


    def download_and_replace_update(self):
        try:
            # URL tải về tệp cập nhật .exe
            update_url = "https://github.com/duy97ct/motcua_auto/raw/main/Motcua_auto.exe"
            
            # Đường dẫn tới tệp .exe hiện tại
            current_exe_path = os.path.abspath(__file__)

            # Tạo một tên tạm cho tệp tải về
            temp_exe_path = current_exe_path + ".temp"

            # Tải xuống tệp cập nhật
            response = requests.get(update_url, stream=True)
            response.raise_for_status()

            # Lưu tệp cập nhật tạm thời
            with open(temp_exe_path, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)

            # Thay thế tệp hiện tại bằng tệp cập nhật
            os.replace(temp_exe_path, current_exe_path)

            # Hiển thị thông báo và thoát ứng dụng
            messagebox.showinfo("Cập nhật", "Đã tải xuống cập nhật. Vui lòng khởi động lại ứng dụng.")
            self.root.destroy()
            
            # Khởi động lại ứng dụng
            subprocess.Popen([current_exe_path])

        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể cập nhật ứng dụng: {e}")

    def open_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, file_path)
            self.file_path = file_path

    def start_thread(self):
        self.stop_flag = False
        self.stop_button.config(state=tk.NORMAL)
        threading.Thread(target=self.start_automation).start()

    def stop_automation(self):
        self.stop_flag = True

    def start_automation(self):
        # Đọc tệp Excel
        df = pd.read_excel(self.file_path)

        # Thiết lập Chrome WebDriver
        service = Service("chromedriver.exe")
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        driver = webdriver.Chrome(service=service, options=options)
        wait = WebDriverWait(driver, 10)

        # Lấy tên file để xác định các trường dữ liệu
        filename = os.path.basename(self.file_path)
        if filename == "Sample1.xlsx":
            url_col = "URL1"
            username_col = "Username1"
            password_col = "Password1"
        elif filename == "Sample2.xlsx":
            url_col = "URL2"
            username_col = "Username2"
            password_col = "Password2"
        else:
            messagebox.showerror("Lỗi", "Tên file không khớp với mẫu đã định nghĩa.")
            return

        # Duyệt qua từng hàng và thực hiện các tác vụ
        for index, row in df.iterrows():
            if self.stop_flag:
                break

            url = row[url]
            username = row[admin]
            password = row[pass]

            # Mở URL và thực hiện đăng nhập
            driver.get(url)
            wait.until(EC.presence_of_element_located((By.NAME, "username"))).send_keys(username)
            driver.find_element(By.NAME, "password").send_keys(password)
            driver.find_element(By.NAME, "password").send_keys(Keys.RETURN)

            # Thực hiện các thao tác theo yêu cầu
            if self.checkbox_value3.get():
                self.sync_thu_tuc_chung(driver)
            if self.checkbox_value4.get():
                self.sync_dich_vu_cong(driver)
            if self.checkbox_value5.get():
                self.configure_holidays(driver)
            if self.checkbox_value6.get():
                self.sync_fields(driver)

        driver.quit()
        self.stop_button.config(state=tk.DISABLED)

    def sync_thu_tuc_chung(self, driver):
        # Thực hiện đồng bộ thủ tục chung
        pass

    def sync_dich_vu_cong(self, driver):
        # Thực hiện đồng bộ dịch vụ công
        pass

    def configure_holidays(self, driver):
        # Thực hiện cấu hình ngày nghỉ lễ
        pass
    
    def sync_fields(self, driver):
        # Thực hiện đồng bộ lĩnh vực
        pass
    
    def download_sample_file(self):
        url = "https://example.com/sample.xlsx"
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        
        if save_path:
            response = requests.get(url)
            with open(save_path, 'wb') as f:
                f.write(response.content)
            messagebox.showinfo("Thành công", "Đã tải xuống tệp mẫu thành công.")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
