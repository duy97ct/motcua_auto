import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from tkinter import ttk, filedialog, messagebox
import tkinter as tk
import os
import sys

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("PHẦN MỀM TỰ ĐỘNG HÓA THAO TÁC HỆ THỐNG MỘT CỬA ĐIỆN TỬ")
        self.root.geometry("700x550")

        self.file_path = ""
        self.stop_flag = False

        # Thiết lập font chữ tổng thể
        self.default_font = ("Helvetica", 13)
        self.button_font = ("Helvetica", 12, "bold")
        self.title_font = ("Helvetica", 16, "bold")
        self.signature_font = ("Helvetica", 9)

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

        self.checkbox_dongboTTC = tk.IntVar() #đồng bộ TTC
        self.checkbox_dongboDVC = tk.IntVar() #đồng bộ DVC
        self.checkbox_dongbolv = tk.IntVar() #đồng bộ lĩnh vực
        self.checkbox_offnamsau = tk.IntVar() #Cấu hình nghỉ lễ năm sau
        self.checkbox_offnamnay = tk.IntVar() #Cấu hình nghỉ lễ  năm nay
        self.checkbox_copysovb = tk.IntVar() #sao chép sổ vb

        self.checkbox1 = tk.Checkbutton(self.checkbox_frame, text="Đồng bộ Thủ Tục Chung", variable=self.checkbox_dongboTTC, font=self.default_font)
        self.checkbox1.pack(anchor='w')

        self.checkbox2 = tk.Checkbutton(self.checkbox_frame, text="Đồng bộ Dịch Vụ Công", variable=self.checkbox_dongboDVC, font=self.default_font)
        self.checkbox2.pack(anchor='w')

        self.checkbox4 = tk.Checkbutton(self.checkbox_frame, text="Đồng bộ Lĩnh Vực", variable=self.checkbox_dongbolv, font=self.default_font)
        self.checkbox4.pack(anchor='w')

        self.checkbox4 = tk.Checkbutton(self.checkbox_frame, text="Cấu hình Ngày Nghỉ Lễ", variable=self.checkbox_offnamnay, font=self.default_font)
        self.checkbox4.pack(anchor='w')
             
        # Khung chứa checkbutton "Cấu hình quy trình tự động" và nút "Đính kèm"
        self.checkbox_quytrinh_frame = tk.Frame(self.checkbox_frame)
        self.checkbox_quytrinh_frame.pack(anchor='w')

        self.checkbox_quytrinh = tk.IntVar()
        self.checkbox6 = tk.Checkbutton(self.checkbox_quytrinh_frame, text="Cấu hình quy trình tự động", variable=self.checkbox_quytrinh, font=self.default_font)
        self.checkbox6.pack(side=tk.LEFT)

        self.attach_button = tk.Button(self.checkbox_quytrinh_frame, text="Đính kèm", command=self.attach_file, font=self.default_font)
        self.attach_button.pack(side=tk.LEFT, padx=5)
        
        self.attached_file_label = tk.Label(self.checkbox_quytrinh_frame, text="", font=self.default_font)
        self.attached_file_label.pack(side=tk.LEFT)
        
        self.attached_file_path = None

        self.checkbox5 = tk.Checkbutton(self.checkbox_frame, text="Sao chép Sổ Văn Bản", variable=self.checkbox_copysovb, font=self.default_font)
        self.checkbox5.pack(anchor='w')     

        # Các nút điều khiển
        self.control_frame = tk.Frame(root)
        self.control_frame.pack(pady=10)

        self.start_button = tk.Button(self.control_frame, text="START", command=self.start_thread, font=self.button_font, fg="green")
        self.start_button.grid(row=0, column=0, padx=5)

        self.stop_button = tk.Button(self.control_frame, text="STOP", command=self.stop_automation, state=tk.DISABLED, font=self.button_font, fg="red")
        self.stop_button.grid(row=0, column=1, padx=5)

        # Khung chứa các nút tải file
        self.download_frame = tk.Frame(root)
        self.download_frame.pack(pady=5)

        # Nút "Tải file mẫu"
        self.download_button = tk.Button(self.download_frame, text="Tải file mẫu", command=self.download_sample_file, font=self.button_font, fg="blue")
        self.download_button.pack(side=tk.LEFT, padx=5)

        # Nút "Tải file quy trình"
        self.download_button_quytrinh = tk.Button(self.download_frame, text="Cấu hình quy trình", command=self.open_quy_trinh_window, font=self.button_font, fg="blue")
        self.download_button_quytrinh.pack(side=tk.LEFT, padx=5)

        # Chữ ký
        self.signature_label = tk.Label(root, text="Phòng Ứng dụng CNTT - Trung tâm Công nghệ thông tin và Truyền thông",fg="purple", font=self.signature_font, anchor='e')
        self.signature_label.pack(side=tk.BOTTOM, padx=10, pady=10, anchor='se')

    def open_quy_trinh_window(self):
        quy_trinh_window = tk.Toplevel(self.root)
        quy_trinh_window.title("CẤU HÌNH QUY TRÌNH")
        quy_trinh_window.geometry("900x600")

        self.quy_trinh_data = []
        self.form_entries = []  # Khởi tạo form_entries
        self.luan_chuyen_entries = []  # Khởi tạo luan_chuyen_entries

        # Khung chứa Quy trình
        self.quy_trinh_frame = tk.LabelFrame(quy_trinh_window, text="Quy trình", font=self.default_font)
        self.quy_trinh_frame.pack(fill="x", padx=10, pady=5)

        tk.Label(self.quy_trinh_frame, text="Tên quy trình:", font=self.default_font).grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.ten_quy_trinh_entry = tk.Entry(self.quy_trinh_frame, font=self.default_font, width=40)
        self.ten_quy_trinh_entry.grid(row=0, column=1, padx=5, pady=5)

        tk.Label(self.quy_trinh_frame, text="Bí danh quy trình:", font=self.default_font).grid(row=1, column=0, padx=5, pady=5, sticky='w')
        self.bi_danh_entry = tk.Entry(self.quy_trinh_frame, font=self.default_font, width=40)
        self.bi_danh_entry.grid(row=1, column=1, padx=5, pady=5)

        # Khung chứa Danh sách Form
        self.danh_sach_form_frame = tk.LabelFrame(quy_trinh_window, text="Danh sách Form", font=self.default_font)
        self.danh_sach_form_frame.pack(fill="x", padx=10, pady=5)

        # Thêm tiêu đề của các cột
        headers = ["ID", "Tên Form", "Mã Action", "Thời gian", "Nhóm người dùng"]
        for col, header in enumerate(headers):
            tk.Label(self.danh_sach_form_frame, text=header, font=self.default_font).grid(row=0, column=col, padx=5, pady=5)

        self.add_form_entry(1)  # Thêm form entry đầu tiên

        self.add_form_button = tk.Button(self.danh_sach_form_frame, text="Thêm 1 nhóm", command=self.add_form_entry, font=self.button_font)
        self.add_form_button.grid(row=999, column=0, columnspan=5, pady=5)

        # Nút "Xong"
        self.done_button = tk.Button(self.danh_sach_form_frame, text="Xong", command=self.save_form_state, font=self.button_font)
        self.done_button.grid(row=999, column=4, columnspan=5, pady=5)

        # Khung chứa Danh sách Luân chuyển
        self.danh_sach_luan_chuyen_frame = tk.LabelFrame(quy_trinh_window, text="Danh sách Luân chuyển", font=self.default_font)
        self.danh_sach_luan_chuyen_frame.pack(fill="x", padx=10, pady=5)

        self.add_luan_chuyen_entry()  # Thêm luân chuyển entry đầu tiên

        self.add_luan_chuyen_button = tk.Button(self.danh_sach_luan_chuyen_frame, text="Thêm 1 hàng", command=self.add_luan_chuyen_entry, font=self.button_font)
        self.add_luan_chuyen_button.grid(row=999, column=0, columnspan=3, pady=5)

        # Nút tải về
        self.download_button = tk.Button(quy_trinh_window, text="Tải về", command=self.export_to_excel, font=self.button_font)
        self.download_button.pack(pady=5)

    def add_form_entry(self, id=None):
        if id is None:
            id = len(self.form_entries) + 1

        row = len(self.form_entries) + 1

        id_label = tk.Label(self.danh_sach_form_frame, text=str(id), font=self.default_font)
        id_label.grid(row=row, column=0, padx=5, pady=5)

        ten_form_entry = tk.Entry(self.danh_sach_form_frame, font=self.default_font)
        ten_form_entry.grid(row=row, column=1, padx=5, pady=5)

        action_menu = ttk.Combobox(self.danh_sach_form_frame, values=["thêm mới", "chuyển xử lý", "trình phê duyệt", "chuyển ban hành", "chuyển trả kết quả"], font=self.default_font, state='readonly')
        action_menu.grid(row=row, column=2, padx=5, pady=5)

        thoi_gian_entry = tk.Entry(self.danh_sach_form_frame, font=self.default_font)
        thoi_gian_entry.grid(row=row, column=3, padx=5, pady=5)

        nhom_nguoi_dung_entry = tk.Entry(self.danh_sach_form_frame, font=self.default_font)
        nhom_nguoi_dung_entry.grid(row=row, column=4, padx=5, pady=5)

        self.form_entries.append((id_label, ten_form_entry, action_menu, thoi_gian_entry, nhom_nguoi_dung_entry))

        self.update_luan_chuyen_menus()

    def add_luan_chuyen_entry(self):
        row = len(self.luan_chuyen_entries) + 1

        from_form_label = tk.Label(self.danh_sach_luan_chuyen_frame, text="Từ form", font=self.default_font)
        from_form_label.grid(row=row, column=0, padx=5, pady=5)

        from_form_menu = ttk.Combobox(self.danh_sach_luan_chuyen_frame, values=[form[1].get() for form in self.form_entries], font=self.default_font, state='readonly')
        from_form_menu.grid(row=row, column=1, padx=5, pady=5)

        to_form_label = tk.Label(self.danh_sach_luan_chuyen_frame, text="Đến form", font=self.default_font)
        to_form_label.grid(row=row, column=2, padx=5, pady=5)

        to_form_menu = ttk.Combobox(self.danh_sach_luan_chuyen_frame, values=[form[1].get() for form in self.form_entries], font=self.default_font, state='readonly')
        to_form_menu.grid(row=row, column=3, padx=5, pady=5)

        self.luan_chuyen_entries.append((from_form_label, from_form_menu, to_form_label, to_form_menu))

    def update_luan_chuyen_menus(self):
        form_names = [form[1].get() for form in self.form_entries]
        for entry in self.luan_chuyen_entries:
            entry[1]['values'] = form_names
            entry[3]['values'] = form_names

    def save_form_state(self):
        self.update_luan_chuyen_menus()
        messagebox.showinfo("Thông báo", "Trạng thái của danh sách Form đã được lưu và cập nhật danh sách luân chuyển.")

    def export_to_excel(self):
        ten_quy_trinh = self.ten_quy_trinh_entry.get()
        bi_danh = self.bi_danh_entry.get()

        quy_trinh_data = []
        for entry in self.form_entries:
            id = entry[0].cget("text")
            ten_form = entry[1].get()
            action = entry[2].get()
            thoi_gian = entry[3].get()
            nhom_nguoi_dung = entry[4].get()
            quy_trinh_data.append([ten_quy_trinh, bi_danh, id, ten_form, action, thoi_gian, nhom_nguoi_dung])

        luan_chuyen_data = []
        for entry in self.luan_chuyen_entries:
            from_form = entry[1].get()
            to_form = entry[3].get()
            luan_chuyen_data.append([from_form, to_form])

        df_quy_trinh = pd.DataFrame(quy_trinh_data, columns=["Tên quy trình", "Bí danh", "ID", "Tên Form", "Mã Action", "Thời gian", "Nhóm người dùng"])
        df_luan_chuyen = pd.DataFrame(luan_chuyen_data, columns=["Từ form", "Đến form"])

        wb = Workbook()
        ws_quy_trinh = wb.active
        ws_quy_trinh.title = "QuyTrinh"

        # Thêm tên quy trình và bí danh vào tiêu đề
        ws_quy_trinh.append(["Tên quy trình:", ten_quy_trinh])
        ws_quy_trinh.append(["Bí danh:", bi_danh])
        ws_quy_trinh.append([])  # Thêm dòng trống để tách biệt

        for row in dataframe_to_rows(df_quy_trinh, index=False, header=True):
            ws_quy_trinh.append(row)

        for cell in ws_quy_trinh[4]:  # In đậm tiêu đề
            cell.font = cell.font.copy(bold=True)

        ws_luan_chuyen = wb.create_sheet(title="LuanChuyen")
        for row in dataframe_to_rows(df_luan_chuyen, index=False, header=True):
            ws_luan_chuyen.append(row)

        for cell in ws_luan_chuyen[1]:
            cell.font = cell.font.copy(bold=True)

        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            wb.save(file_path)
            messagebox.showinfo("Thông báo", f"File đã được lưu tại {file_path}")

    def attach_file(self):
        self.attached_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.attached_file_path:
            self.attached_file_label.config(text=self.attached_file_path)
            messagebox.showinfo("Thông báo", f"File đã được đính kèm: {self.attached_file_path}")

    def open_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if self.file_path:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, self.file_path)
    
    def download_sample_file(self):
        sample_file_path = os.path.join(sys._MEIPASS, "data.xlsx") if hasattr(sys, "_MEIPASS") else "data.xlsx"
        destination_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if destination_path:
            with open(sample_file_path, 'rb') as sample_file:
                with open(destination_path, 'wb') as destination_file:
                    destination_file.write(sample_file.read())
            messagebox.showinfo("Thông báo", "Tải file mẫu thành công!")
    
    def check_for_update(self):
        messagebox.showinfo("Thông báo", "Chức năng check update hiện chưa được hỗ trợ.")
    
    def start_thread(self):
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        messagebox.showinfo("Thông báo", "Chức năng start hiện chưa được hỗ trợ.")

    def stop_automation(self):
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.stop_flag = True
        messagebox.showinfo("Thông báo", "Chức năng stop hiện chưa được hỗ trợ.")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
