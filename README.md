# BTVN
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import csv
import datetime
import pandas as pd

employee_data = []

def save_to_csv():

    with open("employees.csv", "w", newline="", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow(["Mã", "Tên", "Đơn vị", "Chức danh", "Ngày sinh", "Giới tính", "Số CMND", "Nơi cấp", "Ngày cấp"])
        writer.writerows(employee_data)
    messagebox.showinfo("Thành công", "Dữ liệu đã được lưu vào file employees.csv!")

def add_employee():
    """Thêm nhân viên vào danh sách."""
    emp = [
        entry_ma.get(),
        entry_ten.get(),
        entry_don_vi.get(),
        entry_chuc_danh.get(),
        entry_ngay_sinh.get(),
        gender_var.get(),
        entry_cmnd.get(),
        entry_noi_cap.get(),
        entry_ngay_cap.get(),
    ]
    if all(emp):
        employee_data.append(emp)
        messagebox.showinfo("Thành công", "Nhân viên đã được thêm!")
    else:
        messagebox.showerror("Lỗi", "Vui lòng điền đầy đủ thông tin!")

def birthday_today():

    today = datetime.datetime.now().strftime("%d/%m")
    birthdays = [emp for emp in employee_data if emp[4][:5] == today]
    if birthdays:
        birthday_message = "\n".join([f"{emp[1]} ({emp[0]})" for emp in birthdays])
        messagebox.showinfo("Sinh nhật hôm nay", f"Nhân viên có sinh nhật hôm nay:\n{birthday_message}")
    else:
        messagebox.showinfo("Thông báo", "Không có nhân viên nào sinh nhật hôm nay.")

def export_to_excel():

    if not employee_data:
        messagebox.showerror("Lỗi", "Không có dữ liệu để xuất!")
        return
    df = pd.DataFrame(employee_data, columns=["Mã", "Tên", "Đơn vị", "Chức danh", "Ngày sinh", "Giới tính", "Số CMND", "Nơi cấp", "Ngày cấp"])
    df["Ngày sinh"] = pd.to_datetime(df["Ngày sinh"], format="%d/%m/%Y")
    df = df.sort_values(by="Ngày sinh", ascending=False)
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if save_path:
        df.to_excel(save_path, index=False)
        messagebox.showinfo("Thành công", f"Dữ liệu đã được xuất ra file {save_path}")


root = tk.Tk()
root.title("Quản lý thông tin nhân viên")

fields = [
    ("Mã", 20),
    ("Tên", 30),
    ("Đơn vị", 30),
    ("Chức danh", 30),
    ("Ngày sinh (DD/MM/YYYY)", 15),
    ("Số CMND", 20),
    ("Nơi cấp", 30),
    ("Ngày cấp (DD/MM/YYYY)", 15),
]

entries = {}
for idx, (label, width) in enumerate(fields):
    ttk.Label(root, text=label).grid(row=idx, column=0, padx=5, pady=5, sticky=tk.W)
    entry = ttk.Entry(root, width=width)
    entry.grid(row=idx, column=1, padx=5, pady=5)
    entries[label] = entry

entry_ma = entries["Mã"]
entry_ten = entries["Tên"]
entry_don_vi = entries["Đơn vị"]
entry_chuc_danh = entries["Chức danh"]
entry_ngay_sinh = entries["Ngày sinh (DD/MM/YYYY)"]
entry_cmnd = entries["Số CMND"]
entry_noi_cap = entries["Nơi cấp"]
entry_ngay_cap = entries["Ngày cấp (DD/MM/YYYY)"]

gender_var = tk.StringVar(value="Nam")
ttk.Label(root, text="Giới tính").grid(row=4, column=2, padx=5, pady=5, sticky=tk.W)
ttk.Radiobutton(root, text="Nam", variable=gender_var, value="Nam").grid(row=4, column=3, padx=5, pady=5)
ttk.Radiobutton(root, text="Nữ", variable=gender_var, value="Nữ").grid(row=4, column=4, padx=5, pady=5)

ttk.Button(root, text="Thêm nhân viên", command=add_employee).grid(row=10, column=0, padx=5, pady=10, sticky=tk.W)
ttk.Button(root, text="Sinh nhật hôm nay", command=birthday_today).grid(row=10, column=1, padx=5, pady=10, sticky=tk.W)
ttk.Button(root, text="Xuất toàn bộ danh sách", command=export_to_excel).grid(row=10, column=2, padx=5, pady=10, sticky=tk.W)
ttk.Button(root, text="Lưu dữ liệu", command=save_to_csv).grid(row=10, column=3, padx=5, pady=10, sticky=tk.W)

root.mainloop()
