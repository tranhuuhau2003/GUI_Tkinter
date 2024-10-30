import email
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
import os
import smtplib
import sqlite3
from email import encoders
import pandas as pd
import tkinter as tk
from tkinter import *
from tkinter import filedialog, messagebox, Toplevel, ttk
from tkinter import Frame, Tk
from PIL import Image, ImageTk
from tkinter import Label, Entry, Button,  Radiobutton, IntVar
import matplotlib.pyplot as plt
from matplotlib import rcParams
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import schedule
import time
import threading
import imaplib
from email.header import decode_header
from datetime import datetime
from datetime import timedelta
import re
from collections import defaultdict
from tkinter import scrolledtext
import random

# Biến toàn cục để lưu tên file tóm tắt
summary_file = 'TongHopSinhVienVangCacLop.xlsx'

chart_frame = None

# Đăng ký adapter datetime
sqlite3.register_adapter(datetime, lambda d: d.timestamp())
sqlite3.register_converter("timestamp", lambda t: datetime.fromtimestamp(t))

# Biến toàn cục lưu trữ dữ liệu sinh viên
global df_sinh_vien, ma_lop, ten_mon_hoc
df_sinh_vien, ma_lop, ten_mon_hoc = None, None, None

def load_data():
    try:
        Tk().withdraw()  # Ẩn cửa sổ chính
        file_path = filedialog.askopenfilename(title="Chọn file Excel", filetypes=[("Excel files", "*.xlsx;*.xls")])
        
        if not file_path:
            print("Không có file nào được chọn.")
            return None, None, None, None, None

        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Không tìm thấy file tại: {file_path}")

        # Đọc file Excel
        df = pd.read_excel(file_path, header=None)
        df = df.fillna('')

        # Lấy thông tin cần thiết
        dot = df.iloc[5, 2]
        ma_lop = df.iloc[7, 2]
        ten_mon_hoc = df.iloc[8, 2]

        # Lấy dữ liệu sinh viên
        df_sinh_vien = df.iloc[13:, [1, 2, 3, 4, 5, 6, 9, 12, 15, 18, 21, 24, 25, 26, 27]]
        df_sinh_vien.columns = ['MSSV', 'Họ đệm', 'Tên', 'Giới tính', 'Ngày sinh', '11/06/2024', '18/06/2024', '25/06/2024', '02/07/2024', '09/07/2024', '23/07/2024', 'Vắng có phép', 'Vắng không phép', 'Tổng số tiết', '(%) vắng']

        # Chuyển đổi các cột phần trăm vắng từ ',' sang '.'
        if '(%) vắng' in df_sinh_vien.columns:
            df_sinh_vien['(%) vắng'] = df_sinh_vien['(%) vắng'].apply(lambda x: str(x).replace(',', '.') if isinstance(x, str) else x)

        # Chuyển đổi cột vắng có phép và không phép
        df_sinh_vien['Vắng có phép'] = pd.to_numeric(df_sinh_vien['Vắng có phép'], errors='coerce').fillna(0)
        df_sinh_vien['Vắng không phép'] = pd.to_numeric(df_sinh_vien['Vắng không phép'], errors='coerce').fillna(0)

        # Tính tổng buổi vắng
        df_sinh_vien['Tổng buổi vắng'] = df_sinh_vien['Vắng có phép'] + df_sinh_vien['Vắng không phép']
        
        # Lấy danh sách MSSV
        mssv_list = df_sinh_vien['MSSV'].tolist()        
        
        return df_sinh_vien, dot, ma_lop, ten_mon_hoc, mssv_list
    except Exception as e:
        print(f"Lỗi khi đọc dữ liệu từ Excel: {e}")
        return None, None, None, None, None

    
def add_data_to_sqlite(df_sinh_vien, dot, ma_lop, ten_mon_hoc, mssv_list):
    try:
        conn = sqlite3.connect('students.db')
        cursor = conn.cursor()

        # Tạo bảng mới với các cột chính xác
        cursor.execute("""CREATE TABLE IF NOT EXISTS students (
                            id INTEGER PRIMARY KEY AUTOINCREMENT,
                            mssv TEXT,
                            ho_dem TEXT,
                            ten TEXT,
                            gioi_tinh TEXT,
                            ngay_sinh TEXT,
                            "11/06/2024" TEXT,
                            "18/06/2024" TEXT,
                            "25/06/2024" TEXT,
                            "02/07/2024" TEXT,
                            "09/07/2024" TEXT,
                            "23/07/2024" TEXT,
                            vang_co_phep INTEGER,
                            vang_khong_phep INTEGER,
                            tong_so_tiet INTEGER,
                            ty_le_vang REAL,
                            tong_buoi_vang INTEGER,
                            dot TEXT,
                            ma_lop TEXT,
                            ten_mon_hoc TEXT,
                            email_student TEXT
                        )""")

        # Tạo bảng parents
        cursor.execute('''CREATE TABLE IF NOT EXISTS parents (
                            mssv TEXT PRIMARY KEY,
                            email_ph TEXT  -- Email của phụ huynh
                        )''')

        # Tạo bảng teachers
        cursor.execute('''CREATE TABLE IF NOT EXISTS teachers (
                            mssv TEXT PRIMARY KEY,
                            email_gvcn TEXT  -- Email của giáo viên chủ nhiệm
                        )''')
        
        # Tạo bảng TBM
        cursor.execute('''CREATE TABLE IF NOT EXISTS tbm (
                            mssv TEXT PRIMARY KEY,
                            email_tbm TEXT  -- Email của trưởng bộ môn
                        )''')
        
        conn.commit()

        # Thêm dữ liệu mới vào bảng students
        for _, row in df_sinh_vien.iterrows():
            try:
                # Kiểm tra xem cặp MSSV và Mã lớp đã tồn tại chưa
                cursor.execute("SELECT COUNT(*) FROM students WHERE mssv = ? AND ma_lop = ?", (row['MSSV'], ma_lop))
                exists = cursor.fetchone()[0] > 0
                
                # Nếu không tồn tại thì thêm vào
                if not exists:
                    values_to_insert = (
                        str(row['MSSV']),
                        str(row['Họ đệm']),
                        str(row['Tên']),
                        str(row['Giới tính']),
                        str(row['Ngày sinh']),
                        str(row['11/06/2024']),
                        str(row['18/06/2024']),
                        str(row['25/06/2024']),
                        str(row['02/07/2024']),
                        str(row['09/07/2024']),
                        str(row['23/07/2024']),
                        int(float(row['Vắng có phép'])),
                        int(float(row['Vắng không phép'])),
                        int(float(row['Tổng số tiết'])),
                        float(row['(%) vắng']),
                        int(row['Tổng buổi vắng']),
                        dot,
                        ma_lop,
                        ten_mon_hoc
                    )

                    cursor.execute("""INSERT INTO students (
                                        mssv, ho_dem, ten, gioi_tinh, ngay_sinh, 
                                        "11/06/2024", "18/06/2024", "25/06/2024", "02/07/2024", "09/07/2024", "23/07/2024",
                                        vang_co_phep, vang_khong_phep, tong_so_tiet, ty_le_vang, tong_buoi_vang,
                                        dot, ma_lop, ten_mon_hoc) 
                                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""", values_to_insert)
                else:
                    print(f"MSSV: {row['MSSV']} với mã lớp: {ma_lop} đã tồn tại. Không thêm vào DB.")

            except Exception as e:
                print(f"Lỗi khi thêm sinh viên {row['MSSV']}: {e}")
        
        # Thêm email vào bảng students
        for mssv in mssv_list:
            email_student = f"tranhuuhauthh@gmail.com"  # Tạo email mẫu cho sinh viên
            cursor.execute('UPDATE students SET email_student = ? WHERE mssv = ?', (email_student, mssv))
            
        # Xóa dữ liệu cũ trước khi thêm dữ liệu mới vào bảng parents
        cursor.execute("DELETE FROM parents")
        for mssv in mssv_list:
            email_ph = f"tranhuuhauthh@gmail.com"  # Tạo email mẫu cho phụ huynh
            cursor.execute('INSERT OR IGNORE INTO parents (mssv, email_ph) VALUES (?, ?)', (mssv, email_ph))

        # Xóa dữ liệu cũ trước khi thêm dữ liệu mới vào bảng teachers
        cursor.execute("DELETE FROM teachers")
        for mssv in mssv_list:
            email_gvcn = f"tranhuuhauthh@gmail.com"  # Tạo email mẫu cho giáo viên chủ nhiệm
            cursor.execute('INSERT OR IGNORE INTO teachers (mssv, email_gvcn) VALUES (?, ?)', (mssv, email_gvcn))
            
        # Xóa dữ liệu cũ trước khi thêm dữ liệu mới vào bảng tbm
        cursor.execute("DELETE FROM tbm")
        for mssv in mssv_list:
            email_tbm = f"tranhuuhauthh@gmail.com"  # Tạo email mẫu cho trưởng bộ môn
            cursor.execute('INSERT OR IGNORE INTO tbm (mssv, email_tbm) VALUES (?, ?)', (mssv, email_tbm))

        conn.commit()   
        conn.close()
    except Exception as e:
        print(f"Lỗi khi thêm dữ liệu vào SQLite: {e}")

def load_from_excel_to_treeview(tree):
    df_sinh_vien, dot, ma_lop, ten_mon_hoc, mssv_list = load_data()

    if df_sinh_vien is not None:
        add_data_to_sqlite(df_sinh_vien, dot, ma_lop, ten_mon_hoc, mssv_list)  # Thêm dữ liệu vào SQLite

        # Xóa dữ liệu hiện tại trong Treeview
        for row in tree.get_children():
            tree.delete(row)

        # Kết nối đến cơ sở dữ liệu để lấy dữ liệu
        conn = sqlite3.connect('students.db')
        cursor = conn.cursor()

        # Lấy dữ liệu từ bảng students
        cursor.execute("SELECT * FROM students")
        rows = cursor.fetchall()
        conn.close()

        # Hiển thị dữ liệu đã tải vào Treeview với cột STT
        for index, row in enumerate(rows):
            stt = index + 1  # Số thứ tự

            # Chỉ lấy các cột cần thiết
            data_to_insert = [
                row[1],  # MSSV
                row[2],  # Họ đệm
                row[3],  # Tên
                row[4],  # Giới tính
                row[5],  # Ngày sinh
                row[12],  # Vắng có phép
                row[13],  # Vắng không phép
                row[14],  # Tổng số tiết
                row[15],  # (%) vắng
                row[16],  # Tổng buổi vắng
                row[17],  # Đợt
                row[18],  # Mã lớp
                row[19]   # Tên môn học
            ]

            # Thêm dữ liệu vào Treeview, bao gồm STT
            tree.insert('', 'end', values=[stt] + data_to_insert)  # Thêm dữ liệu vào Treeview
            
    update_button_states()
    
      
def clear_table(tree):
    # Kết nối đến cơ sở dữ liệu
    conn = sqlite3.connect('students.db')  
    cursor = conn.cursor()
    
    try:
        # Xóa dữ liệu trong các bảng
        rows_deleted = 0
        cursor.execute("DELETE FROM students")
        rows_deleted += cursor.rowcount
        
        cursor.execute("DELETE FROM parents")
        rows_deleted += cursor.rowcount
        
        cursor.execute("DELETE FROM teachers")
        rows_deleted += cursor.rowcount

        # Xác nhận thay đổi
        conn.commit()
        print(f"Dữ liệu đã được xóa thành công từ các bảng. Số lượng dòng đã xóa: {rows_deleted}")
       
    except sqlite3.Error as e:
        print(f"Đã xảy ra lỗi khi xóa dữ liệu: {e}")
    finally:
        # Đóng kết nối
        cursor.close()
        conn.close()
    
    refresh_treeview(tree)
    
    update_button_states()
    
def refresh_treeview(tree):
    # Xóa dữ liệu hiện tại trong treeview
    for item in tree.get_children():
        tree.delete(item)

    # Kết nối đến SQLite và lấy dữ liệu
    conn = sqlite3.connect('students.db')
    cursor = conn.cursor()
     # Chỉ lấy các cột cần thiết
    cursor.execute("""
        SELECT MSSV, ho_dem, ten, gioi_tinh, ngay_sinh, 
               vang_co_phep, vang_khong_phep, tong_so_tiet, 
               ty_le_vang, tong_buoi_vang, dot, ma_lop, ten_mon_hoc 
        FROM students
    """)
    rows = cursor.fetchall()
    
    for index, row in enumerate(rows):
        # Chèn dữ liệu vào TreeView với cột STT
        stt = index + 1  # Tính STT, bắt đầu từ 1
        tree.insert('', 'end', values=(stt,  # STT
            row[0],  # MSSV
            row[1],  # Họ đệm
            row[2],  # Tên
            row[3],  # Giới tính
            row[4],  # Ngày sinh
            row[5],  # Vắng có phép
            row[6],  # Vắng không phép
            row[7],  # Tổng số tiết
            row[8],  # (%) vắng
            row[9],  # Tổng buổi vắng
            row[10],  # Đợt
            row[11],  # Mã lớp
            row[12]   # Tên môn học
        ))
    
    conn.close()

def add_student(tree):
    # Tạo một cửa sổ mới để thêm sinh viên
    window = Toplevel()
    window.title("Thêm Sinh Viên")
    window.geometry("250x450")
    window.configure(bg="#F2D0D3")  # Thiết lập màu nền cho cửa sổ

    labels = ["MSSV", "Họ đệm", "Tên", "Giới tính", "Ngày sinh", 
              "Vắng có phép", "Vắng không phép", "Tổng số tiết", 
              'Đợt', 'Mã lớp', 'Tên môn học']
    entries = []

    font_style = ("Times New Roman", 9)

    for i, label in enumerate(labels):
        Label(window, text=label, font=font_style, bg="#F2D0D3").grid(row=i, column=0, padx=10, pady=5, sticky='w')
        
        if label == "Giới tính":
            gender_var = IntVar()
            Radiobutton(window, text="Nam", variable=gender_var, value=1, font=font_style, bg="#F2D0D3").grid(row=i, column=1, padx=10, pady=5, sticky='w')
            Radiobutton(window, text="Nữ", variable=gender_var, value=2, font=font_style, bg="#F2D0D3").grid(row=i, column=1, columnspan=2, padx=10, pady=5, sticky='e')
            entries.append(gender_var)
        else:
            entry = Entry(window, font=font_style)
            entry.grid(row=i, column=1, padx=10, pady=5)
            entries.append(entry)

    def save_student():
        conn = sqlite3.connect('students.db')
        cursor = conn.cursor()
        try:
            mssv = entries[0].get()
            ho_dem = entries[1].get()
            ten = entries[2].get()
            gioi_tinh = "Nam" if entries[3].get() == 1 else "Nữ"
            ngay_sinh = entries[4].get()
            vang_co_phep = int(entries[5].get())
            vang_khong_phep = int(entries[6].get())
            tong_so_tiet = int(entries[7].get())
            tong_buoi_vang = vang_co_phep + vang_khong_phep
            ty_le_vang = round((tong_buoi_vang / tong_so_tiet) * 100, 1) if tong_so_tiet > 0 else 0

            dot = entries[8].get()
            ma_lop = entries[9].get()
            ten_mon_hoc = entries[10].get()

            values_to_insert = (
                mssv, ho_dem, ten, gioi_tinh, ngay_sinh, 
                vang_co_phep, vang_khong_phep, tong_so_tiet, 
                ty_le_vang, tong_buoi_vang, dot, ma_lop, ten_mon_hoc
            )
            
            cursor.execute("""INSERT INTO students 
                              (mssv, ho_dem, ten, gioi_tinh, ngay_sinh, 
                               vang_co_phep, vang_khong_phep, tong_so_tiet, 
                               ty_le_vang, tong_buoi_vang, dot, ma_lop, ten_mon_hoc) 
                              VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""", 
                           values_to_insert)
            conn.commit()
            messagebox.showinfo("Thành công", "Đã thêm sinh viên thành công.")
            refresh_treeview(tree)
            window.destroy()
        except sqlite3.IntegrityError:
            messagebox.showerror("Lỗi", "MSSV đã tồn tại.")
        except ValueError:
            messagebox.showerror("Lỗi", "Vui lòng nhập số hợp lệ cho các trường vắng.")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Có lỗi xảy ra: {e}")
        finally:
            conn.close()

    # Nút Lưu với kiểu đẹp, màu nền và viền tròn
    save_button = Button(
        window, 
        text="Lưu", 
        command=save_student, 
        font=("Times New Roman", 9, "bold"),  # Giảm kích thước font
        bg="#F2A2C0",  # Màu nền cho nút
        fg="black",  # Màu chữ trắng
        padx=10, pady=5,  # Giảm độ đệm để nút nhỏ hơn
        relief=RAISED, 
        bd=2  # Độ dày viền
    )
    save_button.grid(row=len(labels), column=0, columnspan=2, pady=15)
    save_button.configure(highlightbackground="#F2A2C0", highlightthickness=2) 
    
    update_button_states()


def edit_student(tree):
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("Chọn sinh viên", "Vui lòng chọn sinh viên để chỉnh sửa.")
        return

    # Tạo một cửa sổ mới để chỉnh sửa sinh viên
    window = Toplevel()
    window.title("Chỉnh Sửa Sinh Viên")
    
    # Đặt kích thước cho cửa sổ
    window.geometry("270x400")  # Tăng kích thước để vừa với các ô nhập

    # Đặt màu nền cho cửa sổ
    window.configure(bg="#F2D0D3")

    # Lấy dữ liệu của sinh viên đã chọn
    item_values = tree.item(selected_item, "values")
    
    # Danh sách các nhãn cho trường cần hiển thị
    labels = ["MSSV", "Họ đệm", "Tên", "Giới tính", "Ngày sinh", 
              "Vắng có phép", "Vắng không phép", "Tổng số tiết", 
              "Đợt", "Mã lớp", "Tên môn học"]
    
    # Chỉ lấy các giá trị cần thiết từ item_values
    values_to_display = item_values[1:9] + item_values[11:14]  
    
    entries = []
    gender_var = IntVar()  # Biến để lưu giá trị giới tính

    # Thay đổi kiểu chữ cho các nhãn và ô nhập
    font_style = ("Times New Roman", 9) 

    # Hiển thị thông tin hiện có của sinh viên vào các trường thông tin để chỉnh sửa
    for i, (label, value) in enumerate(zip(labels, values_to_display)):  
        Label(window, text=label, font=font_style, bg="#F2D0D3").grid(row=i, column=0, padx=10, pady=5, sticky='w')  # Căn trái
        
        if label == "Giới tính":
            # Đặt giá trị cho radio button dựa vào dữ liệu
            gioi_tinh = item_values[4].strip()
            if gioi_tinh == "Nam":
                gender_var.set(1)  # Đặt giá trị 1 cho Nam
            elif gioi_tinh == "Nữ":
                gender_var.set(2)  # Đặt giá trị 2 cho Nữ
            
            # Thêm nút radio cho giới tính
            Radiobutton(window, text="Nam", variable=gender_var, value=1, font=font_style, bg="#F2D0D3").grid(row=i, column=1, padx=10, pady=5, sticky='w')
            Radiobutton(window, text="Nữ", variable=gender_var, value=2, font=font_style, bg="#F2D0D3").grid(row=i, column=1, columnspan=2, padx=10, pady=5, sticky='e')
        elif label in ["MSSV", "Đợt", "Mã lớp", "Tên môn học"]:
            label_value = Label(window, text=value, font=font_style, bg="#F2D0D3")  # Hiển thị dưới dạng Label
            label_value.grid(row=i, column=1, padx=10, pady=5)  # Đặt bên cạnh nhãn
        else:
            entry = Entry(window, font=font_style)  # Đặt kiểu chữ cho ô nhập
            entry.insert(0, value)  # Điền giá trị hiện tại vào ô nhập
            entry.grid(row=i, column=1, padx=10, pady=5)  # Đặt ô nhập bên cạnh nhãn
            entries.append(entry)

    def update_student():
        conn = sqlite3.connect('students.db')
        cursor = conn.cursor()

        # Lấy giá trị từ các trường đã nhập
        mssv_cu = item_values[1]  # MSSV cũ từ item_values
        ho_dem = entries[0].get()
        gioi_tinh = "Nam" if gender_var.get() == 1 else "Nữ"  # Lấy giá trị giới tính từ radio button
        ten = entries[1].get()
        ngay_sinh = entries[2].get()
        vang_co_phep = int(entries[3].get())
        vang_khong_phep = int(entries[4].get())
        tong_so_tiet = int(entries[5].get())
        dot = item_values[11]  # Đợt cũ
        ma_lop = item_values[12]  # Mã lớp cũ
        ten_mon_hoc = item_values[13]  # Tên môn cũ

        # Tính tổng buổi vắng
        tong_buoi_vang = vang_co_phep + vang_khong_phep
        # Tính % vắng
        if tong_so_tiet > 0:
            ty_le_vang = round((tong_buoi_vang / tong_so_tiet) * 100, 1)  # Làm tròn tới 1 chữ số thập phân
        else:
            ty_le_vang = 0

        # Cập nhật thông tin sinh viên
        try:
            cursor.execute("""UPDATE students SET 
                              ho_dem = ?, ten = ?, gioi_tinh = ?, ngay_sinh = ?, 
                              vang_co_phep = ?, vang_khong_phep = ?, tong_so_tiet = ?, 
                              ty_le_vang = ?, tong_buoi_vang = ?, dot = ?, ma_lop = ?, ten_mon_hoc = ?
                              WHERE mssv = ?""", 
                           (ho_dem, ten, gioi_tinh, ngay_sinh, vang_co_phep, vang_khong_phep, 
                            tong_so_tiet, ty_le_vang, tong_buoi_vang, dot, ma_lop, ten_mon_hoc, mssv_cu))
            conn.commit()
            messagebox.showinfo("Thành công", "Đã cập nhật sinh viên thành công.")
            refresh_treeview(tree)  # Cập nhật Treeview
            window.destroy()
        except Exception as e:
            messagebox.showerror("Lỗi", f"Có lỗi xảy ra: {e}")
        finally:
            conn.close()

    Button(window, text="Cập nhật", command=update_student, font=font_style, bg="#F2A2C0", fg="black").grid(row=len(values_to_display), column=0, columnspan=2, pady=20)

def delete_student(tree):
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("Chọn sinh viên", "Vui lòng chọn sinh viên để xóa.")
        return

    item_values = tree.item(selected_item, "values")
    mssv = item_values[1]  # Lấy MSSV của sinh viên được chọn

    # Hiển thị hộp thoại xác nhận
    confirm = messagebox.askyesno("Xác nhận xóa", f"Bạn có chắc chắn muốn xóa sinh viên có MSSV: {mssv}?")
    if not confirm:
        return  # Nếu người dùng không xác nhận, thoát khỏi hàm

    conn = sqlite3.connect('students.db')
    cursor = conn.cursor()
    cursor.execute("DELETE FROM students WHERE mssv=?", (mssv,))
    conn.commit()
    conn.close()

    refresh_treeview(tree)  # Cập nhật Treeview
    messagebox.showinfo("Thành công", "Đã xóa sinh viên thành công.")
    update_button_states()


def view_details(tree):
    conn = sqlite3.connect('students.db')
    cursor = conn.cursor()
    selected_item = tree.selection()
    
    if selected_item:
        item_data = tree.item(selected_item, 'values')
        mssv = item_data[1]
        
        query = '''
            SELECT ho_dem, ten, gioi_tinh, ngay_sinh, dot, ma_lop, ten_mon_hoc,
                "vang_co_phep", "vang_khong_phep", "tong_so_tiet", ty_le_vang, tong_buoi_vang,
                "11/06/2024", "18/06/2024", "25/06/2024", 
                "02/07/2024", "09/07/2024", "23/07/2024"
            FROM students
            WHERE mssv = ?
        '''
        cursor.execute(query, (mssv,))
        details_data = cursor.fetchone()

        if details_data:
            time_off = []
            date_columns = ["11/06/2024", "18/06/2024", "25/06/2024", 
                            "02/07/2024", "09/07/2024", "23/07/2024"]
            for i, date in enumerate(date_columns, start=12):
                if details_data[i] in ["K", "P"]:
                    time_off.append(date)
            
            detail_window = tk.Toplevel()
            detail_window.title("Chi tiết thông tin sinh viên")
            detail_window.geometry("650x500+450+150")
            detail_window.configure(bg="#F2D0D3")
            
            labels = [
                ("MSSV:", mssv),
                ("Họ tên:", f"{details_data[0]} {details_data[1]}"),
                ("Giới tính:", details_data[2]),
                ("Ngày sinh:", details_data[3]),
                ("Đợt:", details_data[4]),
                ("Mã lớp:", details_data[5]),
                ("Tên môn học:", details_data[6]),
                ("Số tiết nghỉ có phép:", details_data[7]),
                ("Số tiết nghỉ không phép:", details_data[8]),
                ("Tổng số tiết:", details_data[9]),
                ("Tỷ lệ vắng:", f"{details_data[10]}%"),
                ("Tổng buổi vắng:", details_data[11]),
                ("Thời gian nghỉ:", ', '.join(time_off) if time_off else "Không có")
            ]
            
            for label_text, data in labels:
                frame = tk.Frame(detail_window, bg="#F2D0D3")
                frame.pack(anchor="w", padx=20, pady=2)
                
                label = tk.Label(frame, text=label_text, font=("Times New Roman", 11, "bold"), bg="#F2D0D3")
                label.pack(side="left")
                
                value = tk.Label(frame, text=data, font=("Times New Roman", 11), bg="#F2D0D3")
                value.pack(side="left")
            
            # Thiết kế nút đóng với màu sắc, bo tròn và hiệu ứng di chuột
            def on_enter(e):
                close_button['background'] = "#D98880"  # Thay đổi màu khi di chuột
                close_button['foreground'] = "white"

            def on_leave(e):
                close_button['background'] = "#F2A2C0"  # Trở về màu gốc khi rời chuột
                close_button['foreground'] = "black"

            close_button = tk.Button(
                detail_window, 
                text="Đóng", 
                font=("Times New Roman", 11, "bold"), 
                command=detail_window.destroy,
                bg="#F2A2C0",      # Màu nền nút
                fg="black",        # Màu chữ
                activebackground="#CD6155",  # Màu khi nhấn
                activeforeground="white",
                relief="flat",     # Loại bỏ viền nổi
                padx=10, pady=5,   # Tăng kích thước nút
                borderwidth=2      # Độ dày viền
            )
            close_button.pack(pady=20)

            # Gắn hiệu ứng di chuột vào nút
            close_button.bind("<Enter>", on_enter)
            close_button.bind("<Leave>", on_leave)
        
        else:
            messagebox.showerror("Lỗi", "Không tìm thấy thông tin sinh viên.")
    else:
        messagebox.showwarning("Cảnh báo", "Vui lòng chọn một sinh viên để xem chi tiết.")
    
    conn.close()

def sort_students_by_absences(tree):
    # Xóa dữ liệu hiện tại trong treeview
    for item in tree.get_children():
        tree.delete(item)

    # Kết nối đến SQLite và lấy dữ liệu
    conn = sqlite3.connect('students.db')
    cursor = conn.cursor()

    # Sắp xếp theo tổng buổi vắng giảm dần
     # Chỉ lấy các cột cần thiết
    cursor.execute("""
        SELECT MSSV, ho_dem, ten, gioi_tinh, ngay_sinh, 
               vang_co_phep, vang_khong_phep, tong_so_tiet, 
               ty_le_vang, tong_buoi_vang, dot, ma_lop, ten_mon_hoc 
        FROM students
        ORDER BY tong_buoi_vang DESC
    """)
    rows = cursor.fetchall()

    # Biến đếm để đánh số thứ tự
    stt = 1

    for row in rows:
        # Chèn dữ liệu vào TreeView với STT
        item_id = tree.insert('', 'end', values=(
            stt,  # STT - cột đầu tiên
            row[0],  # MSSV
            row[1],  # Họ đệm
            row[2],  # Tên
            row[3],  # Giới tính
            row[4],  # Ngày sinh
            row[5],  # Vắng có phép
            row[6],  # Vắng không phép
            row[7],  # Tổng số tiết
            row[8],  # (%) vắng
            row[9],  # Tổng buổi vắng
            row[10], # Đợt
            row[11], # Mã lớp
            row[12]  # Tên môn học
        ))

        # Kiểm tra tỷ lệ vắng và bôi đỏ nếu >= 50.0
        if row[8] >= 50.0:  # Giả sử cột 8 là tỷ lệ vắng
            tree.item(item_id, tags=('highlight',))

        # Tăng số thứ tự cho lần lặp tiếp theo
        stt += 1
    # Định nghĩa kiểu bôi đỏ cho tag
    tree.tag_configure('highlight', foreground='red')

    conn.close()

def search_students(tree, search_by, search_value):
    # Xóa dữ liệu hiện tại trong treeview
    for item in tree.get_children():
        tree.delete(item)

    # Kết nối tới CSDL SQLite
    conn = sqlite3.connect('students.db')
    cursor = conn.cursor()

    # Câu truy vấn dựa trên tiêu chí tìm kiếm
    query = "SELECT mssv, ho_dem, ten, gioi_tinh, ngay_sinh, vang_co_phep, vang_khong_phep, tong_so_tiet, ty_le_vang, tong_buoi_vang, dot, ma_lop, ten_mon_hoc FROM students WHERE "

    if search_by == "MSSV":
        query += "mssv LIKE ?"
        search_value = '%' + search_value.strip() + '%'
    elif search_by == "Tên":
        query += "(ho_dem || ' ' || ten) LIKE ?"
        search_value = '%' + search_value.strip() + '%'
    elif search_by == "Tỷ lệ vắng":
        try:
            search_value = float(search_value)  # Chuyển thành float để so sánh số liệu
            query += "ty_le_vang >= ?"
        except ValueError:
            print("Giá trị tìm kiếm tỷ lệ vắng phải là số.")
            conn.close()
            return
    else:
        conn.close()
        return

    # Thực thi câu truy vấn
    cursor.execute(query, (search_value,))
    rows = cursor.fetchall()

    # Biến đếm để đánh số thứ tự
    stt = 1
    for row in rows:
        # Chèn dữ liệu vào TreeView với STT
        tree.insert('', 'end', values=(
            stt,  # STT
            row[0],  # MSSV
            row[1],  # Họ đệm
            row[2],  # Tên
            row[3],  # Giới tính
            row[4],  # Ngày sinh
            row[5],  # Vắng có phép
            row[6],  # Vắng không phép
            row[7],  # Tổng số tiết
            row[8],  # (%) vắng
            row[9],  # Tổng buổi vắng
            row[10], # Đợt
            row[11], # Mã lớp
            row[12]  # Tên môn học
        ))

        stt += 1

    conn.close()

# Thêm giao diện tìm kiếm vào hệ thống chính
def add_search_interface(center_frame, tree):
    search_frame = Frame(center_frame, bg="#F2A2C0", bd=1)  # Giảm chiều cao bằng cách giảm bd
    search_frame.pack(side='top', fill='x', padx=5, pady=3)

    search_by_var = StringVar(value="MSSV")
    search_by_menu = OptionMenu(search_frame, search_by_var, "MSSV", "Tên", "Tỷ lệ vắng")
    search_by_menu.config(bg="#F2A2C0")
    search_by_menu.pack(side='left', padx=2)

    # Entry tìm kiếm
    search_entry = Entry(search_frame, font=("Times New Roman", 12), bd=3, width=17)
    search_entry.pack(side='left', padx=2)

    # Nút tìm kiếm
    search_button = Button(search_frame, text="Tìm", command=lambda: search_students(tree, search_by_var.get(), search_entry.get()), bg="#F2A2C0", font=("Times New Roman", 10))
    search_button.pack(side='left', padx=2)

# Hàm khởi tạo cơ sở dữ liệu tonghopsv
def initialize_database():
    conn = sqlite3.connect('tonghopsv.db', detect_types=sqlite3.PARSE_DECLTYPES)
    cursor = conn.cursor()

    # Tạo bảng nếu chưa tồn tại
    cursor.execute("""CREATE TABLE IF NOT EXISTS tonghopsv (
                        mssv TEXT PRIMARY KEY,
                        ho_dem TEXT,
                        ten TEXT,
                        gioi_tinh TEXT,
                        ngay_sinh TIMESTAMP,
                        vang_co_phep INTEGER,
                        vang_khong_phep INTEGER,
                        tong_so_tiet INTEGER,
                        ty_le_vang REAL,
                        tong_buoi_vang INTEGER,
                        dot TEXT,
                        ma_lop TEXT,
                        ten_mon_hoc TEXT
                    )""")
    
    # Xóa dữ liệu trong bảng khi khởi động
    cursor.execute("DELETE FROM tonghopsv")
    conn.commit()
    conn.close()
    
# Hàm lưu sinh viên vào SQLite dànhcho tonghopsv
def save_students_to_sqlite(df):
    print("Đang lưu sinh viên vào SQLite...")
    conn = sqlite3.connect('tonghopsv.db', detect_types=sqlite3.PARSE_DECLTYPES)
    cursor = conn.cursor()

    for _, row in df.iterrows():
        try:
            # Kiểm tra nếu 'ngay_sinh' là datetime, chuyển đổi thành chuỗi
            ngay_sinh_value = row['Ngày sinh']
            if isinstance(ngay_sinh_value, pd.Timestamp):  # Nếu là kiểu pandas Timestamp
                ngay_sinh_value = ngay_sinh_value.strftime('%Y-%m-%d')  # Chuyển đổi thành chuỗi

            cursor.execute("""INSERT OR IGNORE INTO tonghopsv (
                                mssv, ho_dem, ten, gioi_tinh, ngay_sinh, 
                                vang_co_phep, vang_khong_phep, tong_so_tiet, 
                                ty_le_vang, tong_buoi_vang, dot, ma_lop, ten_mon_hoc
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                           (row['MSSV'], row['Họ đệm'], row['Tên'], row['Giới tính'],
                            ngay_sinh_value,  # Dùng giá trị đã chuyển đổi
                            row['Vắng có phép'], row['Vắng không phép'],
                            row['Tổng số tiết'], row['(%) vắng'], row['Tổng buổi vắng'],
                            row['Đợt'], row['Mã lớp'], row['Tên môn học']))
        except Exception as e:
            print(f"Lỗi khi thêm sinh viên {row['MSSV']}: {e}")

    conn.commit()
    conn.close()

def send_email_with_ssl(summary_file):
    sender_email = "carotneee4@gmail.com"
    app_password = "bgjx tavb oxba ickr"  
    receiver_email = "vokhanhlinh04112k3@gmail.com"
    subject = "Tổng hợp sinh viên vắng nhiều"
    body = "Đính kèm là danh sách sinh viên vắng >= 50%."

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(sender_email, app_password)
            server.sendmail(sender_email, receiver_email, msg.as_string())
        print("Email đã được gửi thành công.")
    except Exception as e:
        print(f"Lỗi khi gửi email: {e}")
        
def send_email(to_address, subject, message):
    """Send email to the recipient."""
    from_address = "carotneee4@gmail.com"
    password = "bgjx tavb oxba ickr"  
    
    # Initialize email
    msg = MIMEMultipart()
    msg['From'] = from_address
    msg['To'] = to_address
    msg['Subject'] = subject

    # Email content
    msg.attach(MIMEText(message, 'plain'))

    # Configure SMTP server to send email
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(from_address, password)
        text = msg.as_string()
        server.sendmail(from_address, to_address, text)
        server.quit()
        print(f"Email sent to {to_address}")
    except Exception as e:
        print(f"Failed to send email to {to_address}: {e}")

def send_warning_emails():
    """Check and send warning emails for students."""
    # Connect to SQLite database
    connection = sqlite3.connect('students.db')
    cursor = connection.cursor()

    try:
        query = """
        SELECT mssv, ho_dem, ten, ma_lop, vang_co_phep, vang_khong_phep, tong_so_tiet, ty_le_vang
        FROM students
        """
        cursor.execute(query)
        records = cursor.fetchall()

        # Biến lưu trữ địa chỉ email đã gửi
        sent_emails = set()

        for row in records:
            mssv, ho_dem, ten, ma_lop, vang_co_phep, vang_khong_phep, tong_so_tiet, ty_le_vang = row
            
            # Get student and related emails
            student_email = get_student_email(cursor, mssv)
            parent_email = get_parent_email(cursor, mssv)
            homeroom_teacher_email = get_teacher_email(cursor, mssv)
            tbm_email = get_tbm_email(cursor, mssv)

            # Check and send warnings based on absence rate
            if ty_le_vang >= 50:
                subject = "Cảnh báo học vụ: Vắng học quá 50%"
                message = (f"Sinh viên {ho_dem} {ten} (Mã lớp: {ma_lop}) đã vắng hơn 50% số buổi học.")
                send_email(student_email, subject, message)
                send_email(parent_email, subject, message)
                send_email(homeroom_teacher_email, subject, message)
                send_email(tbm_email, subject, message)

                # Thêm vào tập hợp địa chỉ email đã gửi
                sent_emails.update([student_email, parent_email, homeroom_teacher_email, tbm_email])

            elif ty_le_vang >= 20:
                subject = "Cảnh báo học vụ: Vắng học quá 20%"
                message = f"Sinh viên {ho_dem} {ten} (Mã lớp: {ma_lop}) đã vắng hơn 20% số buổi học."
                send_email(student_email, subject, message)

                # Thêm vào tập hợp địa chỉ email đã gửi
                sent_emails.add(student_email)

        # Thông báo chỉ một lần sau khi hoàn thành gửi email
        if sent_emails:
            email_list = ', '.join(sent_emails)  # Chuyển đổi tập hợp thành chuỗi
            messagebox.showinfo("Email Success", f"Email đã gửi thành công tới: {email_list}")

    except Exception as e:
        messagebox.showerror("Email Error", f"Có lỗi xảy ra khi gửi email: {e}")  # Thông báo lỗi
    finally:
        connection.close()

def get_student_email(cursor, mssv):
    """Retrieve student email from the database based on MSSV."""
    query = f"SELECT email_student FROM students WHERE mssv = ?"
    cursor.execute(query, (mssv,))
    result = cursor.fetchone()
    return result[0] if result else None

def get_parent_email(cursor, mssv):
    """Retrieve parent email from the database based on MSSV."""
    query = f"SELECT email_ph FROM parents WHERE mssv = ?"
    cursor.execute(query, (mssv,))
    result = cursor.fetchone()
    return result[0] if result else None

def get_teacher_email(cursor, mssv):
    """Retrieve homeroom teacher email from the database based on MSSV."""
    query = f"SELECT email_gvcn FROM teachers WHERE mssv = ?"
    cursor.execute(query, (mssv,))
    result = cursor.fetchone()
    return result[0] if result else None

def get_tbm_email(cursor, mssv):
    """Retrieve TBM email from the database based on MSSV."""
    query = f"SELECT email_tbm FROM tbm WHERE mssv = ?"
    cursor.execute(query, (mssv,))
    result = cursor.fetchone()
    return result[0] if result else None

def load_and_summarize_students(tree):
    global class_codes
    class_codes = []  # Khởi tạo danh sách để lưu mã lớp

    # Xóa tất cả mục trong giao diện (tree) ngay từ đầu
    for item in tree.get_children():
        tree.delete(item)

    excel_files = filedialog.askopenfilenames(title='Chọn các file Excel', filetypes=[("Excel files", "*.xlsx;*.xls")])
    
    # Kiểm tra xem có tệp Excel hợp lệ không
    if not excel_files or not all(os.path.exists(f) for f in excel_files):
        print("Không có tệp Excel hợp lệ được chọn!")
        return

    all_data = []

    for file in excel_files:
        df = pd.read_excel(file, header=None)
        df = df.fillna('')

        dot = df.iloc[5, 2]
        ma_lop = df.iloc[7, 2]
        ten_mon_hoc = df.iloc[8, 2]

        df_sinh_vien = df.iloc[13:, [1, 2, 3, 4, 5, 24, 25, 26, 27]]
        df_sinh_vien.columns = ['MSSV', 'Họ đệm', 'Tên', 'Giới tính', 'Ngày sinh', 'Vắng có phép', 'Vắng không phép', 'Tổng số tiết', '(%) vắng']

        df_sinh_vien['(%) vắng'] = df_sinh_vien['(%) vắng'].apply(lambda x: str(x).replace(',', '.') if isinstance(x, str) else x)
        df_sinh_vien['Vắng có phép'] = pd.to_numeric(df_sinh_vien['Vắng có phép'], errors='coerce').fillna(0)
        df_sinh_vien['Vắng không phép'] = pd.to_numeric(df_sinh_vien['Vắng không phép'], errors='coerce').fillna(0)
        df_sinh_vien['Tổng buổi vắng'] = df_sinh_vien['Vắng có phép'] + df_sinh_vien['Vắng không phép']

        df_sinh_vien['Đợt'] = dot
        df_sinh_vien['Mã lớp'] = ma_lop
        df_sinh_vien['Tên môn học'] = ten_mon_hoc

        # Xử lý cột Ngày sinh
        df_sinh_vien['Ngày sinh'] = pd.to_datetime(df_sinh_vien['Ngày sinh'], errors='coerce')

        all_data.append(df_sinh_vien)

        save_students_to_sqlite(df_sinh_vien)  # Gọi hàm lưu sinh viên vào SQLite

    combined_data = pd.concat(all_data, ignore_index=True)
    combined_data['(%) vắng'] = pd.to_numeric(combined_data['(%) vắng'], errors='coerce').fillna(0)

    combined_data.rename(columns={
        'Họ đệm': 'ho_dem',
        'Tên': 'ten',
        'Giới tính': 'gioi_tinh',
        'Ngày sinh': 'ngay_sinh',
        'Vắng có phép': 'vang_co_phep',
        'Vắng không phép': 'vang_khong_phep',
        'Tổng số tiết': 'tong_so_tiet',
        '(%) vắng': 'ty_le_vang',
        'Tổng buổi vắng': 'tong_buoi_vang'
    }, inplace=True)

    # Xóa tất cả mục trong giao diện (tree) sau khi tổng hợp
    for item in tree.get_children():
        tree.delete(item)

    conn = sqlite3.connect('tonghopsv.db')
    cursor = conn.cursor()

    cursor.execute("SELECT * FROM tonghopsv")
    rows = cursor.fetchall()
        
    # Hiển thị dữ liệu đã tải vào Treeview với cột STT
    for row in rows:
        stt = len(tree.get_children()) + 1  # Tạo STT tự động
        tree.insert('', 'end', values=[stt] + list(row))

    conn.close()

    # Lưu mã lớp của các sinh viên có vắng > 50%
    # class_codes = combined_data[combined_data['ty_le_vang'] >= 50.0]['Mã lớp'].unique().tolist()
      
def save_absent_students_to_excel(threshold=30.0):
    global summary_file, class_codes

    # Kết nối tới cơ sở dữ liệu để lấy dữ liệu sinh viên
    conn = sqlite3.connect('tonghopsv.db')
    cursor = conn.cursor()

    # Truy vấn tất cả sinh viên trong bảng tonghopsv để lấy mã lớp
    query_all_classes = "SELECT DISTINCT `ma_lop` FROM tonghopsv"
    cursor.execute(query_all_classes)
    all_class_rows = cursor.fetchall()

    # Lưu mã lớp của tất cả sinh viên
    class_codes = [row[0] for row in all_class_rows]
    
    # Truy vấn sinh viên có tỷ lệ vắng lớn hơn ngưỡng (threshold) để lưu vào file Excel
    query_absent_students = f"SELECT * FROM tonghopsv WHERE ty_le_vang >= {threshold}"
    cursor.execute(query_absent_students)
    absent_rows = cursor.fetchall()

    conn.close()

    if absent_rows:
        # Tạo DataFrame từ dữ liệu sinh viên có tỷ lệ vắng > threshold
        df_absent_students = pd.DataFrame(absent_rows, columns=[
            'MSSV', 'Họ đệm', 'Tên', 'Giới tính', 'Ngày sinh', 'Vắng có phép', 
            'Vắng không phép', 'Tổng số tiết', '(%) vắng', 'Tổng buổi vắng', 
            'Đợt', 'Mã lớp', 'Tên môn học'])

        # Lưu sinh viên vắng nhiều vào tệp Excel
        summary_file = 'TongHopSinhVienVangNhieu.xlsx'
        df_absent_students.to_excel(summary_file, index=False)

        print(f"Tệp tổng hợp sinh viên vắng nhiều đã được lưu tại: {summary_file}")
        print(f"Mã lớp liên quan (tất cả sinh viên): {class_codes}")

        # Gọi hàm gửi email với tệp Excel đính kèm
        send_email_with_attachment(summary_file, class_codes)
    else:
        print("Không có sinh viên nào vượt quá ngưỡng vắng!")
        return None, []

def send_email_with_attachment(summary_file, class_codes):
    sender_email = "carotneee4@gmail.com" 
    sender_password = "bgjx tavb oxba ickr"
    recipient_email = "tranhuuhauthh@gmail.com"

    # Kiểm tra tệp trước khi gửi
    if not summary_file or not os.path.exists(summary_file):
        print("Không tìm thấy tệp Excel để gửi email.")
        return

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = "Báo cáo sinh viên vắng nhiều"

    # Tạo phần thân email
    if class_codes:
        body = "Đây là báo cáo tổng hợp sinh viên vắng nhiều của tất cả các lớp: " + ', '.join(class_codes)
    else:
        body = "Không có sinh viên nào vượt quá ngưỡng vắng."

    msg.attach(MIMEText(body, 'plain'))

    # Đính kèm tệp Excel nếu có
    try:
        with open(summary_file, "rb") as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={summary_file}')
            msg.attach(part)

        # Gửi email
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()  # Bật chế độ bảo mật
            server.login(sender_email, sender_password)
            server.send_message(msg)
            print("Gửi email thành công!")
            messagebox.showinfo("Email Success", f"Email đã gửi thành công tới {recipient_email}")
    except FileNotFoundError:
        print("Tệp không tồn tại hoặc không thể mở.")
        messagebox.showerror("Email Error", "Tệp không tồn tại hoặc không thể mở.")
    except Exception as e:
        print(f"Có lỗi xảy ra khi gửi email: {e}")
        messagebox.showerror("Email Error", f"Có lỗi xảy ra khi gửi email: {e}")

# Đặt font mặc định là Times New Roman cho biểu đồ
rcParams['font.family'] = 'Times New Roman'    
    
# dùng để vẽ biểu đồ tỷ lệ vắng
def get_data_from_sqlite():
    conn = sqlite3.connect('students.db')
    cursor = conn.cursor()
    
    # Lấy dữ liệu MSSV, Họ tên và tỷ lệ vắng của sinh viên
    cursor.execute("SELECT mssv, ho_dem || ' ' || ten as ho_ten, ty_le_vang FROM students")
    data = cursor.fetchall()
    
    # Lấy dữ liệu lớp và số buổi vắng
    cursor.execute("SELECT ma_lop, SUM(vang_co_phep + vang_khong_phep) AS tong_buoi_vang FROM students GROUP BY ma_lop")
    class_data = cursor.fetchall()

    conn.close()
    
    return data, class_data

def plot_student_absence_chart(student_data):
    fig, ax = plt.subplots(figsize=(16, 8)) 
    
    # Tên sinh viên
    names = [row[1] for row in student_data]
    # Tỷ lệ vắng
    absence_rates = [row[2] for row in student_data]

    # Vẽ biểu đồ cột dọc
    ax.bar(names, absence_rates, color='pink', width=0.4)

    # Thiết lập nhãn cho trục y và tiêu đề với font Times New Roman
    ax.set_ylabel('Tỷ lệ vắng (%)', fontsize=10)
    ax.set_title('Tỷ lệ vắng của sinh viên', fontsize=12)

    # Xoay nhãn trục x và chỉnh kích thước
    ax.tick_params(axis='x', labelsize=8, rotation=45)

    # Đặt nhãn tương ứng với các vị trí của cột
    ax.set_xticks(range(len(names)))  # Đảm bảo các nhãn trên trục x được đặt tại đúng vị trí
    ax.set_xticklabels(names, rotation=45, ha="right")

    # Tạo thêm khoảng trống giữa các cột
    ax.margins(x=0.1)  # Giảm bớt khoảng cách giữa các cột và lề để khớp tên và cột

    # Sử dụng tight_layout để điều chỉnh lại bố cục
    plt.tight_layout()

    return fig

def show_student_chart():
    student_data, _ = get_data_from_sqlite()
    fig = plot_student_absence_chart(student_data)
    
    # Mở cửa sổ mới để hiển thị biểu đồ
    new_window = tk.Toplevel()
    new_window.title("Biểu đồ tỷ lệ vắng sinh viên")
    window_width = 1300  # Chiều rộng của cửa sổ
    window_height = 700  # Chiều cao của cửa sổ
    new_window.geometry(f"{window_width}x{window_height}")

    canvas = FigureCanvasTkAgg(fig, master=new_window)
    canvas.draw()
    canvas.get_tk_widget().pack(fill='both', expand=True)
    
#dùng đê vẽ biểu đồ vắng có phếp / không phép
def get_absence_types_data():
    conn = sqlite3.connect('students.db')
    cursor = conn.cursor()
    
    # Lấy tổng số vắng có phép và vắng không phép
    cursor.execute("""
        SELECT SUM(vang_co_phep) AS tong_vang_co_phep,
               SUM(vang_khong_phep) AS tong_vang_khong_phep
        FROM students
    """)
    absence_data = cursor.fetchone()
    conn.close()
    
    return absence_data

def plot_absence_types_chart(absence_data):
    fig, ax = plt.subplots(figsize=(6, 6))  # Điều chỉnh kích thước biểu đồ tròn

    labels = ['Vắng có phép', 'Vắng không phép']
    values = [absence_data[0], absence_data[1]]
    colors = ['#A0B4F2', 'pink']  

    # Vẽ biểu đồ tròn
    ax.pie(values, labels=labels, colors=colors, autopct='%1.1f%%', startangle=90)

    # Đảm bảo biểu đồ tròn là hình tròn (không bị méo)
    ax.axis('equal')

    # Thiết lập tiêu đề
    ax.set_title('Tỷ lệ vắng có phép và vắng không phép', fontsize=12)

    return fig

def show_absence_types_chart():
    absence_data = get_absence_types_data()
    fig = plot_absence_types_chart(absence_data)

    # Mở cửa sổ mới để hiển thị biểu đồ
    new_window = tk.Toplevel()
    new_window.title("Biểu đồ vắng có phép và vắng không phép")
    window_width = 600  # Chiều rộng của cửa sổ
    window_height = 600  # Chiều cao của cửa sổ
    new_window.geometry(f"{window_width}x{window_height}")

    canvas = FigureCanvasTkAgg(fig, master=new_window)
    canvas.draw()
    canvas.get_tk_widget().pack(fill='both', expand=True)
    
# Chỉnh sửa để kích hoạt các nút sau khi tải file
def enable_buttons():
    add_button.config(state=NORMAL)
    edit_button.config(state=NORMAL)
    delete_button.config(state=NORMAL)
    sort_button.config(state=NORMAL)
    student_chart_button.config(state=NORMAL)
    absence_types_chart_button.config(state=NORMAL)
    send_warning_email_button.config(state=NORMAL)
    view_detail_button.config(state=NORMAL)

# Cập nhật khi tải file thành công sẽ kích hoạt các nút
def load_and_enable():
    load_from_excel_to_treeview(tree)
    enable_buttons()  # Kích hoạt các nút sau khi tải file

def initialize_user_database():
    connection = sqlite3.connect('students.db')
    cursor = connection.cursor()

    # Tạo bảng users với ràng buộc UNIQUE cho username
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL UNIQUE,
            password TEXT NOT NULL
        )
    ''')

    # Kiểm tra xem người dùng admin đã tồn tại chưa
    cursor.execute('SELECT * FROM users WHERE username = ?', ('123',))
    result = cursor.fetchone()

    # Nếu người dùng admin chưa tồn tại, thì thêm vào
    if result is None:
        cursor.execute('''
            INSERT INTO users (username, password) 
            VALUES (?, ?)
        ''', ('123', '123'))

    connection.commit()
    connection.close()

def login():
    username = username_entry.get()
    password = password_entry.get()

    # Connect to SQLite database
    connection = sqlite3.connect('students.db')
    cursor = connection.cursor()

    # Query to check if user exists
    cursor.execute("SELECT * FROM users WHERE username = ? AND password = ?", (username, password))
    result = cursor.fetchone()

    if result:
        messagebox.showinfo("Login Successful", "Welcome!")
        login_window.destroy()  # Close login window and open the main app
        main()  # Call the main app function after successful login
    else:
        messagebox.showerror("Login Failed", "Invalid username or password")
    
    connection.close()

# Hàm hiển thị form đăng nhập
def show_login_form():
    global login_window, username_entry, password_entry

    login_window = Tk()
    login_window.title("Login")

    # Thiết lập kích thước và căn giữa cửa sổ đăng nhập
    window_width = 400
    window_height = 300
    screen_width = login_window.winfo_screenwidth()
    screen_height = login_window.winfo_screenheight()
    position_top = int(screen_height/2 - window_height/2)
    position_right = int(screen_width/2 - window_width/2)
    login_window.geometry(f'{window_width}x{window_height}+{position_right}+{position_top}')
    login_window.configure(bg="#F2D0D3")  # Màu nền xám nhạt

    # Thiết kế nhãn tiêu đề
    title_label = Label(login_window, text="Login", font=("Times New Roman", 24, "bold"), bg="#F2D0D3", fg="#333333")
    title_label.pack(pady=20)

    # Nhãn và ô nhập cho Username
    username_label = Label(login_window, text="Username", font=("Times New Roman", 12), bg="#F2D0D3", fg="#333333")
    username_label.pack(pady=5)
    username_entry = Entry(login_window, font=("Times New Roman", 12), width=30, bd=2, relief="groove")
    username_entry.pack()

    # Nhãn và ô nhập cho Password
    password_label = Label(login_window, text="Password", font=("Times New Roman", 12), bg="#F2D0D3", fg="#333333")
    password_label.pack(pady=5)
    password_entry = Entry(login_window, font=("Times New Roman", 12), width=30, bd=2, relief="groove", show="*")
    password_entry.pack()

    # Nút đăng nhập
    # Nút đăng nhập (giống nút load_button)
    login_button = Button(login_window, text="Login", command=login, bg="#F2A2C0", fg='black', font=("Times New Roman", 10))  
    login_button.pack(pady=40)  # Căn giống với load_button

    # Vòng lặp giao diện
    login_window.mainloop()

#hàm thực hiện việc lập lịch, kiểm tra email và gửi thông báo dựa trên ngày và giờ cụ thể.   
def start_scheduler():
    global class_codes, summary_file  # Mã lớp từ bảng tonghop

    while True:
        now = datetime.now()
        
        # Kiểm tra email và xử lý
        check_emails_and_process()  # Kiểm tra email đến và xử lý
        
        # Kiểm tra xem ngày hiện tại là ngày 1 hoặc 25 và thời gian là đúng 12:00
        if (now.day == 1 or now.day == 25) and now.hour == 12 and now.minute == 00:
            print("Đủ điều kiện gửi email. Gửi email...")

            # Gọi hàm send_email_with_attachment với đường dẫn tệp và mã lớp từ bảng tonghop
            save_absent_students_to_excel()

        else:
            print(f"Hiện tại là {now.strftime('%Y-%m-%d %H:%M:%S')} - Không đủ điều kiện để gửi email.")
        
        time.sleep(40)  # Sau mỗi lần kiểm tra, nó sẽ chờ 40 giây trước khi lặp lại

def check_emails_and_process():
    # Thông tin đăng nhập email
    IMAP_SERVER = "imap.gmail.com"
    EMAIL_ACCOUNT = "tranhuuhauthh@gmail.com"
    PASSWORD = "jmny hcmf voxq ekbj"  

    # Kết nối tới server IMAP
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL_ACCOUNT, PASSWORD)
    mail.select("inbox")

    # Tìm email chưa đọc (Unread emails)
    status, messages = mail.search(None, '(UNSEEN)')
    
    # Kiểm tra xem có email nào chưa đọc
    if status != "OK" or not messages[0]:
        print("Không có email mới")
        return

    email_ids = messages[0].split()

    email_class_codes = []  # Biến lưu trữ mã lớp lấy từ email
    for email_id in email_ids:
        status, msg_data = mail.fetch(email_id, "(RFC822)")
        if status != "OK":
            print(f"Lỗi khi tải email ID {email_id}")
            continue
        
        # Đọc email và giải mã nội dung
        for response_part in msg_data:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])
                subject, encoding = decode_header(msg["Subject"])[0]
                if isinstance(subject, bytes):
                    subject = subject.decode(encoding if encoding else "utf-8")

                from_email = msg.get("From")
                print(f"Đang xử lý email từ: {from_email} - Chủ đề: {subject}")

                # Lấy nội dung email
                if msg.is_multipart():
                    for part in msg.walk():
                        if part.get_content_type() == "text/plain":
                            body = part.get_payload(decode=True).decode(part.get_content_charset())
                            class_codes_from_email = extract_class_codes_from_message(body)
                            email_class_codes.extend(class_codes_from_email)
                            print(f"Mã lớp nhận được từ email: {class_codes_from_email}")

    if email_class_codes:
        send_late_report_email(from_email, email_class_codes)

    mail.logout()

def extract_class_codes_from_message(body):
    # Tìm và tách các mã lớp từ nội dung email theo định dạng đã cho
    match = re.search(r"Đây là báo cáo tổng hợp sinh viên vắng nhiều của tất cả các lớp: (.+)", body)
    if match:
        class_codes = match.group(1).split(", ")
        return class_codes
    return []

def send_late_report_email(from_email, email_class_codes):
    # Kiểm tra hạn chót (giả sử hạn chót là ngày 15 và 30 hàng tháng)
    today = datetime.today()
    if today.day > 15 and today.day < 30:
        # Tạo nội dung báo cáo
        subject = "Báo cáo quản lý về lớp trễ hạn"
        body = f"Người gửi: {from_email}\nLớp: {', '.join(email_class_codes)}\nTình trạng: Trễ hạn"
        recipient_email = "tranhuuhauthh@gmail.com"  # Email quản lý

        send_email(recipient_email, subject, body)

# def send_question(student_email, staff_email, manager_email, question):
#     try:
#         smtp_server = 'smtp.gmail.com'
#         smtp_port = 587
#         smtp_user = 'carotneee4@gmail.com'
#         smtp_password = 'bgjx tavb oxba ickr'  # Mật khẩu thực tế

#         # Tạo nội dung email
#         message = MIMEMultipart()
#         message['From'] = student_email
#         message['To'] = staff_email
#         message['Subject'] = f'Câu hỏi từ sinh viên: {student_email}'

#         body = f'''Chào nhân viên,

#         Sinh viên {student_email} đã gửi câu hỏi:

#         "{question}"

#         Vui lòng trả lời câu hỏi này trong thời gian sớm nhất. Nếu không có phản hồi trong 24 giờ, câu hỏi sẽ được nhắc nhở gửi tới quản lý {manager_email}.

#         Trân trọng,
#         Hệ thống hỗ trợ học vụ
#         '''
#         message.attach(MIMEText(body, 'plain'))

#         server = smtplib.SMTP(smtp_server, smtp_port)
#         server.starttls()
#         server.login(smtp_user, smtp_password)
#         text = message.as_string()
#         server.sendmail(student_email, staff_email, text)
#         server.quit()

#         print(f'Đã gửi email câu hỏi đến nhân viên: {staff_email}')

#         # Theo dõi câu hỏi và lập lịch nhắc nhở
#         track_question(student_email, staff_email, manager_email, question)

#     except Exception as e:
#         print(f'Không thể gửi email: {str(e)}')

# def track_question(student_email, staff_email, manager_email, question):
#     submission_time = datetime.now()
#     print(f"Đã gửi câu hỏi lúc: {submission_time.strftime('%Y-%m-%d %H:%M:%S')}")

#     # Lập lịch kiểm tra phản hồi sau 30 giây
#     schedule.every(10).seconds.do(check_response, student_email, staff_email, manager_email, question)

# def check_response(student_email, staff_email, manager_email, question):
#     print(f"Kiểm tra phản hồi cho câu hỏi từ {student_email}...")
    
#     if not check_if_answered(student_email, question):
#         print("Không có phản hồi. Gửi nhắc nhở cho quản lý.")
#         send_reminder_to_manager(student_email, staff_email, manager_email, question)
#     else:
#         print("Câu hỏi đã được trả lời.")

# def send_reminder_to_manager(student_email, staff_email, manager_email, question):
#     try:
#         smtp_server = 'smtp.gmail.com'
#         smtp_port = 587
#         smtp_user = 'carotneee4@gmail.com'
#         smtp_password = 'bgjx tavb oxba ickr'  # Mật khẩu thực tế

#         message = MIMEMultipart()
#         message['From'] = smtp_user
#         message['To'] = manager_email
#         message['Subject'] = f'Nhắc nhở: Không có phản hồi cho câu hỏi từ {student_email}'

#         body = f'''Chào Quản lý,

#         Sinh viên {student_email} đã gửi câu hỏi tới {staff_email} nhưng chưa nhận được phản hồi trong 30 giây. 
#         Vui lòng kiểm tra và hỗ trợ.

#         Câu hỏi: {question}

#         Trân trọng,
#         Hệ thống hỗ trợ học vụ
#         '''
#         message.attach(MIMEText(body, 'plain'))

#         server = smtplib.SMTP(smtp_server, smtp_port)
#         server.starttls()
#         server.login(smtp_user, smtp_password)
#         text = message.as_string()
#         server.sendmail(smtp_user, manager_email, text)
#         server.quit()

#         print(f'Đã gửi email nhắc nhở đến quản lý: {manager_email}')
#     except Exception as e:
#         print(f'Không thể gửi email nhắc nhở: {str(e)}')

# def check_if_answered(student_email, question):
#     # Thay thế bằng cách kiểm tra dữ liệu thực tế
#     answered_questions = []  # Danh sách câu hỏi đã trả lời
#     return question in answered_questions

# Thiết lập cơ sở dữ liệu để lưu trữ câu hỏi và phản hồi
def update_button_states():
    if len(tree.get_children()) == 0:  # Kiểm tra nếu Treeview rỗng
        # Tắt các nút
        buttons = [add_button, edit_button, delete_button, sort_button, 
                   send_warning_email_button, view_detail_button, 
                   student_chart_button, absence_types_chart_button,
                   summarize_button, send_summary_email_button, refresh_button]
        
        for button in buttons:
            button.config(state=tk.DISABLED)
    else:
        # Bật các nút
        add_button.config(state=tk.NORMAL)
        edit_button.config(state=tk.NORMAL)
        delete_button.config(state=tk.NORMAL)
        sort_button.config(state=tk.NORMAL)
        send_warning_email_button.config(state=tk.NORMAL)
        view_detail_button.config(state=tk.NORMAL)
        student_chart_button.config(state=tk.NORMAL)
        absence_types_chart_button.config(state=tk.NORMAL)
        summarize_button.config(state=tk.NORMAL)
        send_summary_email_button.config(state=tk.NORMAL)
        refresh_button.config(state=tk.NORMAL)
        
def main():
    global df_sinh_vien, ma_lop, ten_mon_hoc, summary_file
    global chart_frame  
    global tree  # Declare tree as a global variable
    global add_button, edit_button, delete_button, sort_button, student_chart_button, absence_types_chart_button, send_warning_email_button, view_detail_button, summarize_button, send_summary_email_button, refresh_button
    root = Tk()
    root.title("Quản Lý Sinh Viên")
    
    # Thay đổi màu nền cho cửa sổ chính
    root.configure(bg="#F2D0D3")  # Màu nền chính

    # Thêm logo vào tiêu đề của ứng dụng
    logo_icon = Image.open("GUI_Tkinter/logoSGu.png")
    logo_icon = logo_icon.resize((32, 32), Image.LANCZOS)
    logo_icon_photo = ImageTk.PhotoImage(logo_icon)
    root.iconphoto(False, logo_icon_photo)

    # Đặt kích thước và vị trí cho giao diện chính
    root.geometry("1500x750+10+20")
    
    # Tạo style cho các nút
    style = ttk.Style()
    style.configure("TButton", font=("Times New Roman", 10), padding=6)

    # Thêm logo vào giao diện
    logo_image = Image.open("GUI_Tkinter/logocnttsgu.png")
    logo_image = logo_image.resize((240, 50), Image.LANCZOS)
    logo_photo = ImageTk.PhotoImage(logo_image)
    logo_label = Label(root, image=logo_photo, bg="#F2D0D3")  # Màu nền logo
    logo_label.image = logo_photo
    logo_label.pack(side=TOP, pady=10)

    # Tạo Treeview để hiển thị dữ liệu sinh viên
    tree = ttk.Treeview(root, columns=('STT', 'MSSV', 'Họ đệm', 'Tên', 'Giới tính', 'Ngày sinh', 'Vắng có phép', 'Vắng không phép', 'Tổng số tiết', '(%) vắng', 'Tổng buổi vắng', 'Đợt', 'Mã lớp', 'Tên môn học'), show='headings')
    
    style.configure("Treeview", font=("Times New Roman", 10), rowheight=25)
    style.configure("Treeview.Heading", font=("Times New Roman", 11, "bold"), background="#4CAF50", foreground="black")
    style.map("Treeview", background=[("selected", "#A3C1DA")], foreground=[("selected", "black")])

    # Tùy chỉnh chiều rộng cột
    tree.column("STT", width=40, anchor="center")  
    tree.column("MSSV", width=100, anchor="center")
    tree.column("Họ đệm", width=150, anchor="center")
    tree.column("Tên", width=80, anchor="center")
    tree.column("Giới tính", width=80, anchor="center")
    tree.column("Ngày sinh", width=120, anchor="center")
    tree.column("Vắng có phép", width=120, anchor="center")
    tree.column("Vắng không phép", width=120, anchor="center")
    tree.column("Tổng số tiết", width=120, anchor="center")
    tree.column("(%) vắng", width=80, anchor="center")
    tree.column("Tổng buổi vắng", width=120, anchor="center")
    tree.column("Đợt", width=100, anchor="center")
    tree.column("Mã lớp", width=100, anchor="center")
    tree.column("Tên môn học", width=150, anchor="center")

    for col in tree['columns']:
        tree.heading(col, text=col)

    tree.pack(fill='both', expand=True, padx=10, pady=10)

    # Tạo frame chứa các nút bên trái
    left_frame = Frame(root, bg="#F2D0D3")  # Màu nền frame bên trái
    left_frame.pack(side=LEFT, padx=10, pady=10, fill='y')

    # Tạo frame chứa các nút ở giữa
    center_frame = Frame(root, bg="#F2D0D3")  # Màu nền frame ở giữa
    center_frame.pack(side=LEFT, padx=10, pady=10, fill='y')

    # Thêm giao diện tìm kiếm vào center_frame trước các nút khác
    add_search_interface(center_frame, tree)  # Thêm interface tìm kiếm lên trên cùng

    # Các nút nằm bên trái, với độ rộng cố định
    button_width = 20
    button_color = "#F2A2C0"  # Thay đổi màu các nút

    load_button = Button(left_frame, text="Tải file", command=lambda: load_from_excel_to_treeview(tree), width=button_width, bg=button_color, fg='black', font=("Times New Roman", 10))
    load_button.pack(anchor='w', pady=5, fill='x')

    # Định nghĩa các nút
    add_button = tk.Button(left_frame, text="Thêm sinh viên", command=lambda: add_student(tree), width=button_width, bg=button_color, fg='black', font=("Times New Roman", 10), state=tk.DISABLED)
    add_button.pack(anchor='w', pady=5, fill='x')

    edit_button = tk.Button(left_frame, text="Sửa sinh viên", command=lambda: edit_student(tree), width=button_width, bg=button_color, fg='black', font=("Times New Roman", 10), state=tk.DISABLED)
    edit_button.pack(anchor='w', pady=5, fill='x')

    delete_button = tk.Button(left_frame, text="Xóa sinh viên", command=lambda: delete_student(tree), width=button_width, bg=button_color, fg='black', font=("Times New Roman", 10), state=tk.DISABLED)
    delete_button.pack(anchor='w', pady=5, fill='x')

    sort_button = tk.Button(left_frame, text="Sắp xếp", command=lambda: sort_students_by_absences(tree), width=button_width, bg=button_color, fg='black', font=("Times New Roman", 10), state=tk.DISABLED)
    sort_button.pack(anchor='w', pady=5, fill='x')

    send_warning_email_button = tk.Button(left_frame, text="Gửi Email cảnh báo", command=send_warning_emails, width=button_width, bg=button_color, fg='black', font=("Times New Roman", 10), state=tk.DISABLED)
    send_warning_email_button.pack(anchor='w', pady=5, fill='x')

    view_detail_button = tk.Button(left_frame, text="Xem Chi Tiết", command=lambda: view_details(tree), width=button_width, bg=button_color, fg='black', font=("Times New Roman", 10), state=tk.DISABLED)
    view_detail_button.pack(anchor='w', pady=5, fill='x')

    # Các nút nằm ở giữa
    student_chart_button = tk.Button(center_frame, text="Biểu đồ % vắng", command=show_student_chart, width=button_width, bg=button_color, fg='black', font=("Times New Roman", 10), state=tk.DISABLED)
    student_chart_button.pack(anchor='center', pady=10)

    absence_types_chart_button = tk.Button(center_frame, text="Biểu đồ vắng P/K", command=show_absence_types_chart, width=button_width, bg=button_color, fg='black', font=("Times New Roman", 10), state=tk.DISABLED)
    absence_types_chart_button.pack(anchor='center', pady=10)

    summarize_button = tk.Button(center_frame, text="Xóa dữ liệu", command=lambda: clear_table(tree), width=button_width, bg=button_color, fg='black', font=("Times New Roman", 10), state=tk.DISABLED)
    summarize_button.pack(anchor='center', pady=10)

    send_summary_email_button = tk.Button(center_frame, text="Gửi Email tổng hợp", 
                                        command=lambda: save_absent_students_to_excel() if summary_file else print("Không có tệp tóm tắt để gửi!"), 
                                        width=button_width, bg=button_color, fg='black', font=("Times New Roman", 10), state=tk.DISABLED)
    send_summary_email_button.pack(anchor='center', pady=10)
    
    refresh_button = Button(center_frame, text="Refresh", command=lambda: refresh_treeview(tree), width=button_width, bg=button_color, fg='black', font=("Times New Roman", 10), state=tk.DISABLED)
    refresh_button.pack(anchor='center', pady=10)
    
    # send_question_button = Button(center_frame, text="send question", command=lambda: send_question('carotneee4@gmail.com', 'tranhuuhauthh@gmail.com', 'tranhuuhauthh@gmail.com', 'Câu hỏi mẫu từ sinh viên'), width=button_width, bg=button_color, fg='black', font=("Times New Roman", 10))
    # send_question_button.pack(anchor='center', pady=10)
    
    initialize_database()
    clear_table(tree)
    refresh_treeview(tree) 
    
       
    # Khởi tạo chart_frame
    chart_frame = Frame(root, bg="#F2D0D3")  # Màu nền chart_frame
    chart_frame.pack(fill='both', expand=True)
    
    update_button_states()
    # Gán hàm cho nút tải file
    # load_button.config(command=load_and_enable)    
      
if __name__ == "__main__":
    # Khởi tạo cơ sở dữ liệu người dùng
    initialize_user_database()
    
    # Tạo một luồng riêng cho chức năng gửi email tự động
    email_thread = threading.Thread(target=start_scheduler)
    email_thread.daemon = True  # Đảm bảo chương trình chính dừng, luồng này cũng sẽ dừng
    email_thread.start()

    # Hiển thị form đăng nhập
    show_login_form()
    
    # Để giữ cho chương trình hoạt động, có thể cần một vòng lặp chính
    while True:
        # schedule.run_pending()  # Thêm dòng này để chạy các tác vụ đã lên lịch
        time.sleep(1)
        
