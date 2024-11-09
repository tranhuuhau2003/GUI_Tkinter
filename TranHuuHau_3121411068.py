import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import sqlite3
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import time
import schedule
import threading
from datetime import datetime
from tabulate import tabulate
        
        
        
        
class ExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Data Viewer")
        self.filepath = ''  # Để trống ban đầu, sẽ được cập nhật khi chọn file
        
        # Kết nối SQLite và tạo bảng nếu chưa có
        self.conn = sqlite3.connect('sinhvien.db')
        self.cursor = self.conn.cursor()
        
    

        self.start_scheduler()
        
        # Giao diện chính
        self.main_menu()
        self.clear_students_table()
    
    def select_file(self):
        self.filepath = filedialog.askopenfilename(title="Chọn File Excel", filetypes=[("Excel Files", "*.xls;*.xlsx")])
        if self.filepath:
            self.load_excel_toDB()

    def load_excel_toDB(self):
        try:
            df = pd.read_excel(self.filepath, engine='xlrd', header=None)
            df = df.fillna('')

            # Lấy thông tin "Đợt", "Mã lớp học phần", "Tên môn học"
            dot = df.iloc[5, 2]  # Đợt nằm ở hàng 5, cột 2
            ma_lop = df.iloc[7, 2]  # Mã lớp học phần nằm ở hàng 7, cột 2
            ten_mon_hoc = df.iloc[8, 2]  # Tên môn học nằm ở hàng 8, cột 2
                        
             # Lấy dữ liệu sinh viên từ hàng 13 trở đi
            df_sinh_vien = df.iloc[13:, [1, 2, 3, 4, 5]]  # Chỉ lấy các cột cần thiết
            df_sinh_vien.columns = ['MSSV', 'Họ đệm', 'Tên', 'Giới tính', 'Ngày sinh']
            
            self.mssv_list = df_sinh_vien['MSSV'].tolist()

            # Xóa bảng cũ nếu tồn tại
            self.cursor.execute("DROP TABLE IF EXISTS students")

            self.cursor.execute(""" 
            CREATE TABLE IF NOT EXISTS students (
                mssv TEXT PRIMARY KEY,
                ho_dem TEXT,
                ten TEXT,
                gioi_tinh TEXT,
                ngay_sinh TEXT,
                dot TEXT,             -- Cột mới: Đợt
                ma_lop TEXT,          -- Cột mới: Mã lớp
                ten_mon_hoc TEXT,    -- Cột mới: Tên môn học
                email_student TEXT
            )
            """)
            
            self.cursor.execute('DELETE FROM students')
            # Lưu dữ liệu vào SQLite
            for index, row in df_sinh_vien.iterrows():
                self.cursor.execute("""
                    INSERT OR IGNORE INTO students (mssv, ho_dem, ten, gioi_tinh, ngay_sinh, dot, ma_lop, ten_mon_hoc)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    row['MSSV'],
                    row['Họ đệm'],
                    row['Tên'],
                    row['Giới tính'],
                    str(row['Ngày sinh']),
                    dot,          # Đợt
                    ma_lop,       # Mã lớp học phần
                    ten_mon_hoc   # Tên môn học
                ))
 

            df_temp = df.iloc[13:, [1, 6, 9, 12, 15, 18, 21, 24, 25, 26, 27]]
            df_temp.columns = ['MSSV', '11/06/2024', '18/06/2024', '25/06/2024', '02/07/2024', '09/07/2024', '23/07/2024', 'Vắng có phép', 'Vắng không phép', 'Tổng số tiết', 'Tỷ lệ vắng']
         

            # Xóa bảng cũ nếu tồn tại
            self.cursor.execute("DROP TABLE IF EXISTS attendance")
            # Tạo bảng nếu chưa có
            self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS attendance (
                MSSV TEXT PRIMARY KEY,
                "11/06/2024" TEXT,
                "18/06/2024" TEXT,
                "25/06/2024" TEXT,
                "02/07/2024" TEXT,
                "09/07/2024" TEXT,
                "23/07/2024" TEXT,
                "Vắng có phép" INTEGER,
                "Vắng không phép" INTEGER,
                "Tổng số tiết" INTEGER,
                "Tỷ lệ vắng" TEXT
            )
            ''')

            self.cursor.execute('DELETE FROM attendance')
            # Chèn dữ liệu vào bảng
            for _, row in df_temp.iterrows():
                self.cursor.execute('''
                    INSERT OR REPLACE INTO attendance (MSSV, "11/06/2024", "18/06/2024", "25/06/2024", 
                    "02/07/2024", "09/07/2024", "23/07/2024", "Vắng có phép", "Vắng không phép", 
                    "Tổng số tiết", "Tỷ lệ vắng")
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''', 
                    (row['MSSV'], row['11/06/2024'], row['18/06/2024'], row['25/06/2024'], 
                    row['02/07/2024'], row['09/07/2024'], row['23/07/2024'], 
                    row['Vắng có phép'], row['Vắng không phép'], row['Tổng số tiết'], row['Tỷ lệ vắng']))


                # Tạo bảng parents
            self.cursor.execute('''
                CREATE TABLE IF NOT EXISTS parents (
                    mssv TEXT PRIMARY KEY,
                    email_ph TEXT  -- Email của phụ huynh
                )
            ''')

            # Tạo bảng teachers
            self.cursor.execute('''
                CREATE TABLE IF NOT EXISTS teachers (
                    mssv TEXT PRIMARY KEY,
                    email_gvcn TEXT  -- Email của giáo viên chủ nhiệm
                )
            ''')
            
             # Tạo bảng tbm
            self.cursor.execute('''
                CREATE TABLE IF NOT EXISTS tbm (
                    mssv TEXT PRIMARY KEY,
                    email_tbm TEXT  -- Email của tbm
                )
            ''')
            
            # Lưu thay đổi 
            self.conn.commit()
            
            # Thêm dữ liệu vào bảng parents
            for mssv in self.mssv_list:
                email_student = f"tranhuuhauthh@gmail.com"  # Tạo email mẫu cho phụ huynh
                self.cursor.execute('UPDATE students SET email_student = ? WHERE mssv = ?', (email_student, mssv))
        
            # Thêm dữ liệu vào bảng parents
            for mssv in self.mssv_list:
                email_ph = f"tranhuuhauthh@gmail.com"  # Tạo email mẫu cho phụ huynh
                self.cursor.execute('INSERT OR IGNORE INTO parents (mssv, email_ph) VALUES (?, ?)', (mssv, email_ph))

            # Thêm dữ liệu vào bảng teachers
            for mssv in self.mssv_list:
                email_gvcn = f"tranhuuhau2003@gmail.com"  # Tạo email mẫu cho giáo viên chủ nhiệm
                self.cursor.execute('INSERT OR IGNORE INTO teachers (mssv, email_gvcn) VALUES (?, ?)', (mssv, email_gvcn))
                
            # Thêm dữ liệu vào bảng teachers
            for mssv in self.mssv_list:
                email_tbm = f"tranhuuhauthh@gmail.com"  # Tạo email mẫu cho giáo viên chủ nhiệm
                self.cursor.execute('INSERT OR IGNORE INTO tbm (mssv, email_tbm) VALUES (?, ?)', (mssv, email_tbm))
            
            self.conn.commit()
            self.load_students_to_treeview()
        except Exception as e:
            print(f"Lỗi khi lưu thông tin vắng: {e}")
                      
    def clear_students_table(self):
        # Kiểm tra xem bảng 'students' có tồn tại hay không
        try:
            self.cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='students'")
            table_exists = self.cursor.fetchone()
            
            if table_exists:
                # Xóa toàn bộ dữ liệu trong bảng 'students' nếu tồn tại
                self.cursor.execute("DELETE FROM students")
                self.conn.commit()
                print("Tất cả dữ liệu trong bảng 'students' đã được xóa.")
            else:
                print("Bảng 'students' không tồn tại.")
        except Exception as e:
            print(f"Đã xảy ra lỗi khi xóa dữ liệu: {e}")

    def load_students_to_treeview(self):
        try:
            # Xóa dữ liệu hiện có trong Treeview
            for row in self.tree.get_children():
                self.tree.delete(row)

            # Truy vấn lấy dữ liệu từ bảng students và attendance
            self.cursor.execute("""
                SELECT students.mssv, students.ho_dem, students.ten, students.ngay_sinh, students.dot, 
                    students.ma_lop, students.ten_mon_hoc, 
                    attendance."Vắng có phép", attendance."Vắng không phép",
                    (attendance."Vắng có phép" + attendance."Vắng không phép") as tong_vang
                FROM students
                LEFT JOIN attendance ON students.mssv = attendance.MSSV
            """)
            rows = self.cursor.fetchall()

            # Chèn dữ liệu vào Treeview
            for i, row in enumerate(rows, 1):
                self.tree.insert('', 'end', values=(i, row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[9]))

        except Exception as e:
            print(f"Lỗi khi tải dữ liệu lên Treeview: {e}")
  
    def search_student(self):
        search_value = self.search_entry.get().lower()
        if search_value:
            query = """
                SELECT students.mssv, students.ho_dem, students.ten, students.ngay_sinh, students.dot, 
                    students.ma_lop, students.ten_mon_hoc, 
                    attendance."Vắng có phép", attendance."Vắng không phép",
                    (attendance."Vắng có phép" + attendance."Vắng không phép") as tong_vang
                FROM students
                LEFT JOIN attendance ON students.mssv = attendance.MSSV
                WHERE LOWER(students.mssv) LIKE ? OR LOWER(students.ten) LIKE ?
            """
            search_pattern = '%' + search_value + '%'
            self.cursor.execute(query, (search_pattern, search_pattern))
            result = self.cursor.fetchall()

            # Xóa dữ liệu hiện có trong Treeview
            self.tree.delete(*self.tree.get_children())

            # Chèn dữ liệu tìm kiếm vào Treeview
            for index, row in enumerate(result, start=1):
                self.tree.insert("", "end", values=(index,) + row)
        else:
            self.load_students_to_treeview()  # Nếu không có giá trị tìm kiếm, load lại toàn bộ dữ liệu
    
    def add_student(self):
        def save_student():
            mssv = mssv_entry.get().strip()  # Xóa khoảng trắng đầu và cuối
            ho_dem = ho_dem_entry.get().strip()
            ten = ten_entry.get().strip()
            gioi_tinh = gioi_tinh_var.get()
            ngay_sinh = ngay_sinh_entry.get().strip()
            ma_lop = ma_lop_entry.get().strip()
            ten_mon_hoc = ten_mon_hoc_entry.get().strip()
            dot = dot_entry.get().strip()
            vang_co_phep = int(vang_co_phep_entry.get() or 0)
            vang_khong_phep = int(vang_khong_phep_entry.get() or 0)
            tong_so_tiet = int(tong_so_tiet_entry.get() or 0)

            # Kiểm tra xem MSSV và Tên có rỗng hay không
            if not mssv or not ten:
                messagebox.showerror("Error", "MSSV và Tên sinh viên không được để trống!")
                return  # Ngừng thực hiện nếu MSSV hoặc Tên rỗng

            # Tính tỷ lệ vắng
            if tong_so_tiet > 0:
                ty_le_vang = (vang_co_phep + vang_khong_phep) / tong_so_tiet * 100
            else:
                ty_le_vang = 0.0  # Nếu không có tiết nào, tỷ lệ vắng sẽ là 0%

            try:
                # Lưu thông tin sinh viên vào bảng students
                self.cursor.execute("""
                    INSERT INTO students (mssv, ho_dem, ten, gioi_tinh, ngay_sinh, ma_lop, ten_mon_hoc, dot)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """, (mssv, ho_dem, ten, gioi_tinh, ngay_sinh, ma_lop, ten_mon_hoc, dot))

                # Lưu thông tin vắng vào bảng attendance
                self.cursor.execute("""
                    INSERT INTO attendance (MSSV, "Vắng có phép", "Vắng không phép", "Tổng số tiết", "Tỷ lệ vắng")
                    VALUES (?, ?, ?, ?, ?)
                """, (mssv, vang_co_phep, vang_khong_phep, tong_so_tiet, f"{ty_le_vang:.1f}"))

                # Lưu thay đổi vào cơ sở dữ liệu
                self.conn.commit()

                messagebox.showinfo("Info", "Student added successfully!")
                add_window.destroy()
                self.load_students_to_treeview()

            except Exception as e:
                messagebox.showerror("Error", f"Lỗi khi thêm sinh viên: {e}")

        # Tạo cửa sổ thêm sinh viên
        add_window = tk.Toplevel(self.root)
        add_window.title("Add Student")

        tk.Label(add_window, text="MSSV:").grid(row=0, column=0, padx=10, pady=10)
        tk.Label(add_window, text="Họ đệm:").grid(row=1, column=0, padx=10, pady=10)
        tk.Label(add_window, text="Tên:").grid(row=2, column=0, padx=10, pady=10)
        tk.Label(add_window, text="Giới tính:").grid(row=3, column=0, padx=10, pady=10)
        tk.Label(add_window, text="Ngày sinh:").grid(row=4, column=0, padx=10, pady=10)
        tk.Label(add_window, text="Mã lớp:").grid(row=5, column=0, padx=10, pady=10)
        tk.Label(add_window, text="Tên môn học:").grid(row=6, column=0, padx=10, pady=10)
        tk.Label(add_window, text="Đợt:").grid(row=7, column=0, padx=10, pady=10)
        tk.Label(add_window, text="Vắng có phép:").grid(row=8, column=0, padx=10, pady=10)
        tk.Label(add_window, text="Vắng không phép:").grid(row=9, column=0, padx=10, pady=10)
        tk.Label(add_window, text="Tổng số tiết:").grid(row=10, column=0, padx=10, pady=10)

        mssv_entry = tk.Entry(add_window)
        ho_dem_entry = tk.Entry(add_window)
        ten_entry = tk.Entry(add_window)
        gioi_tinh_var = tk.StringVar(value="Nam")  # Mặc định giới tính là Nam
        tk.Radiobutton(add_window, text="Nam", variable=gioi_tinh_var, value="Nam").grid(row=3, column=1)
        tk.Radiobutton(add_window, text="Nữ", variable=gioi_tinh_var, value="Nữ").grid(row=3, column=2)
        ngay_sinh_entry = tk.Entry(add_window)
        ma_lop_entry = tk.Entry(add_window)
        ten_mon_hoc_entry = tk.Entry(add_window)
        dot_entry = tk.Entry(add_window)
        vang_co_phep_entry = tk.Entry(add_window)
        vang_khong_phep_entry = tk.Entry(add_window)
        tong_so_tiet_entry = tk.Entry(add_window)

        mssv_entry.grid(row=0, column=1)
        ho_dem_entry.grid(row=1, column=1)
        ten_entry.grid(row=2, column=1)
        ngay_sinh_entry.grid(row=4, column=1)
        ma_lop_entry.grid(row=5, column=1)
        ten_mon_hoc_entry.grid(row=6, column=1)
        dot_entry.grid(row=7, column=1)
        vang_co_phep_entry.grid(row=8, column=1)
        vang_khong_phep_entry.grid(row=9, column=1)
        tong_so_tiet_entry.grid(row=10, column=1)

        tk.Button(add_window, text="Save student", command=save_student, bg="#4CAF50", fg="white").grid(row=11, column=0, columnspan=2, pady=10)
  
    def edit_student(self):
        selected_item = self.tree.selection()

        if not selected_item:
            messagebox.showwarning("Warning", "Vui lòng chọn một sinh viên để chỉnh sửa.")
            return

        item = self.tree.item(selected_item)
        mssv = item['values'][1]

        # Truy vấn để lấy thông tin sinh viên từ cơ sở dữ liệu
        self.cursor.execute("""
            SELECT s.mssv, s.ho_dem, s.ten, s.gioi_tinh, s.ngay_sinh, a."Vắng có phép", a."Vắng không phép", a."Tổng số tiết", a."Tỷ lệ vắng"
            FROM students s
            JOIN attendance a ON s.mssv = a.MSSV
            WHERE s.mssv = ?
        """, (mssv,))
        student_data = self.cursor.fetchone()

        if student_data is None:
            messagebox.showerror("Error", "Không tìm thấy sinh viên trong cơ sở dữ liệu.")
            return

        def save_edit():
            ho_dem = ho_dem_entry.get().strip()
            ten = ten_entry.get().strip()
            gioi_tinh = gioi_tinh_var.get()
            ngay_sinh = ngay_sinh_entry.get().strip()
            vắng_có_phép = int(vang_co_phap_entry.get() or 0)
            vắng_không_phép = int(vang_khong_phap_entry.get() or 0)
            tong_so_tiet = int(tong_so_tiet_entry.get() or 0)

            if not ten:
                messagebox.showerror("Error", "Tên sinh viên không được để trống!")
                return

            # Tính tỷ lệ vắng
            if tong_so_tiet > 0:
                ty_le_vang = (vắng_có_phép + vắng_không_phép) / tong_so_tiet * 100
            else:
                ty_le_vang = 0.0

            try:
                # Cập nhật thông tin sinh viên vào bảng students
                self.cursor.execute("""
                    UPDATE students
                    SET ho_dem = ?, ten = ?, gioi_tinh = ?, ngay_sinh = ?
                    WHERE mssv = ?
                """, (ho_dem, ten, gioi_tinh, ngay_sinh, mssv))

                # Cập nhật thông tin vắng vào bảng attendance
                self.cursor.execute("""
                    UPDATE attendance
                    SET "Vắng có phép" = ?, "Vắng không phép" = ?, "Tổng số tiết" = ?, "Tỷ lệ vắng" = ?
                    WHERE MSSV = ?
                """, (vắng_có_phép, vắng_không_phép, tong_so_tiet, f"{ty_le_vang:.1f}", mssv))

                self.conn.commit()
                messagebox.showinfo("Info", "Cập nhật sinh viên thành công!")
                edit_window.destroy()
                self.load_students_to_treeview()

            except Exception as e:
                messagebox.showerror("Error", f"Lỗi khi cập nhật sinh viên: {e}")

        # Tạo cửa sổ chỉnh sửa
        edit_window = tk.Toplevel(self.root)
        edit_window.title("Edit Student")

        # Lấy dữ liệu hiện tại để hiển thị vào các trường nhập liệu
        tk.Label(edit_window, text="MSSV:").grid(row=0, column=0, padx=10, pady=10)
        tk.Label(edit_window, text=student_data[0]).grid(row=0, column=1, padx=10, pady=10)

        tk.Label(edit_window, text="Họ đệm:").grid(row=1, column=0, padx=10, pady=10)
        tk.Label(edit_window, text="Tên:").grid(row=2, column=0, padx=10, pady=10)
        tk.Label(edit_window, text="Giới tính:").grid(row=3, column=0, padx=10, pady=10)
        tk.Label(edit_window, text="Ngày sinh:").grid(row=4, column=0, padx=10, pady=10)
        tk.Label(edit_window, text="Vắng có phép:").grid(row=5, column=0, padx=10, pady=10)
        tk.Label(edit_window, text="Vắng không phép:").grid(row=6, column=0, padx=10, pady=10)
        tk.Label(edit_window, text="Tổng số tiết:").grid(row=7, column=0, padx=10, pady=10)
        tk.Label(edit_window, text="Tỷ lệ vắng:").grid(row=8, column=0, padx=10, pady=10)

        ho_dem_entry = tk.Entry(edit_window)
        ho_dem_entry.insert(0, student_data[1])
        ten_entry = tk.Entry(edit_window)
        ten_entry.insert(0, student_data[2])

        gioi_tinh_var = tk.StringVar(value=student_data[3])
        tk.Radiobutton(edit_window, text="Nam", variable=gioi_tinh_var, value="Nam").grid(row=3, column=1)
        tk.Radiobutton(edit_window, text="Nữ", variable=gioi_tinh_var, value="Nữ").grid(row=3, column=2)

        ngay_sinh_entry = tk.Entry(edit_window)
        ngay_sinh_entry.insert(0, student_data[4])

        vang_co_phap_entry = tk.Entry(edit_window)
        vang_co_phap_entry.insert(0, student_data[5])

        vang_khong_phap_entry = tk.Entry(edit_window)
        vang_khong_phap_entry.insert(0, student_data[6])

        tong_so_tiet_entry = tk.Entry(edit_window)
        tong_so_tiet_entry.insert(0, student_data[7])

        ho_dem_entry.grid(row=1, column=1)
        ten_entry.grid(row=2, column=1)
        ngay_sinh_entry.grid(row=4, column=1)
        vang_co_phap_entry.grid(row=5, column=1)
        vang_khong_phap_entry.grid(row=6, column=1)
        tong_so_tiet_entry.grid(row=7, column=1)

        tk.Button(edit_window, text="Save Changes", command=save_edit, bg="#4CAF50", fg="white").grid(row=9, column=0, columnspan=2, pady=10)

    def remove_student(self):
        selected_item = self.tree.selection()
        if selected_item:
            # Lấy giá trị MSSV từ dòng được chọn
            student_mssv = self.tree.item(selected_item, "values")[1]  # MSSV nằm ở cột thứ 2 trong Treeview
            confirm = messagebox.askyesno("Xóa sinh viên", f"Bạn có chắc muốn xóa sinh viên với MSSV: {student_mssv}?")
            
            if confirm:
                try:
                    # Xóa sinh viên khỏi bảng students
                    self.cursor.execute("DELETE FROM students WHERE mssv = ?", (student_mssv,))
                    # Xóa thông tin điểm danh của sinh viên trong bảng attendance
                    self.cursor.execute("DELETE FROM attendance WHERE MSSV = ?", (student_mssv,))
                    self.conn.commit()

                    # Cập nhật lại Treeview sau khi xóa
                    self.load_students_to_treeview()

                    messagebox.showinfo("Info", f"Đã xóa sinh viên {student_mssv} thành công!")
                except Exception as e:
                    messagebox.showerror("Error", f"Có lỗi xảy ra: {e}")
            else:
                messagebox.showinfo("Info", "Xóa sinh viên đã bị hủy.")
        else:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn sinh viên để xóa!")
     
    def sort_students_by_absences(self):
        # Truy vấn lấy danh sách sinh viên và thông tin liên quan
        query = '''
            SELECT s.rowid AS STT, s.mssv, s.ho_dem, s.ten, s.ngay_sinh, s.dot, s.ma_lop, s.ten_mon_hoc, 
                a."Vắng có phép" + a."Vắng không phép" AS "Tổng vắng"
            FROM students s
            JOIN attendance a ON s.mssv = a.MSSV
            ORDER BY "Tổng vắng" DESC
        '''
        self.cursor.execute(query)
        sorted_students = self.cursor.fetchall()

        # Xóa dữ liệu cũ trong treeview
        for i in self.tree.get_children():      
            self.tree.delete(i)

        # Thêm dữ liệu đã sắp xếp vào treeview và định dạng
        for student in sorted_students:
            stt, mssv, ho_dem, ten, ngay_sinh, dot, ma_lop, ten_mon_hoc, tong_vang = student

            # Thêm sinh viên vào treeview
            self.tree.insert("", "end", values=(stt, mssv, ho_dem, ten, ngay_sinh, dot, ma_lop, ten_mon_hoc, tong_vang))

            # K1iểm tra và highlight sinh viên có tổng vắng > 10
            if tong_vang >= 10:
                # Chỉ đổi màu chữ cho sinh viên
                self.tree.item(self.tree.get_children()[-1], tags=("highlight",))

        # Định nghĩa tag cho treeview
        self.tree.tag_configure("highlight", foreground="blue")  # Màu chữ đỏ
 
    def send_email(self, to_address, subject, message):
        """Gửi email tới địa chỉ nhận"""
        from_address = "tranhuuhau2003@gmail.com"
        password = "frpq iken dpth pxku"
        
        # Khởi tạo email
        msg = MIMEMultipart()
        msg['From'] = from_address
        msg['To'] = to_address  
        msg['Subject'] = subject

        # Nội dung email
        msg.attach(MIMEText(message, 'plain'))

        # Cấu hình server SMTP để gửi email
        try:
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(from_address, password)
            text = msg.as_string()
            server.sendmail(from_address, to_address, text)
            server.quit()
             # Hiển thị thông báo thành công khi email được gửi
            messagebox.showinfo("Email Success", f"Email đã gửi thành công tới {to_address}")
            print(f"Email sent to {to_address}")
        except Exception as e:
            # Hiển thị thông báo lỗi nếu gửi thất bại
            messagebox.showerror("Email Error", f"Không thể gửi email tới {to_address}: {e}")
            print(f"Failed to send email to {to_address}: {e}")

    def send_warning_emails(self):
        """Kiểm tra và gửi cảnh báo học vụ cho sinh viên"""
        try:
            query = """
            SELECT s.mssv, s.ho_dem, s.ten, 
                a."Vắng có phép", a."Vắng không phép", 
                a."Tổng số tiết", a."Tỷ lệ vắng"
            FROM students s
            JOIN attendance a ON s.mssv = a.MSSV
            """
            self.cursor.execute(query)
            records = self.cursor.fetchall()

            for row in records:
                mssv, ho_dem, ten, vang_co_phep, vang_khong_phep, tong_so_tiet, ty_le_vang = row
                
                # Chuyển đổi tỷ lệ vắng sang kiểu số, thay thế ',' bằng '.'
                try:
                    ty_le_vang = float(ty_le_vang.replace(',', '.'))  # Thay thế ',' bằng '.'
                    print(f"Tỷ lệ vắng cho sinh viên {mssv}: {ty_le_vang}")  # In ra để kiểm tra
                except ValueError:
                    print(f"Lỗi chuyển đổi tỷ lệ vắng cho sinh viên {mssv}: {ty_le_vang}")
                    continue  # Bỏ qua sinh viên này nếu không thể chuyển đổi

                # Lấy email của sinh viên, phụ huynh, GVCN từ cơ sở dữ liệu
                student_email = self.get_student_email(mssv)
                parent_email = self.get_parent_email(mssv)
                teacher_email = self.get_teacher_email(mssv)
                tbm_email = self.get_tbm_email(mssv)

                # Kiểm tra và gửi cảnh báo theo tỷ lệ vắng
                if ty_le_vang >= 50:
                    # Gửi email cho sinh viên, phụ huynh, GVCN
                    subject = "Cảnh báo học vụ: Vắng học quá 50%"
                    message = f"Sinh viên {ho_dem} {ten} đã vắng hơn 50% số buổi học."
                    self.send_email(student_email, subject, message)
                    self.send_email(parent_email, subject, message)
                    self.send_email(teacher_email, subject, message)
                    self.send_email(tbm_email, subject, message)
                elif ty_le_vang >= 20:
                    # Gửi email chỉ cho sinh viên
                    subject = "Cảnh báo học vụ: Vắng học quá 20%"
                    message = f"Sinh viên {ho_dem} {ten} đã vắng hơn 20% số buổi học."
                    self.send_email(student_email, subject, message)

        except Exception as e:
            print(f"Lỗi khi gửi email cảnh báo: {e}")
       
    def get_student_email(self, mssv):
        """Lấy email sinh viên từ cơ sở dữ liệu dựa trên MSSV"""
        # Bạn cần có cột email trong bảng students hoặc một bảng riêng
        query = f"SELECT email_student FROM students WHERE mssv = '{mssv}'"
        self.cursor.execute(query)
        result = self.cursor.fetchone()
        return result[0] if result else None

    def get_parent_email(self, mssv):
        """Lấy email phụ huynh từ cơ sở dữ liệu dựa trên MSSV"""
        query = f"SELECT email_ph FROM parents WHERE mssv = '{mssv}'"
        self.cursor.execute(query)
        result = self.cursor.fetchone()
        return result[0] if result else None

    def get_teacher_email(self, mssv):
        """Lấy email GVCN từ cơ sở dữ liệu dựa trên MSSV"""
        query = f"SELECT email_gvcn FROM teachers WHERE mssv = '{mssv}'"
        self.cursor.execute(query)
        result = self.cursor.fetchone()
        return result[0] if result else None
    
    def get_tbm_email(self, mssv):
        """Lấy email TBM từ cơ sở dữ liệu dựa trên MSSV"""
        query = "SELECT email_tbm FROM tbm WHERE mssv = ?"
        self.cursor.execute(query, (mssv,))
        result = self.cursor.fetchone()
        return result[0] if result else None

    def view_details(self):
        selected_item = self.tree.selection()
        if selected_item:
            item_data = self.tree.item(selected_item, 'values')
            mssv = item_data[1]  # Lấy MSSV từ dữ liệu đã chọn
            
            # Truy vấn thông tin sinh viên
            query = '''
                SELECT s.ho_dem, s.ten, s.dot, s.ma_lop, s.ten_mon_hoc,
                    a."Vắng có phép", a."Vắng không phép",
                    a."11/06/2024", a."18/06/2024", a."25/06/2024", 
                    a."02/07/2024", a."09/07/2024", a."23/07/2024"
                FROM students s
                JOIN attendance a ON s.mssv = a.MSSV
                WHERE s.mssv = ?
            '''
            self.cursor.execute(query, (mssv,))
            details_data = self.cursor.fetchone()
            
            if details_data:
                # Tạo danh sách thời gian nghỉ
                time_off = []
                date_columns = ["11/06/2024", "18/06/2024", "25/06/2024", 
                                "02/07/2024", "09/07/2024", "23/07/2024"]
                # Kiểm tra độ dài của details_data
                for i, date in enumerate(date_columns, start=8):  # bắt đầu từ chỉ số 8
                    if i < len(details_data) and details_data[i] in ["K", "P"]:  # Kiểm tra độ dài trước
                        time_off.append(date)

                # Tạo chuỗi chi tiết thông tin sinh viên
                details = (
                    f"MSSV: {mssv}\n"
                    f"Họ tên: {details_data[0]} {details_data[1]}\n"
                    f"Đợt: {details_data[2]}\n"
                    f"Mã lớp: {details_data[3]}\n"
                    f"Tên môn học: {details_data[4]}\n"
                    f"Số ngày nghỉ có phép: {details_data[5]}\n"
                    f"Số ngày nghỉ không phép: {details_data[6]}\n"
                    f"Thời gian nghỉ: {', '.join(time_off) if time_off else 'Không có'}"
                )
                messagebox.showinfo("Chi tiết thông tin sinh viên", details)
            else:
                messagebox.showerror("Lỗi", "Không tìm thấy thông tin sinh viên.")
        else:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn một sinh viên để xem chi tiết.")
 
    # def generate_report_and_send_email(self):
    #     # Tạo kết nối SQLite mới trong hàm này
    #     conn = sqlite3.connect('students.db')
    #     cursor = conn.cursor()

    #     try:
    #         # Lấy dữ liệu sinh viên vắng nhiều
    #         query = """
    #         SELECT mssv, ho_dem, ten, gioi_tinh, ngay_sinh, vắng_có_phép, vắng_không_phép, tong_so_tiet, ty_le_vang
    #         FROM students
    #         WHERE ty_le_vang >= 20
    #         """
            
    #         # Thực hiện truy vấn
    #         cursor.execute(query)
    #         results = cursor.fetchall()
            
    #         # Chuyển đổi kết quả thành DataFrame
    #         columns = [column[0] for column in cursor.description]  # Lấy tên cột
    #         df = pd.DataFrame(results, columns=columns)

    #         # Lưu dữ liệu vào file Excel
    #         file_path = 'tong_hop_sinh_vien_vang_nhieu.xlsx'
    #         df.to_excel(file_path, index=False)

    #         # Gửi email với tệp đính kèm
    #         self.send_email_with_attachment(file_path)  # Đảm bảo rằng phương thức này có sẵn trong lớp của bạn

    #     except Exception as e:
    #         print(f"Có lỗi xảy ra trong quá trình tạo báo cáo và gửi email: {e}")

    #     finally:
    #         conn.close()  # Đảm bảo đóng kết nối

    def create_tonghop_table(self):
     
        # Tạo bảng mới với các cột cần thiết
        self.cursor.execute(""" 
        CREATE TABLE IF NOT EXISTS tonghop (
            mssv TEXT PRIMARY KEY,
            ho_dem TEXT,
            ten TEXT,
            ngay_sinh TEXT,
            dot TEXT,          -- Đợt
            ma_lop TEXT,       -- Mã lớp
            ten_mon_hoc TEXT,  -- Tên môn học
            vang_co_phep INTEGER,
            vang_khong_phep INTEGER
        )
        """)
    
        # Lưu thay đổi và đóng kết nối
        self.conn.commit()
        print("Bảng 'tonghop' đã được tạo thành công (nếu chưa tồn tại).")

    def clear_tonghop_table(self):
        # Xóa toàn bộ dữ liệu trong bảng 'tonghop'
        self.cursor.execute("DELETE FROM tonghop")
        
        # Lưu thay đổi và đóng kết nối
        self.conn.commit()
        print("Tất cả dữ liệu trong bảng 'tonghop' đã được xóa.")
        self.load_tonghop_to_treeview()
        
    def load_excel_tonghop(self):
        # tạo bảng tonghop nếu chưa tồn tại
        self.create_tonghop_table()
        # Sau đó, load dữ liệu vào Treeview (nếu cần thiết)
        self.load_tonghop_to_treeview()
        # Mở hộp thoại để chọn file
        filepath = filedialog.askopenfilename(
            title="Chọn file Excel",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )

        # Kiểm tra nếu không có file nào được chọn
        if not filepath:
            print("Không có file nào được chọn.")
            return
        
        try:
            df = pd.read_excel(filepath, engine='xlrd', header=None)  # Đọc file Excel
            df = df.fillna('')  # Thay thế các giá trị null bằng chuỗi rỗng

            # Lấy thông tin "Đợt", "Mã lớp học phần", "Tên môn học"
            dot = df.iloc[5, 2]  # Đợt nằm ở hàng 5, cột 2
            ma_lop = df.iloc[7, 2]  # Mã lớp học phần nằm ở hàng 7, cột 2
            ten_mon_hoc = df.iloc[8, 2]  # Tên môn học nằm ở hàng 8, cột 2

            # Lấy dữ liệu sinh viên từ hàng 13 trở đi
            df_sinh_vien = df.iloc[13:, [1, 2, 3, 4, 5, 24, 25]]  # Chỉ lấy các cột cần thiết
            df_sinh_vien.columns = ['MSSV', 'Họ đệm', 'Tên', 'Giới tính', 'Ngày sinh', 'Vắng có phép', 'Vắng không phép']
            
            self.mssv_list = df_sinh_vien['MSSV'].tolist()

       

            # Lưu dữ liệu vào SQLite (bảng tonghop)
            for _, row in df_sinh_vien.iterrows():
                self.cursor.execute("""
                    INSERT OR IGNORE INTO tonghop (mssv, ho_dem, ten, ngay_sinh, dot, ma_lop, ten_mon_hoc, vang_co_phep, vang_khong_phep)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    row['MSSV'],
                    row['Họ đệm'],
                    row['Tên'],
                    str(row['Ngày sinh']),
                    dot,           # Đợt
                    ma_lop,        # Mã lớp học phần
                    ten_mon_hoc,   # Tên môn học
                    int(row['Vắng có phép']),   # Vắng có phép
                    int(row['Vắng không phép']) # Vắng không phép
                ))
        
            # Lưu thay đổi vào cơ sở dữ liệu
            self.conn.commit()

            # Sau đó, load dữ liệu vào Treeview (nếu cần thiết)
            self.load_tonghop_to_treeview()

            print("Dữ liệu đã được lưu vào bảng 'tonghop' thành công!")
        except Exception as e:
            print(f"Lỗi khi lưu thông tin vắng: {e}")

    def load_tonghop_to_treeview(self):
        # Xóa tất cả các dòng hiện tại trong Treeview
        for item in self.tree.get_children():
            self.tree.delete(item)
        try:
            # Truy vấn dữ liệu từ bảng tonghop
            self.cursor.execute("SELECT mssv, ho_dem, ten, ngay_sinh, dot, ma_lop, ten_mon_hoc, vang_co_phep, vang_khong_phep FROM tonghop")
            rows = self.cursor.fetchall()

            # Duyệt qua từng dòng dữ liệu và thêm vào Treeview
            for index, row in enumerate(rows, start=1):  # Thêm STT bằng cách đếm index
                # Tính tổng vắng có phép và không phép
                tong_vang = row[7] + row[8]  # vang_co_phep + vang_khong_phep
                
                # Thêm dữ liệu vào Treeview
                self.tree.insert('', 'end', values=(index, row[0], row[1], row[2], row[3], row[4], row[5], row[6], tong_vang))

            print("Dữ liệu đã được tải lên Treeview thành công!")
        except Exception as e:
            print(f"Lỗi khi tải dữ liệu lên Treeview: {e}")
            
    def generate_report(self):
        # Tạo kết nối SQLite mới trong hàm này
        conn = sqlite3.connect('sinhvien.db')
        cursor = conn.cursor()

        try:
            # Lấy dữ liệu sinh viên có tổng số vắng >= 15
            query = """
            SELECT mssv, ho_dem, ten, ngay_sinh, dot, ma_lop, ten_mon_hoc, 
                vang_co_phep, vang_khong_phep, 
                (vang_co_phep + vang_khong_phep) AS tong_vang
            FROM tonghop
            WHERE (vang_co_phep + vang_khong_phep) >= 15
            """
            
            # Thực hiện truy vấn
            cursor.execute(query)
            results = cursor.fetchall()

            # Kiểm tra số lượng kết quả
            if not results:
                messagebox.showinfo("Không có dữ liệu", "Không có sinh viên nào có tổng số vắng >= 15.")
                return
            
            # Chuyển đổi kết quả thành DataFrame
            columns = [column[0] for column in cursor.description]  # Lấy tên cột
            df = pd.DataFrame(results, columns=columns)

            # Lưu dữ liệu vào file Excel trong thư mục hiện tại
            file_path = 'tong_hop_sinh_vien_vang_hon_15.xlsx'  # Thay đổi đường dẫn
            df.to_excel(file_path, index=False)

            # Gửi email với tệp đính kèm
            self.send_email_with_attachment(file_path)  # Đảm bảo rằng phương thức này có sẵn trong lớp của bạn
            
            # Hiển thị thông báo thành công sau khi tạo báo cáo và gửi email
            messagebox.showinfo("Thành công", "Báo cáo đã được tạo và gửi email thành công.")

        except Exception as e:
            print(f"Có lỗi xảy ra trong quá trình tạo báo cáo và gửi email: {e}")
            messagebox.showerror("Lỗi", f"Có lỗi xảy ra: {e}")

        finally:
            conn.close()  # Đảm bảo đóng kết nối

    def send_email_with_attachment(self, file_path):
        # Cấu hình thông tin email
        sender_email = "tranhuuhau2003@gmail.com"  # Địa chỉ email của bạn
        sender_password = "frpq iken dpth pxku"  # Mật khẩu email của bạn
        recipient_email = "tranhuuhauthh@gmail.com"  # Địa chỉ email của người nhận

        # Tạo một đối tượng MIMEMultipart
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = "Báo cáo sinh viên vắng nhiều"

        # Tạo phần thân email
        body = "Xin chào,\n\nĐây là báo cáo tổng hợp sinh viên vắng nhiều.\n\nTrân trọng!"
        msg.attach(MIMEText(body, 'plain'))

        # Đính kèm tệp Excel
        attachment = open(file_path, "rb")
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename= {file_path}')
        msg.attach(part)

        # Gửi email
        try:
            with smtplib.SMTP('smtp.gmail.com', 587) as server:  # Thay 'smtp.example.com' với máy chủ SMTP của bạn
                server.starttls()  # Bật chế độ bảo mật
                server.login(sender_email, sender_password)
                server.send_message(msg)
                print("Email đã được gửi thành công!")
                messagebox.showinfo("Email Success", f"Email đã gửi thành công tới {recipient_email}")  # Thông báo thành công
        except Exception as e:
            print(f"Có lỗi xảy ra khi gửi email: {e}")
            messagebox.showerror("Email Error", f"Có lỗi xảy ra khi gửi email: {e}")  # Thông báo lỗi
        finally:
            attachment.close()
   
    def start_scheduler(self):
        def run_scheduler():
            while True:
                now = datetime.now()
                # Kiểm tra xem có phải là ngày 24 và thời gian là 19:38 hay không
                if now.day == 24 and now.hour == 20 and now.minute == 5:
                    print("Đủ điều kiện gửi email. Gửi email...")
                    
                    # Đường dẫn tới tệp Excel cần đính kèm
                    file_path = 'tong_hop_sinh_vien_vang_nhieu.xlsx'
                    # Gọi hàm send_email_with_attachment với đường dẫn tệp
                    self.send_email_with_attachment(file_path)

                else:
                    print(f"Hiện tại là {now.strftime('%Y-%m-%d %H:%M:%S')} - Không đủ điều kiện để gửi email.")
                
                time.sleep(60)  # Kiểm tra mỗi 10 giây

        threading.Thread(target=run_scheduler, daemon=True).start()
        print("Scheduler đã được khởi động.")
              
    def main_menu(self):
        # Đặt kích thước và vị trí cửa sổ khi mở
        self.root.geometry("1400x750+70+15")  # Kích thước: 1000x700, vị trí: x=300, y=100

        self.frame = tk.Frame(self.root, bg="#f0f0f0", width=1000, height=700)
        self.frame.pack_propagate(False)
        self.frame.pack(fill='both', expand=True, padx=10, pady=10)

        # Các button nằm trên Treeview
        search_frame = tk.Frame(self.frame, bg="#f0f0f0")
        search_frame.pack(fill='x', pady=5)

        tk.Label(search_frame, text="Search by MSSV or Name:", font=("Arial", 8)).pack(side="left", padx=5)
        self.search_entry = tk.Entry(search_frame, font=("Arial", 8), width=15)
        self.search_entry.pack(side="left", padx=5)

        search_button = tk.Button(search_frame, text="Search", command=self.search_student, bg="#008CBA", fg="white", font=("Arial", 8), width=7)  # Giữ nút Search nhỏ hơn
        search_button.pack(side="left", padx=5)
        
        load_button = tk.Button(search_frame, text="Refresh", command=self.load_students_to_treeview, bg="#008CBA", fg="white", font=("Arial", 8), width=10)
        load_button.pack(side="left", padx=5)

        update_frame = tk.Frame(self.frame, bg="#f0f0f0")
        update_frame.pack(fill='x', pady=5)

        add_button = tk.Button(update_frame, text="Add Student", command=self.add_student, bg="#008CBA", fg="white", font=("Arial", 8), width=10)
        add_button.pack(side="left", padx=5)

        edit_button = tk.Button(update_frame, text="Edit Student", command=self.edit_student, bg="#008CBA", fg="white", font=("Arial", 8), width=10)
        edit_button.pack(side="left", padx=5)

        remove_button = tk.Button(update_frame, text="Remove Student", command=self.remove_student, bg="#008CBA", fg="white", font=("Arial", 8), width=10)
        remove_button.pack(side="left", padx=5)

        viewDetails_button = tk.Button(update_frame, text="View Details", command=self.view_details, bg="#008CBA", fg="white", font=("Arial", 8), width=10)
        viewDetails_button.pack(side="left", padx=5)
        
        sort_button = tk.Button(update_frame, text="Sort Student", command=self.sort_students_by_absences, bg="#008CBA", fg="white", font=("Arial", 8), width=10)
        sort_button.pack(side="left", padx=5)

        sendMail_button = tk.Button(update_frame, text="Send Mail", command=self.send_warning_emails, bg="#008CBA", fg="white", font=("Arial", 8), width=10)
        sendMail_button.pack(side="left", padx=5)

        summary_button = tk.Button(update_frame, text="Summary", command=self.load_excel_tonghop, bg="#008CBA", fg="white", font=("Arial", 8), width=10)
        summary_button.pack(side="left", padx=5)

        sendReport_button = tk.Button(update_frame, text="Send Report", command=self.generate_report, bg="#008CBA", fg="white", font=("Arial", 8), width=10)
        sendReport_button.pack(side="left", padx=5)
        
        clear_sumary = tk.Button(update_frame, text="Clear Summary", command=self.clear_tonghop_table, bg="#008CBA", fg="white", font=("Arial", 8), width=10)
        clear_sumary.pack(side="left", padx=5)

        

        # Treeview
        self.tree_frame = tk.Frame(self.frame)
        self.tree_frame.pack(fill='both', expand=True)

        self.tree = ttk.Treeview(self.tree_frame, columns=['STT', 'MSSV', 'Họ đệm', 'Tên', 'Ngày Sinh', 'Đợt', 'Mã lớp', 'Tên môn học', 'Tổng vắng'], show="headings")
        self.tree.pack(side='left', fill='both', expand=True)

        self.tree.column('STT', width=30, anchor='center')
        self.tree.column('MSSV', width=70, anchor='center')
        self.tree.column('Họ đệm', width=100, anchor='w')
        self.tree.column('Tên', width=70, anchor='w')
        self.tree.column('Ngày Sinh', width=70, anchor='w')
        self.tree.column('Đợt', width=70, anchor='center')
        self.tree.column('Mã lớp', width=70, anchor='center')
        self.tree.column('Tên môn học', width=100, anchor='w')
        self.tree.column('Tổng vắng', width=70, anchor='center')

        for col in self.tree["columns"]:
            self.tree.heading(col, text=col, command=lambda _col=col: self.sort_column(_col, False))

        self.scrollbar = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        self.scrollbar.pack(side='right', fill='y')
        self.tree.configure(yscrollcommand=self.scrollbar.set)

        # Nút chọn file nằm dưới Treeview
        self.select_file_button = tk.Button(self.frame, text="Choose File", command=self.select_file, bg="#008CBA", fg="white", font=("Arial", 10))
        self.select_file_button.pack(pady=10)


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelApp(root)
    root.mainloop()
