import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from fpdf import FPDF
from datetime import datetime
from PyPDF2 import PdfMerger
import sqlite3
import glob
import os
import re
import getpass


class CustomDialog(simpledialog.Dialog):
    def __init__(self, parent, title, prompt, ok_text="Add", cancel_text="Finish"):
        self.prompt = prompt
        self.ok_text = ok_text
        self.cancel_text = cancel_text
        self.should_continue = True
        super().__init__(parent, title=title)

    def body(self, master):
        tk.Label(master, text=self.prompt).pack()
        self.entry = tk.Entry(master)
        self.entry.pack()
        self.entry.focus_set()
        self.grab_set()
        return self.entry

    def buttonbox(self):
        box = tk.Frame(self)
        w = tk.Button(box, text=self.ok_text, width=10, command=self.add_entry, default=tk.ACTIVE)
        w.pack(side=tk.LEFT, padx=5, pady=5)
        w = tk.Button(box, text=self.cancel_text, width=10, command=self.finish_entry)
        w.pack(side=tk.RIGHT, padx=5, pady=5)
        self.bind("<Return>", self.add_entry)
        self.bind("<Escape>", self.finish_entry)
        box.pack()

    def add_entry(self, event=None):
        result = self.entry.get().strip()
        self.result = result
        self.should_continue = True
        self.ok()

    def finish_entry(self, event=None):
        result = self.entry.get().strip()
        if result != "":
            self.result = result
        self.should_continue = False
        self.destroy()

    def destroy(self):
        self.grab_release()  # 释放焦点锁定
        super().destroy()

class MyCustomError(Exception):
    pass

class Ui_MainWindow:
    DATABASE_PATH = r'\\pcba1\illumina\EDHR\DATABASE\EDHR.db'
    TAR_PART = r'\\172.18.8.10\L01'
    PDFSAVE_PATH = r'\\172.18.8.10\labelprinting\EDHR_PDF'
    TEMP_PATH = r'\\pcba1\illumina\EDHR\temp'
    try:
        with sqlite3.connect(DATABASE_PATH) as conn:
            pass
    except sqlite3.Error:
        DATABASE_PATH = r'D:\EDHR-COP\database\EDHRPDF.db'
        TAR_PART = r'D:\\EDHR-COP\\temp\\TAR\\'
        PDFSAVE_PATH = r'D:\\EDHR-COP\\temp\\EDHR_PDF'
        TEMP_PATH = r'D:\\EDHR-COP\\temp\\TempPdf'

    def __init__(self, root, badge_number):
        self.root = root
        self.root.title("PDF Converter")
        self.center_window(305, 220)

        self.label = ttk.Label(root, text="Serial Number:")
        self.label.place(x=30, y=10)

        self.label = ttk.Label(root, text=r'Stations: '+self.get_computer_name()+ r' -----> '+ name)
        self.label.place(x=30, y=200)

        self.combo = ttk.Combobox(root, values=self.get_options_from_database(), width=18, height=5)
        self.combo.place(x=125, y=10)
        
        self.button_enabled = True
        self.button = ttk.Button(root, text="Convert", command=self.convertClicked)
        self.button.place(x=20, y=50)
        self.root.bind("<Return>", lambda event: self.convertClicked() if self.button_enabled else None)

        self.button1 = ttk.Button(root, text="M_Merge", command=self.M_MergeClicked)
        self.button1.place(x=115, y=50)

        self.button2 = ttk.Button(root, text="PDF_Merge", command=self.mergeClicked)
        self.button2.place(x=210, y=50)

        self.model = tk.StringVar()
        self.text = tk.Text(root, width=35, height=4, font=('Arial', 10), padx=10, pady=10, spacing1=5, spacing2=5)
        self.text.place(x=20, y=90)

        
        self.update_button_state()
        self.computer_name = self.get_computer_name()
        self.last_arrow_right_press_time = datetime.now()
        self.all_merged_pdf_paths = []
        self.serial_number = None
        self.serial_number_M = None
        self.Topserial = None
        self.data_M = None

        def on_arrow_right_double_press(event):
            current_time = datetime.now()
            time_difference = current_time - self.last_arrow_right_press_time
            if time_difference.total_seconds() < 1:  
                self.onepdfconr()

            self.last_arrow_right_press_time = current_time

        root.bind("<Right>", on_arrow_right_double_press)


    def update_button_state(self):
        try:
            with sqlite3.connect(self.DATABASE_PATH) as conn:
                cursor = conn.cursor()
                cursor.execute('SELECT convertor, M_Merge, PDF_Merge FROM Settings WHERE id = 1')
                settings_data = cursor.fetchone()
                if settings_data:
                    Convert_enabled, m_merge_enabled, pdf_merge_enabled, = settings_data
                    self.button['state'] = 'normal' if Convert_enabled else 'disabled'
                    self.button1['state'] = 'normal' if m_merge_enabled else 'disabled'
                    self.button2['state'] = 'normal' if pdf_merge_enabled else 'disabled'

        except sqlite3.Error as e:
            print("SQLite error:", e)

    def get_computer_name(self):
        try:
            computer_name = getpass.getuser()
            return computer_name
        except Exception as e:
            print(f"Error getting computer name: {e}")
            return None


    def center_window(self, width, height):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        x_coordinate = (screen_width - width) // 2
        y_coordinate = (screen_height - height) // 2

        self.root.geometry(f"{width}x{height}+{x_coordinate}+{y_coordinate}")

    def get_options_from_database(self):
        try:
            with sqlite3.connect(self.DATABASE_PATH) as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT Fodelname FROM SaveFodel")
                options = [row[0] for row in cursor.fetchall()]
                return options
        except sqlite3.Error as e:
            print("SQLite error:", e)
            return []

    #以下为新加代码
    def check_and_convert(self, file_paths):
        pattern = re.compile(r'[a-zA-Z]{3}\d{7}|[a-zA-Z]{6}\d{7}')
        for file_path in file_paths:
            file_name = os.path.basename(file_path)
            match = re.search(pattern, file_name)
            if not match:
                messagebox.showerror('Error', '你选的什么鬼！！！.')
                return
            serial_number = match.group()
            with open(file_path, 'r') as tar_file:
                lines = tar_file.readlines()
            test_name = None
            for line in lines:
                stripped_line = line.strip()
                if stripped_line.startswith('P'):
                    test_name = stripped_line.replace('P', '', 1)
            print(test_name)
            test_result = any('TP' in line for line in lines)
            print(test_result)
            if test_result:
                pdf = FPDF()
                pdf.add_page()
                pdf.set_font("Arial", size=10)
                line_height = 6
                for line in lines:
                    pdf.multi_cell(0, line_height, txt=line.strip())
                # 生成PDF文件的路径
                date = datetime.now().strftime("%Y%m%d_%H%M%S")
                save_name = f"{serial_number}_{test_name}_{date}.pdf"
                s = f"Normal select"
                out_path = os.path.join(self.TEMP_PATH)
                normal_path = os.path.normpath(out_path)
                if not os.path.exists(out_path):
                    os.makedirs(out_path)
                pdf_file_path = os.path.join(out_path, save_name)
                # 保存PDF文件
                pdf.output(pdf_file_path)
                with sqlite3.connect(self.DATABASE_PATH) as conn:
                    cursor = conn.cursor()
                    cursor.execute('''INSERT INTO SaveFilelist (Model, PartNumber, SerialNumber, PartName, Savename, date, PC_Name, name)
                                      VALUES (?, ?, ?, ?, ?, ?, ?, ?)''',
                                   (s, s, serial_number, test_name, save_name, date, self.computer_name, name))
                    conn.commit()
                
                self.show_message(f"转换成功:{serial_number}\n文件保存到：\n{normal_path}", "#00F5FF")
            else:
                self.show_message(f"你为啥要选个Fail的\n{test_name}\n的文件", "#EE2C2C")


    def onepdfconr(self):
        file_paths = filedialog.askopenfilenames(filetypes=[("TAR Files", "*.tar")])
        self.check_and_convert(file_paths)
    # 新加以上代码用于单个转换

    def M_MergeClicked(self):
        self.all_merged_pdf_paths = []
        if hasattr(self, 'dialog') and self.dialog:
            self.dialog.destroy()
        Topfolder = self.combo.get()
        if not Topfolder:
            self.show_message("Please select a save folder", "yellow")
            return

        self.Topserial = simpledialog.askstring("Input Top serial number ", "Enter Top serial number:")
        if not self.Topserial:
            self.show_message("Please Enter Top serial number", "yellow")
            return

        serial_numbers = []
        i = 1
        while True:
            prompt = f"Enter Sub serial number {i} : "
            dialog = CustomDialog(self.root, "Enter Sub Serial Number", prompt)

            if not dialog.should_continue:
                if dialog.result is not None:
                    serial_numbers.append(dialog.result)
                break

            result = dialog.result
            if result is not None:
                if not result:
                    messagebox.showwarning('Warning', 'Please enter Serial number.')
                    return

                letters = ''.join(filter(str.isalpha, result))
                if not 3 <= len(letters) <= 6:
                    messagebox.showerror('Error', 'Please enter a valid serial number.')
                    return

                if (len(letters) == 3 and len(result) != 10) or (
                        len(letters) == 6 and len(result) != 13):
                    messagebox.showerror('Error', 'Please enter a valid serial number.')
                    return
                serial_numbers.append(result)
            i += 1

        print("serial numbers list:", serial_numbers)

        for self.serial_number_M in serial_numbers:
            letters = ''.join(filter(str.isalpha, self.serial_number_M))
            try:
                with sqlite3.connect(self.DATABASE_PATH) as conn:
                    cursor = conn.cursor()
                    cursor.execute('SELECT * FROM PDF WHERE Flag LIKE ?', (f'{letters}%',))
                    self.data_M = cursor.fetchone()

                if self.data_M:
                    print(f'Data: {self.data_M}')
                    success, result_message = self.takefoler_M(self.data_M)
                    if not success:
                        print(f'Stopping the loop due to an error in takefoler_M: {result_message}')
                        break
                else:
                    messagebox.showwarning('Warning', 'Part no found. Contact the administrator to add.')

            except sqlite3.Error as e:
                print("SQLite error:", e)
                messagebox.showerror('Error', 'Failed to fetch data: ' + str(e))
                break
        self.generate_output_path()

    def takefoler_M(self, data_M):
        try:
            non_empty_data = data_M[5:]
            non_empty_columns = [value for value in non_empty_data if value]
            pdffolder = [os.path.join(self.TAR_PART, column_value, 'archive') for column_value in non_empty_columns]
            print(f'pdffodel:', pdffolder)
            if pdffolder:
                pdf_merger = self.create_pdf_merger()

                for folder in pdffolder:
                    tar_files_M = self.get_tar_files_M(folder)
                    print(f'tarfile:', tar_files_M)

                    if not tar_files_M:
                        self.show_message(f'{self.serial_number_M}_{os.path.basename(os.path.dirname(folder))} No Test', "red")
                        return False, "No Test"  # 返回一个元组

                    latest_tar_file = self.get_latest_tar_file(tar_files_M)
                    test_result = self.check_sixth_line(latest_tar_file)
                    if not test_result:
                        self.show_message(f'{self.serial_number_M}_{os.path.basename(os.path.dirname(folder))} test Failed', "red")
                        return False, "Test Failed"  # 返回一个元组

                    if not os.path.exists(self.TEMP_PATH):
                        os.makedirs(self.TEMP_PATH)

                    temp_pdf_path = os.path.join(self.TEMP_PATH, f"{self.serial_number_M}_{os.path.basename(os.path.dirname(folder))}.pdf")
                    print(f'tempfile:', temp_pdf_path)

                    pdf = self.create_pdf_from_tar_M(latest_tar_file)
                    pdf.output(temp_pdf_path)
                    pdf_merger.append(temp_pdf_path)

                self.saveMergedPdf_M(pdf_merger)
                return True, "Success"  # 返回一个元组
            else:
                self.showDataNotFoundWarning()
                return False, "Data Not Found"  # 返回一个元组
        except Exception as e:
            print("Error in takefoler_M:", e)
            messagebox.showerror('Error', 'An error occurred in takefoler_M: ' + str(e))
            raise MyCustomError("Error in takefoler_M")

    def saveMergedPdf_M(self, pdf_merger):
        datetime_str = datetime.now().strftime("%y%m%d_%H")
        model = self.data_M[3]
        part_name = self.data_M[2]
        savename = f"{self.serial_number_M}_{part_name}_{datetime_str}.pdf"
        output_folder = os.path.join(self.PDFSAVE_PATH, model, part_name)
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        merged_pdf_path = os.path.join(output_folder, savename)
        pdf_merger.write(merged_pdf_path)
        pdf_merger.close()

        self.all_merged_pdf_paths.append(merged_pdf_path)
        return merged_pdf_path

    def get_tar_files_M(self, folder):
        tar_files_M = glob.glob(os.path.join(folder, f'*{self.serial_number_M}*'))
        return tar_files_M

    def create_pdf_from_tar_M(self, tar_file_path):
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=10)

        with open(tar_file_path, 'r') as tar_file:
            lines = tar_file.readlines()
            for line in lines:
                pdf.multi_cell(0, 6, txt=line.strip())

        return pdf

    def generate_output_path(self):
        TopSerial = self.Topserial
        Topfolder = self.combo.get()

        date = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_folder = os.path.join(self.PDFSAVE_PATH, Topfolder)

        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        output_filename = os.path.join(output_folder, f"{TopSerial}-{Topfolder}-{date}.pdf")
        print(f'lastfodel:', output_filename)

        with open(output_filename, "wb") as output_file:
            if not self.all_merged_pdf_paths:
                self.show_message("No input sub serial number.", "yellow")
                return
            pdf_merger = self.create_pdf_merger()
            print(f'allpdf:', self.all_merged_pdf_paths)
            for merged_path in self.all_merged_pdf_paths:
                pdf_merger.append(merged_path)
            pdf_merger.write(output_file)
            pdf_merger.close()

        self.show_message(f"PDF merged successfully.\nOutput:{TopSerial}-{Topfolder}-{date}.pdf", "green")

        return output_filename


    def mergeClicked(self):
        Topfolder = self.combo.get()
        if not Topfolder:
            self.show_message("Please select a save folder", "yellow")
            return
        serialA = simpledialog.askstring("Input serial number ", "Enter serial number:")
        if not serialA:
            self.show_message("Please Enter serial number", "yellow")
            return

        folder_path = filedialog.askdirectory(title="Select Folder")
        if not folder_path:
            self.show_message("Please select a file folder", "yellow")
            return

        pdf_merger = PdfMerger()

        pdf_files = [file for file in os.listdir(folder_path) if file.lower().endswith(".pdf")]

        if not pdf_files:
            self.show_message("No PDF files found in the selected folder", "yellow")
            return

        pdf_files.sort()

        for pdf_file in pdf_files:
            file_path = os.path.join(folder_path, pdf_file)
            pdf_merger.append(file_path)

        date = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = os.path.join(self.PDFSAVE_PATH, Topfolder)
        if not os.path.exists(output_file):
            os.makedirs(output_file)

        output_filename = os.path.join(output_file, f"{serialA}-{Topfolder}-{date}.pdf")

        with open(output_filename, "wb") as output_file:
            pdf_merger.write(output_file)

        self.show_message(f"PDF merged successfully.\nOutput:{serialA}-{Topfolder}-{date}.pdf", "green")

    def show_message(self, message, color):
        self.model.set(message)
        self.text.delete(1.0, tk.END)  # 清空Text组件内容
        self.text.insert(tk.END, message)  # 插入新的文本
        self.text.configure(bg=color)

    def convertClicked(self):
        self.model.set("")  # 清空列表

        self.serial_number = self.combo.get()  # 更新成员变量
        if not self.serial_number:
            messagebox.showwarning('Warning', 'Please enter Serial number.')
            return

        letters = ''.join(filter(str.isalpha, self.serial_number))
        if not 3 <= len(letters) <= 6:
            messagebox.showerror('Error', 'Please enter a valid serial number.')
            return

        if (len(letters) == 3 and len(self.serial_number) != 10) or (len(letters) == 6 and len(self.serial_number) != 13):
            messagebox.showerror('Error', 'Please enter a valid serial number.')
            return

        try:
            with sqlite3.connect(self.DATABASE_PATH) as conn:
                cursor = conn.cursor()
                cursor.execute('SELECT * FROM PDF WHERE Flag LIKE ?', (f'{letters}%',))
                data = cursor.fetchone()

            if data:
                self.take_folder(data)
            else:
                messagebox.showwarning('Warning', 'Part no found. Contact the administrator to add.')

        except sqlite3.Error as e:
            print("SQLite error:", e)
            messagebox.showerror('Error', 'Failed to fetch data: ' + str(e))

    def take_folder(self, data):
        non_empty_data = data[5:]
        non_empty_columns = [value for value in non_empty_data if value]
        pdffolder = [os.path.join(self.TAR_PART, column_value, 'archive') for column_value in non_empty_columns]

        if pdffolder:
            pdf_merger = self.create_pdf_merger()

            for folder in pdffolder:
                tar_files = self.get_tar_files(folder)
                if not tar_files:
                    self.show_message(f'{self.serial_number}_{os.path.basename(os.path.dirname(folder))} No Test', "red")
                    return

                latest_tar_file = self.get_latest_tar_file(tar_files)
                test_result = self.check_sixth_line(latest_tar_file)
                if not test_result:
                    self.show_message(f'{self.serial_number}_{os.path.basename(os.path.dirname(folder))} test Failed', "red")
                    return

                if not os.path.exists(self.TEMP_PATH):
                    os.makedirs(self.TEMP_PATH)

                temp_pdf_path = os.path.join(self.TEMP_PATH, f"{self.serial_number}_{os.path.basename(os.path.dirname(folder))}.pdf")
                pdf = self.create_pdf_from_tar(latest_tar_file)
                
                pdf.output(temp_pdf_path)
                pdf_merger.append(temp_pdf_path)

            self.save_merged_pdf(pdf_merger, data)
        else:
            messagebox.showwarning('Warning', 'Contact the administrator to add test folder.')

    def create_pdf_merger(self):
        return PdfMerger()

    def get_tar_files(self, folder):
        tar_files = glob.glob(os.path.join(folder, f'*{self.serial_number}*'))
        return tar_files

    def get_latest_tar_file(self, tar_files):
        tar_files.sort(key=os.path.getmtime, reverse=True)
        return tar_files[0]

    def check_sixth_line(self, tar_file_path):
        with open(tar_file_path, 'r') as tar_file:
            lines = tar_file.readlines()
            sixth_line = lines[5].strip()
            if "TP" not in sixth_line:
                return False
            else:
                return True

    def create_pdf_from_tar(self, tar_file_path):
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=10)

        with open(tar_file_path, 'r') as tar_file:
            lines = tar_file.readlines()
            for line in lines:
                pdf.multi_cell(0, 6, txt=line.strip())

        return pdf

    def save_merged_pdf(self, pdf_merger, data):
        try:
            date = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
            datetime_str = datetime.now().strftime("%y%m%d_%H")
            model = data[3]
            part_name = data[2]
            PartNumber = data[1]
            savename = f"{self.serial_number}_{part_name}_{datetime_str}.pdf"
            output_folder = os.path.join(self.PDFSAVE_PATH, model, part_name)

            if not os.path.exists(output_folder):
                os.makedirs(output_folder)

            merged_pdf_path = os.path.join(output_folder, savename)
            pdf_merger.write(merged_pdf_path)
            pdf_merger.close()

            with sqlite3.connect(self.DATABASE_PATH) as conn:
                cursor = conn.cursor()
                cursor.execute('''INSERT INTO SaveFilelist (Model, PartNumber, SerialNumber, PartName, Savename, date, PC_Name, name)
                                  VALUES (?, ?, ?, ?, ?, ?, ?, ?)''',
                               (model, PartNumber, self.serial_number, part_name, savename, date, self.computer_name, name))
                conn.commit()

            self.show_message(f"PDF Files Convert Successfully:\n{savename}", "green")

        except Exception as e:
            print("Error:", e)
            messagebox.showerror('Error', 'An error occurred while saving merged PDF: ' + str(e))

    def showDataNotFoundWarning(self):
        messagebox.showerror('Error', 'Contact the administrator to add test folder.')

def initialize_ui(employee_info):
    root = tk.Tk()
    ui = Ui_MainWindow(root, employee_info)
    root.mainloop()

def get_employee_info():
    while True:
        badge_number = simpledialog.askstring("Employee Badge Number", "Enter 4-digit badge number:\n请输入4位数工牌号码：")
        if badge_number is None:
            sys.exit()
        elif not badge_number.isdigit() or len(badge_number) != 4:
            messagebox.showwarning('Warning', 'Please enter a valid 4-digit number.')
        else:          
            employee_info = get_employee_by_badge_number(badge_number)            
            if employee_info:
                return employee_info
            else:
                messagebox.showwarning('Warning', 'Invalid badge number. Please try again.')

def get_employee_by_badge_number(badge_number):
    try:
        db_path = r'\\pcba1\illumina\EDHR\DATABASE\EDHR.db'
        with sqlite3.connect(db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT name, badge_number FROM employees WHERE badge_number = ?', (badge_number,))
            employee_info = cursor.fetchone()
            conn.commit()
        return employee_info
    except sqlite3.Error as e:
        messagebox.showwarning('Warning', 'Please enter a valid 4-digit number.')
        get_employee_info()
        return get_employee_info()

employee_info = get_employee_info()

if employee_info:
    name, badge_number = employee_info
else:
    print()
if __name__ == "__main__":
    initialize_ui(employee_info)
    

