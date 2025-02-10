import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from dbfread import DBF
import os
import csv
import re
import chardet

class ModernButton(ttk.Button):
    def __init__(self, master=None, **kwargs):
        super().__init__(master, **kwargs)
        self.configure(style='Modern.TButton')

class DBFConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("DBF to TXT Converter")
        self.root.geometry("800x600")
        self.root.configure(bg='white')
        
        # Define styles and colors
        self.setup_styles()
        
        # Create main container
        main_frame = ttk.Frame(root, style='Main.TFrame')
        main_frame.pack(fill='both', expand=True, padx=30, pady=20)
        
        # Header
        header_frame = ttk.Frame(main_frame, style='Main.TFrame')
        header_frame.pack(fill='x', pady=(0, 20))
        
        title_label = ttk.Label(
            header_frame,
            text="DBF to TXT Converter",
            style='Header.TLabel'
        )
        title_label.pack()
        
        # File selection area
        file_frame = ttk.Frame(main_frame, style='Card.TFrame')
        file_frame.pack(fill='x', pady=10, ipady=20)
        
        self.file_label = ttk.Label(
            file_frame,
            text="ไม่ได้เลือกไฟล์",
            style='FileLabel.TLabel'
        )
        self.file_label.pack(pady=(10, 15))
        
        select_btn = ModernButton(
            file_frame,
            text="เลือกไฟล์ DBF",
            command=self.select_file
        )
        select_btn.pack()
        
        # Progress area
        progress_frame = ttk.Frame(main_frame, style='Card.TFrame')
        progress_frame.pack(fill='x', pady=20, ipady=20)
        
        self.status_label = ttk.Label(
            progress_frame,
            text="พร้อมแปลงไฟล์",
            style='Status.TLabel'
        )
        self.status_label.pack(pady=(0, 10))
        
        self.progress = ttk.Progressbar(
            progress_frame,
            orient="horizontal",
            length=500,
            mode="determinate",
            style='Modern.Horizontal.TProgressbar'
        )
        self.progress.pack(pady=10, padx=30, fill="x")
        
        # Convert button
        self.convert_btn = ModernButton(
            main_frame,
            text="แปลงไฟล์",
            command=self.select_file
        )
        self.convert_btn.pack(pady=20)

    def setup_styles(self):
        style = ttk.Style()
        
        # Configure colors (Black & White)
        style.configure('Main.TFrame', background='white')
        style.configure('Card.TFrame',
            background='white',
            relief='solid',
            borderwidth=1,
            bordercolor='black'
        )
        
        # Configure labels
        style.configure('Header.TLabel',
            font=('Tahoma', 24, 'bold'),
            background='white',
            foreground='black'
        )
        
        style.configure('FileLabel.TLabel',
            font=('Tahoma', 12),
            background='white',
            foreground='black'
        )
        
        style.configure('Status.TLabel',
            font=('Tahoma', 11),
            background='white',
            foreground='black'
        )
        
        # Configure buttons (Change foreground to black)
        style.configure('Modern.TButton',
            font=('Tahoma', 11),
            padding=10,
            background='black',
            foreground='black'  # เปลี่ยนเป็นสีดำ
        )
        
        # Configure progress bar
        style.configure('Modern.Horizontal.TProgressbar',
            troughcolor='white',
            background='black',
            thickness=10
        )

    def detect_encoding(self, dbf_file):
        with open(dbf_file, 'rb') as f:
            raw_data = f.read(10000)
            result = chardet.detect(raw_data)
        return result['encoding']
    
    def select_file(self):
        dbf_file = filedialog.askopenfilename(
            filetypes=[("DBF files", "*.dbf"), ("All files", "*.*")],
            title="เลือกไฟล์ DBF"
        )
        if not dbf_file:
            return
        self.file_label.config(text=f"ไฟล์ที่เลือก: {os.path.basename(dbf_file)}")
        self.convert_file(dbf_file)
    
    def convert_file(self, dbf_file):
        try:
            self.status_label.config(text="กำลังแปลงไฟล์...", foreground='black')
            self.progress["value"] = 0
            self.root.update_idletasks()
            
            encoding = self.detect_encoding(dbf_file)
            try:
                table = DBF(dbf_file, encoding=encoding)
            except UnicodeDecodeError:
                table = DBF(dbf_file, encoding='cp1252')
            
            output_file = os.path.splitext(dbf_file)[0] + ".txt"
            
            with open(output_file, 'w', encoding='utf-8', newline='') as f:
                writer = csv.writer(f, delimiter='|')
                field_names = [name.replace("DCIPEscreen", "DCIP/E_screen") for name in table.field_names]
                writer.writerow(field_names)
                
                total_records = len(table)
                
                for i, record in enumerate(table):
                    row = [self.clean_value(v) for v in record.values()]
                    writer.writerow(row)
                    if i % 10 == 0:
                        self.progress["value"] = min(100, (i / total_records) * 100)
                        self.status_label.config(
                            text=f"กำลังแปลงไฟล์... {int(self.progress['value'])}%",
                            foreground='black'
                        )
                        self.root.update_idletasks()
            
            self.progress["value"] = 100
            self.status_label.config(
                text=f"แปลงไฟล์เสร็จสิ้น: {os.path.basename(output_file)}",
                foreground='green'
            )
            
            if messagebox.askyesno("เสร็จสิ้น", "ต้องการเปิดไฟล์หรือไม่?"):
                os.startfile(output_file)
                
        except Exception as e:
            messagebox.showerror("Error", f"เกิดข้อผิดพลาด: {str(e)}")
            self.status_label.config(
                text="เกิดข้อผิดพลาดในการแปลงไฟล์",
                foreground='red'
            )
            self.progress["value"] = 0

    def clean_value(self, value):
        if value is None:
            return ''
        if isinstance(value, (int, float)):
            return str(value)
        
        value = str(value).strip()
        value = value.replace("DCIPEscreen", "DCIP/E_screen")
        return re.sub(r'[^0-9a-zA-Z./_]', '', value)

if __name__ == "__main__":
    root = tk.Tk()
    app = DBFConverterApp(root)
    root.mainloop()
