import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from dbfread import DBF
import os
import csv
from datetime import datetime
import re

class DBFConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("DBF to TXT Converter")
        self.root.geometry("800x600")
        
        # กำหนดธีมสี
        self.COLORS = {
            'primary': '#2563eb',        
            'secondary': '#3b82f6',      
            'background': '#f8fafc',     
            'surface': '#ffffff',        
            'success': '#22c55e',        
            'warning': '#f59e0b',       
            'error': '#ef4444',          
            'text': '#000000',           
            'text_secondary': '#000000'   
        }

        # ตั้งค่าพื้นหลังหลัก
        self.root.configure(bg=self.COLORS['background'])

        # กำหนดสไตล์
        self.style = ttk.Style()
        self.style.configure("Main.TFrame", background=self.COLORS['background'])
        self.style.configure("Card.TFrame", background=self.COLORS['surface'])
        
        # สไตล์สำหรับปุ่ม
        self.style.configure(
            "Primary.TButton",
            background=self.COLORS['primary'],
            foreground='black',  
            font=('Tahoma', 12, 'bold'),
            padding=10
        )
        
        # สไตล์สำหรับ Label
        self.style.configure(
            "Header.TLabel",
            background=self.COLORS['background'],
            foreground='black',  
            font=('Tahoma', 24, 'bold')
        )
        
        self.style.configure(
            "SubHeader.TLabel",
            background=self.COLORS['surface'],
            foreground='black',  
            font=('Tahoma', 12)
        )

        # สร้าง Main Container
        self.main_container = ttk.Frame(root, style="Main.TFrame")
        self.main_container.pack(pady=30, padx=40, fill="both", expand=True)

        # Header Section
        header = ttk.Label(
            self.main_container,
            text="DBF to TXT Converter",
            style="Header.TLabel"
        )
        header.pack(pady=(0, 30))

        # Card Container
        self.card = ttk.Frame(self.main_container, style="Card.TFrame")
        self.card.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Add shadow effect to card (using canvas)
        canvas = tk.Canvas(
            self.card,
            background=self.COLORS['surface'],
            highlightthickness=0,
            height=400
        )
        canvas.pack(fill="both", expand=True, padx=2, pady=2)

        # File Selection Area
        self.file_frame = ttk.Frame(canvas, style="Card.TFrame")
        self.file_frame.pack(fill="x", padx=30, pady=20)

        self.file_label = ttk.Label(
            self.file_frame,
            text="ไม่ได้เลือกไฟล์",
            style="SubHeader.TLabel"
        )
        self.file_label.pack(pady=10)

        select_btn = ttk.Button(
            self.file_frame,
            text="เลือกไฟล์ DBF",
            style="Primary.TButton",
            command=self.select_file
        )
        select_btn.pack(pady=10)

        # Progress Area
        self.progress_frame = ttk.Frame(canvas, style="Card.TFrame")
        self.progress_frame.pack(fill="x", padx=30, pady=20)

        self.progress = ttk.Progressbar(
            self.progress_frame,
            orient="horizontal",
            length=500,
            mode="determinate",
            style="Progress.Horizontal.TProgressbar"
        )
        self.progress.pack(pady=15, fill="x")

        self.status_label = ttk.Label(
            self.progress_frame,
            text="พร้อมแปลงไฟล์",
            style="SubHeader.TLabel"
        )
        self.status_label.pack()

        # Convert Button
        self.convert_btn = ttk.Button(
            canvas,
            text="แปลงไฟล์",
            style="Primary.TButton",
            command=self.select_file
        )
        self.convert_btn.pack(pady=20)

        # Footer
        footer_frame = ttk.Frame(self.main_container, style="Main.TFrame")
        footer_frame.pack(fill="x", pady=20)
        
        version_label = ttk.Label(
            footer_frame,
            text="Version 1.1.0",
            foreground='black',  
            font=('Tahoma', 10)
        )
        version_label.pack(side="right")

        # สร้างสไตล์สำหรับ Progressbar
        self.style.configure(
            "Progress.Horizontal.TProgressbar",
            troughcolor=self.COLORS['background'],
            background=self.COLORS['primary'],
            thickness=10
        )

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
            self.status_label.config(text="กำลังแปลงไฟล์...", foreground='black')  # เปลี่ยนเป็นสีดำ
            self.progress["value"] = 0
            self.root.update_idletasks()
            
            table = DBF(dbf_file, encoding='utf-8')
            output_file = os.path.splitext(dbf_file)[0] + ".txt"
            
            with open(output_file, 'w', encoding='utf-8', newline='') as f:
                writer = csv.writer(f, delimiter='|')
                writer.writerow(table.field_names)
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
                foreground='Green'  
            )
            if messagebox.askyesno("เสร็จสิ้น", "ต้องการเปิดไฟล์หรือไม่?"):
                os.startfile(output_file)
                
        except Exception as e:
            messagebox.showerror("Error", f"เกิดข้อผิดพลาด: {str(e)}")
            self.status_label.config(
                text="เกิดข้อผิดพลาดในการแปลงไฟล์",
                foreground='Red'  
            )
            self.progress["value"] = 0

    def clean_value(self, value):
        if value is None:
            return ''
        if isinstance(value, (int, float)):
            return str(value)
        return re.sub(r'[^0-9a-zA-Z.]', '', str(value).strip())

if __name__ == "__main__":
    root = tk.Tk()
    app = DBFConverterApp(root)
    root.mainloop()
