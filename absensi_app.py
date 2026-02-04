"""
Aplikasi GUI untuk Sistem Absensi Karyawan
Modern, User-Friendly Interface dengan Tkinter
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
import os
from attendance_processor import AttendanceProcessor


class AttendanceApp:
    """Aplikasi GUI untuk pengolahan data absensi"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Sistem Absensi Karyawan - v1.0")
        self.root.geometry("900x700")
        self.root.configure(bg='#f0f0f0')
        
        # Processor instance
        self.processor = AttendanceProcessor()
        
        # Variables
        self.csv_file_path = tk.StringVar()
        self.selected_employee = tk.StringVar()
        self.selected_month = tk.IntVar(value=datetime.now().month)
        self.selected_year = tk.IntVar(value=datetime.now().year)
        self.output_filename = tk.StringVar()
        
        # Setup UI
        self.setup_ui()
        
    def setup_ui(self):
        """Setup semua komponen UI"""
        
        # Header
        header_frame = tk.Frame(self.root, bg='#2c3e50', height=80)
        header_frame.pack(fill='x', pady=(0, 20))
        header_frame.pack_propagate(False)
        
        title_label = tk.Label(
            header_frame,
            text="📊 SISTEM ABSENSI KARYAWAN",
            font=('Arial', 20, 'bold'),
            fg='white',
            bg='#2c3e50'
        )
        title_label.pack(pady=15)
        
        subtitle_label = tk.Label(
            header_frame,
            text="Proses data absensi dengan mudah dan cepat",
            font=('Arial', 10),
            fg='#ecf0f1',
            bg='#2c3e50'
        )
        subtitle_label.pack()
        
        # Main Container
        main_frame = tk.Frame(self.root, bg='#f0f0f0')
        main_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        # === SECTION 1: Input File ===
        self.create_section(main_frame, "1️⃣ Pilih File Data Absensi", 0)
        
        file_frame = tk.Frame(main_frame, bg='white', relief='ridge', bd=2)
        file_frame.grid(row=1, column=0, sticky='ew', pady=(0, 20), padx=5, ipady=10)
        
        tk.Label(
            file_frame,
            text="File CSV:",
            font=('Arial', 10, 'bold'),
            bg='white'
        ).grid(row=0, column=0, padx=15, pady=10, sticky='w')
        
        file_entry = tk.Entry(
            file_frame,
            textvariable=self.csv_file_path,
            font=('Arial', 10),
            width=50,
            state='readonly'
        )
        file_entry.grid(row=0, column=1, padx=5, pady=10)
        
        browse_btn = tk.Button(
            file_frame,
            text="📁 Browse",
            command=self.browse_file,
            font=('Arial', 10, 'bold'),
            bg='#3498db',
            fg='white',
            cursor='hand2',
            relief='flat',
            padx=20,
            pady=5
        )
        browse_btn.grid(row=0, column=2, padx=15, pady=10)
        
        # === SECTION 2: Pilih Karyawan ===
        self.create_section(main_frame, "2️⃣ Pilih Karyawan", 2)
        
        employee_frame = tk.Frame(main_frame, bg='white', relief='ridge', bd=2)
        employee_frame.grid(row=3, column=0, sticky='ew', pady=(0, 20), padx=5, ipady=10)
        
        tk.Label(
            employee_frame,
            text="Nama Karyawan:",
            font=('Arial', 10, 'bold'),
            bg='white'
        ).grid(row=0, column=0, padx=15, pady=10, sticky='w')
        
        self.employee_combo = ttk.Combobox(
            employee_frame,
            textvariable=self.selected_employee,
            font=('Arial', 10),
            width=60,
            state='readonly'
        )
        self.employee_combo.grid(row=0, column=1, padx=15, pady=10, columnspan=2)
        
        # === SECTION 3: Periode ===
        self.create_section(main_frame, "3️⃣ Pilih Periode Laporan", 4)
        
        period_frame = tk.Frame(main_frame, bg='white', relief='ridge', bd=2)
        period_frame.grid(row=5, column=0, sticky='ew', pady=(0, 20), padx=5, ipady=10)
        
        # Bulan
        tk.Label(
            period_frame,
            text="Bulan:",
            font=('Arial', 10, 'bold'),
            bg='white'
        ).grid(row=0, column=0, padx=15, pady=10, sticky='w')
        
        months = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
                  'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember']
        
        self.month_combo = ttk.Combobox(
            period_frame,
            values=months,
            font=('Arial', 10),
            width=15,
            state='readonly'
        )
        self.month_combo.grid(row=0, column=1, padx=10, pady=10)
        self.month_combo.current(datetime.now().month - 1)
        
        # Tahun
        tk.Label(
            period_frame,
            text="Tahun:",
            font=('Arial', 10, 'bold'),
            bg='white'
        ).grid(row=0, column=2, padx=15, pady=10, sticky='w')
        
        years = list(range(2020, 2031))
        self.year_combo = ttk.Combobox(
            period_frame,
            values=years,
            font=('Arial', 10),
            width=10,
            state='readonly'
        )
        self.year_combo.grid(row=0, column=3, padx=10, pady=10)
        self.year_combo.set(datetime.now().year)
        
        # === SECTION 4: Output ===
        self.create_section(main_frame, "4️⃣ Nama File Output", 6)
        
        output_frame = tk.Frame(main_frame, bg='white', relief='ridge', bd=2)
        output_frame.grid(row=7, column=0, sticky='ew', pady=(0, 20), padx=5, ipady=10)
        
        tk.Label(
            output_frame,
            text="Nama File:",
            font=('Arial', 10, 'bold'),
            bg='white'
        ).grid(row=0, column=0, padx=15, pady=10, sticky='w')
        
        output_entry = tk.Entry(
            output_frame,
            textvariable=self.output_filename,
            font=('Arial', 10),
            width=50
        )
        output_entry.grid(row=0, column=1, padx=5, pady=10)
        
        tk.Label(
            output_frame,
            text=".xlsx",
            font=('Arial', 10),
            bg='white'
        ).grid(row=0, column=2, padx=5, pady=10)
        
        # Auto-generate button
        auto_btn = tk.Button(
            output_frame,
            text="🔄 Auto",
            command=self.auto_generate_filename,
            font=('Arial', 9),
            bg='#95a5a6',
            fg='white',
            cursor='hand2',
            relief='flat',
            padx=10,
            pady=3
        )
        auto_btn.grid(row=0, column=3, padx=10, pady=10)
        
        # === SECTION 5: Action Buttons ===
        button_frame = tk.Frame(main_frame, bg='#f0f0f0')
        button_frame.grid(row=8, column=0, pady=20)
        
        # Process Button
        process_btn = tk.Button(
            button_frame,
            text="▶️ PROSES DATA",
            command=self.process_data,
            font=('Arial', 12, 'bold'),
            bg='#27ae60',
            fg='white',
            cursor='hand2',
            relief='flat',
            padx=30,
            pady=15,
            width=20
        )
        process_btn.pack(side='left', padx=10)
        
        # Reset Button
        reset_btn = tk.Button(
            button_frame,
            text="🔄 RESET",
            command=self.reset_form,
            font=('Arial', 12, 'bold'),
            bg='#e74c3c',
            fg='white',
            cursor='hand2',
            relief='flat',
            padx=30,
            pady=15,
            width=15
        )
        reset_btn.pack(side='left', padx=10)
        
        # === Status Bar ===
        self.status_var = tk.StringVar(value="Siap. Silakan pilih file data absensi.")
        status_bar = tk.Label(
            self.root,
            textvariable=self.status_var,
            font=('Arial', 9),
            bg='#34495e',
            fg='white',
            anchor='w',
            relief='sunken'
        )
        status_bar.pack(side='bottom', fill='x')
        
        # Configure grid weights
        main_frame.columnconfigure(0, weight=1)
        
    def create_section(self, parent, title, row):
        """Helper untuk membuat section header"""
        section_label = tk.Label(
            parent,
            text=title,
            font=('Arial', 11, 'bold'),
            bg='#f0f0f0',
            fg='#2c3e50',
            anchor='w'
        )
        section_label.grid(row=row, column=0, sticky='w', pady=(10, 5), padx=5)
    
    def browse_file(self):
        """Browse dan load file CSV"""
        filename = filedialog.askopenfilename(
            title="Pilih File Data Absensi",
            filetypes=[
                ("CSV Files", "*.csv"),
                ("All Files", "*.*")
            ]
        )
        
        if filename:
            self.csv_file_path.set(filename)
            self.status_var.set("Memuat file...")
            self.root.update()
            
            # Load CSV
            success, message = self.processor.load_csv(filename)
            
            if success:
                # Populate employee list
                employees = self.processor.get_employee_list()
                self.employee_combo['values'] = employees
                
                if employees:
                    self.employee_combo.current(0)
                
                self.status_var.set(f"✓ {message}")
                messagebox.showinfo("Sukses", message)
            else:
                self.status_var.set(f"✗ {message}")
                messagebox.showerror("Error", message)
    
    def auto_generate_filename(self):
        """Generate nama file otomatis"""
        month = self.month_combo.current() + 1
        year = self.year_combo.get()
        month_name = self.month_combo.get()
        
        filename = f"Rekap_Absensi_{month_name}_{year}"
        self.output_filename.set(filename)
    
    def process_data(self):
        """Proses data absensi"""
        
        # Validasi input
        if not self.csv_file_path.get():
            messagebox.showwarning("Peringatan", "Silakan pilih file CSV terlebih dahulu!")
            return
        
        if not self.selected_employee.get():
            messagebox.showwarning("Peringatan", "Silakan pilih karyawan!")
            return
        
        if not self.output_filename.get():
            messagebox.showwarning("Peringatan", "Silakan masukkan nama file output!")
            return
        
        # Get parameters
        employee = self.selected_employee.get()
        month = self.month_combo.current() + 1
        year = int(self.year_combo.get())
        output_file = self.output_filename.get()
        
        if not output_file.endswith('.xlsx'):
            output_file += '.xlsx'
        
        self.status_var.set("Memproses data...")
        self.root.update()
        
        # Process
        success, message, df = self.processor.process_employee_attendance(
            employee, month, year
        )
        
        if success:
            # Save to Excel
            save_success, save_message = self.processor.save_to_excel(df, output_file)
            
            if save_success:
                self.status_var.set(f"✓ Berhasil! Data disimpan ke: {output_file}")
                
                result_msg = f"""
Proses Berhasil! ✓

Karyawan: {employee}
Periode: {self.month_combo.get()} {year}
Total Hari: {len(df)}
File Output: {output_file}

Data telah disimpan dengan sukses!
                """
                messagebox.showinfo("Sukses", result_msg.strip())
            else:
                self.status_var.set(f"✗ {save_message}")
                messagebox.showerror("Error", save_message)
        else:
            self.status_var.set(f"✗ {message}")
            messagebox.showerror("Error", message)
    
    def reset_form(self):
        """Reset form ke kondisi awal"""
        self.csv_file_path.set("")
        self.selected_employee.set("")
        self.employee_combo['values'] = []
        self.month_combo.current(datetime.now().month - 1)
        self.year_combo.set(datetime.now().year)
        self.output_filename.set("")
        self.status_var.set("Form direset. Siap untuk input baru.")
        
        # Reset processor
        self.processor = AttendanceProcessor()


def main():
    """Main function"""
    root = tk.Tk()
    app = AttendanceApp(root)
    
    # Center window
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')
    
    root.mainloop()


if __name__ == "__main__":
    main()
