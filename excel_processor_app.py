import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import re
from datetime import datetime
import threading


class ExcelDataProcessor:
    def __init__(self):
        self.source_folder = ""
        self.rekap_file = ""
        self.progress_callback = None
        
    def set_progress_callback(self, callback):
        self.progress_callback = callback
        
    def update_progress(self, message, value=None):
        if self.progress_callback:
            self.progress_callback(message, value)
    
    def clean_value(self, value):
        """Menghapus tanda ':' dari value"""
        if isinstance(value, str):
            return value.replace(':', '').strip()
        return value
    
    def clean_column_n(self, value):
        """Menghapus semua tanda '.' dari data kolom N"""
        if isinstance(value, str):
            return value.replace('.', '').strip()
        return value
    
    def extract_ship_number(self, filename):
        """Mengambil 4 karakter pertama sebelum '_' sebagai nomor kapal"""
        base_name = os.path.basename(filename)
        match = re.match(r'^([A-Za-z0-9]{4})_', base_name)
        if match:
            return match.group(1)
        return ""
    
    def find_keyword_value(self, df, keyword, sheet_name=None):
        """Mencari keyword dan mengambil nilai di kolom setelahnya"""
        try:
            for idx, row in df.iterrows():
                for col_idx, cell in enumerate(row):
                    if isinstance(cell, str) and keyword in cell.upper():
                        # Ambil nilai di kolom setelahnya
                        if col_idx + 1 < len(row):
                            value = row.iloc[col_idx + 1]
                            if pd.notna(value):
                                return self.clean_value(str(value))
                        # Jika keyword ada di akhir baris, cek baris berikutnya di kolom yang sama
                        if col_idx < len(row):
                            if idx + 1 < len(df):
                                value = df.iloc[idx + 1, col_idx]
                                if pd.notna(value):
                                    return self.clean_value(str(value))
            return ""
        except Exception as e:
            print(f"Error finding keyword {keyword}: {e}")
            return ""
    
    def find_keyword_cell(self, df, keyword):
        """Mencari posisi cell yang mengandung keyword"""
        for idx, row in df.iterrows():
            for col_idx, cell in enumerate(row):
                if isinstance(cell, str) and keyword.upper() in cell.upper():
                    return idx, col_idx
        return None, None
    
    def process_packing_list(self, filepath):
        """Proses file packing list"""
        try:
            df = pd.read_excel(filepath, header=None)
            result = {}
            
            # Row ke-2 adalah nama supplier (index 1 karena 0-based)
            if len(df) > 1:
                supplier = df.iloc[1, 0] if pd.notna(df.iloc[1, 0]) else ""
                result['supplier'] = self.clean_value(str(supplier))
            
            # Cari "Ocean Vessel"
            ocean_vessel = self.find_keyword_value(df, "Ocean Vessel")
            result['ocean_vessel'] = ocean_vessel
            
            # Cari "INVOICE NO."
            invoice_no = self.find_keyword_value(df, "INVOICE NO.")
            result['invoice_no'] = invoice_no
            
            # Cari "B/L NO."
            bl_no = self.find_keyword_value(df, "B/L NO.")
            result['bl_no'] = bl_no
            
            # Cari "Date:"
            date_val = self.find_keyword_value(df, "Date:")
            result['date'] = date_val
            
            # Cari "Net Weight"
            net_weight = self.find_keyword_value(df, "Net Weight")
            result['net_weight'] = net_weight
            
            # Nomor kapal dari nama file
            ship_number = self.extract_ship_number(filepath)
            result['ship_number'] = ship_number
            
            return result
        except Exception as e:
            self.update_progress(f"Error processing packing list {filepath}: {str(e)}")
            return None
    
    def process_invoice(self, filepath):
        """Proses file invoice dengan 2 sheet: attachment dan CI"""
        try:
            result = {}
            
            # Baca sheet CI
            try:
                df_ci = pd.read_excel(filepath, sheet_name='CI', header=None)
                
                # Cari "B/L NO." di sheet CI
                bl_no = self.find_keyword_value(df_ci, "B/L NO.")
                result['bl_no'] = bl_no
                
                # Cari "INVOICE NO." di sheet CI
                invoice_no = self.find_keyword_value(df_ci, "INVOICE NO.")
                result['invoice_no'] = invoice_no
                
                # Cari "Date:" di sheet CI
                date_val = self.find_keyword_value(df_ci, "Date:")
                result['date'] = date_val
                
            except Exception as e:
                self.update_progress(f"Sheet CI not found or error: {str(e)}")
            
            # Baca sheet attachment untuk detail barang
            try:
                df_attachment = pd.read_excel(filepath, sheet_name='attachment', header=None)
                details = []
                
                # Mulai dari row yang memiliki angka 1 di kolom A
                start_row = None
                for idx, row in df_attachment.iterrows():
                    if pd.notna(row.iloc[0]) and str(row.iloc[0]).strip() == '1':
                        start_row = idx
                        break
                
                if start_row is not None:
                    # Proses setiap baris selama kolom A ada nomornya
                    current_num = 1
                    for idx in range(start_row, len(df_attachment)):
                        row = df_attachment.iloc[idx]
                        
                        # Cek apakah kolom A masih ada nomor urut
                        if pd.isna(row.iloc[0]):
                            break
                        
                        try:
                            seq_num = int(float(str(row.iloc[0]).strip()))
                            if seq_num != current_num:
                                break
                            current_num += 1
                        except:
                            break
                        
                        # Ekstrak data sesuai kolom
                        detail = {}
                        
                        # Kolom A - nomor urut -> kolom M rekap sheet 2
                        detail['seq_no'] = self.clean_value(str(row.iloc[0])) if pd.notna(row.iloc[0]) else ""
                        
                        # Kolom B - nomor kontrak -> kolom Y rekap sheet 2
                        detail['contract_no'] = self.clean_value(str(row.iloc[1])) if pd.notna(row.iloc[1]) else ""
                        
                        # Kolom C, D, E -> kolom O, P, Q rekap sheet 2
                        detail['col_c'] = self.clean_value(str(row.iloc[2])) if pd.notna(row.iloc[2]) else ""
                        detail['col_d'] = self.clean_value(str(row.iloc[3])) if pd.notna(row.iloc[3]) else ""
                        detail['col_e'] = self.clean_value(str(row.iloc[4])) if pd.notna(row.iloc[4]) else ""
                        
                        # Kolom F, G, H, I -> kolom S, T, U, V rekap sheet 2
                        detail['col_f'] = self.clean_value(str(row.iloc[5])) if pd.notna(row.iloc[5]) else ""
                        detail['col_g'] = self.clean_value(str(row.iloc[6])) if pd.notna(row.iloc[6]) else ""
                        detail['col_h'] = self.clean_value(str(row.iloc[7])) if pd.notna(row.iloc[7]) else ""
                        detail['col_i'] = self.clean_value(str(row.iloc[8])) if pd.notna(row.iloc[8]) else ""
                        
                        # Kolom L -> kolom X rekap sheet 2
                        detail['col_l'] = self.clean_value(str(row.iloc[11])) if pd.notna(row.iloc[11]) else ""
                        
                        # Kolom N -> kolom AA rekap sheet 2 (hapus semua '.')
                        detail['col_n'] = self.clean_column_n(str(row.iloc[13])) if pd.notna(row.iloc[13]) else ""
                        
                        details.append(detail)
                
                result['details'] = details
                
            except Exception as e:
                self.update_progress(f"Sheet attachment not found or error: {str(e)}")
                result['details'] = []
            
            return result
        except Exception as e:
            self.update_progress(f"Error processing invoice {filepath}: {str(e)}")
            return None
    
    def process_files(self):
        """Proses semua file dalam folder sumber"""
        try:
            if not os.path.isdir(self.source_folder):
                raise Exception("Folder sumber tidak valid")
            
            # Dapatkan semua file _pl dan _iv
            files = os.listdir(self.source_folder)
            pl_files = [f for f in files if '_pl' in f.lower()]
            iv_files = [f for f in files if '_iv' in f.lower()]
            
            self.update_progress(f"Ditemukan {len(pl_files)} file packing list dan {len(iv_files)} file invoice", 10)
            
            # Kelompokkan file berdasarkan nomor kapal (4 karakter pertama)
            ship_data = {}
            
            # Proses packing list
            for i, pl_file in enumerate(pl_files):
                ship_num = self.extract_ship_number(pl_file)
                if not ship_num:
                    continue
                    
                filepath = os.path.join(self.source_folder, pl_file)
                pl_data = self.process_packing_list(filepath)
                
                if pl_data:
                    if ship_num not in ship_data:
                        ship_data[ship_num] = {'pl': [], 'iv': []}
                    ship_data[ship_num]['pl'].append(pl_data)
                
                progress = 10 + (i / max(len(pl_files), 1)) * 40
                self.update_progress(f"Memproses packing list: {pl_file}", progress)
            
            # Proses invoice
            for i, iv_file in enumerate(iv_files):
                ship_num = self.extract_ship_number(iv_file)
                if not ship_num:
                    continue
                    
                filepath = os.path.join(self.source_folder, iv_file)
                iv_data = self.process_invoice(filepath)
                
                if iv_data:
                    if ship_num not in ship_data:
                        ship_data[ship_num] = {'pl': [], 'iv': []}
                    ship_data[ship_num]['iv'].append(iv_data)
                
                progress = 50 + (i / max(len(iv_files), 1)) * 40
                self.update_progress(f"Memproses invoice: {iv_file}", progress)
            
            self.update_progress("Menyimpan data ke file rekap...", 90)
            self.save_to_rekap(ship_data)
            
            self.update_progress("Selesai!", 100)
            return True
            
        except Exception as e:
            self.update_progress(f"Error: {str(e)}")
            return False
    
    def save_to_rekap(self, ship_data):
        """Simpan data ke file rekap dengan 2 sheet"""
        try:
            # Siapkan data untuk sheet 1 (Rekap BL)
            rekap_bl_data = []
            
            # Siapkan data untuk sheet 2 (Detail BL)
            detail_bl_data = []
            
            for ship_num, data in ship_data.items():
                for pl in data['pl']:
                    # Sheet 1: Rekap BL
                    # Kolom A-J dan L-N
                    row_bl = {
                        'A': '',  # Akan diisi sesuai kebutuhan
                        'B': '',  # Akan diisi sesuai kebutuhan
                        'C': pl.get('ship_number', ''),  # Nomor kapal (4 karakter)
                        'D': '',  # Akan diisi sesuai kebutuhan
                        'E': pl.get('ocean_vessel', ''),  # Nama kapal dari Ocean Vessel
                        'F': pl.get('supplier', ''),  # Supplier dari row 2 packing list
                        'G': '',  # Akan diisi sesuai kebutuhan
                        'H': pl.get('bl_no', ''),  # B/L NO.
                        'I': pl.get('invoice_no', ''),  # INVOICE NO.
                        'J': pl.get('date', ''),  # Date
                        'L': '',  # Akan diisi sesuai kebutuhan
                        'M': pl.get('net_weight', ''),  # Net Weight
                        'N': ''  # Akan diisi sesuai kebutuhan
                    }
                    rekap_bl_data.append(row_bl)
                    
                    # Sheet 2: Detail BL
                    # Cari invoice yang sesuai berdasarkan BL NO atau Invoice NO
                    for iv in data['iv']:
                        if iv.get('bl_no') == pl.get('bl_no') or iv.get('invoice_no') == pl.get('invoice_no'):
                            # Tambahkan detail dari attachment
                            for detail in iv.get('details', []):
                                row_detail = {
                                    'A': '',  # Will be filled
                                    'B': '',  # Will be filled
                                    'C': pl.get('ship_number', ''),  # Nomor kapal
                                    'D': '',  # Will be filled
                                    'E': pl.get('ocean_vessel', ''),  # Nama kapal
                                    'F': pl.get('supplier', ''),  # Supplier
                                    'G': '',  # Will be filled
                                    'H': pl.get('bl_no', ''),  # B/L NO.
                                    'I': iv.get('bl_no', ''),  # B/L NO. dari invoice CI
                                    'J': iv.get('invoice_no', ''),  # INVOICE NO. dari invoice CI
                                    'K': iv.get('date', ''),  # Date dari invoice CI
                                    'L': '',  # Will be filled
                                    'M': detail.get('seq_no', ''),  # Nomor urut dari attachment col A
                                    'N': '',  # Will be filled
                                    'O': detail.get('col_c', ''),  # Col C dari attachment
                                    'P': detail.get('col_d', ''),  # Col D dari attachment
                                    'Q': detail.get('col_e', ''),  # Col E dari attachment
                                    'R': '',  # Will be filled
                                    'S': detail.get('col_f', ''),  # Col F dari attachment
                                    'T': detail.get('col_g', ''),  # Col G dari attachment
                                    'U': detail.get('col_h', ''),  # Col H dari attachment
                                    'V': detail.get('col_i', ''),  # Col I dari attachment
                                    'W': '',  # Will be filled
                                    'X': detail.get('col_l', ''),  # Col L dari attachment
                                    'Y': detail.get('contract_no', ''),  # Col B (kontrak) dari attachment
                                    'Z': '',  # Will be filled
                                    'AA': detail.get('col_n', ''),  # Col N dari attachment (tanpa '.')
                                }
                                detail_bl_data.append(row_detail)
            
            # Buat DataFrame untuk sheet 1
            if rekap_bl_data:
                df_rekap_bl = pd.DataFrame(rekap_bl_data)
                # Urutkan kolom sesuai A-J, L-N
                columns_order = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'L', 'M', 'N']
                df_rekap_bl = df_rekap_bl[[col for col in columns_order if col in df_rekap_bl.columns]]
            else:
                df_rekap_bl = pd.DataFrame(columns=['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'L', 'M', 'N'])
            
            # Buat DataFrame untuk sheet 2
            if detail_bl_data:
                df_detail_bl = pd.DataFrame(detail_bl_data)
                # Urutkan kolom sesuai kebutuhan
                columns_order = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 
                                'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA']
                df_detail_bl = df_detail_bl[[col for col in columns_order if col in df_detail_bl.columns]]
            else:
                df_detail_bl = pd.DataFrame(columns=['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 
                                                      'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA'])
            
            # Simpan ke Excel dengan 2 sheet
            with pd.ExcelWriter(self.rekap_file, engine='openpyxl') as writer:
                df_rekap_bl.to_excel(writer, sheet_name='Rekap BL', index=False)
                df_detail_bl.to_excel(writer, sheet_name='Detail BL', index=False)
            
            self.update_progress(f"Data berhasil disimpan ke {self.rekap_file}")
            
        except Exception as e:
            raise Exception(f"Error saving to rekap: {str(e)}")


class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        
        self.title("Aplikasi Olah Data Excel - Import Barang")
        self.geometry("800x600")
        self.resizable(True, True)
        
        self.processor = ExcelDataProcessor()
        self.processor.set_progress_callback(self.update_progress)
        
        self.setup_ui()
    
    def setup_ui(self):
        # Frame utama
        main_frame = ttk.Frame(self, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Konfigurasi grid
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Aplikasi Olah Data Excel Impor Barang", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Frame untuk source folder
        source_frame = ttk.LabelFrame(main_frame, text="Folder Sumber", padding="10")
        source_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        source_frame.columnconfigure(1, weight=1)
        
        ttk.Label(source_frame, text="Folder:").grid(row=0, column=0, sticky=tk.W)
        self.source_var = tk.StringVar()
        self.source_entry = ttk.Entry(source_frame, textvariable=self.source_var, width=50)
        self.source_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=10)
        
        self.browse_btn = ttk.Button(source_frame, text="Browse", command=self.browse_source)
        self.browse_btn.grid(row=0, column=2)
        
        # Frame untuk rekap file
        rekap_frame = ttk.LabelFrame(main_frame, text="File Rekap", padding="10")
        rekap_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        rekap_frame.columnconfigure(1, weight=1)
        
        ttk.Label(rekap_frame, text="File:").grid(row=0, column=0, sticky=tk.W)
        self.rekap_var = tk.StringVar()
        self.rekap_entry = ttk.Entry(rekap_frame, textvariable=self.rekap_var, width=50)
        self.rekap_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=10)
        
        self.rekap_browse_btn = ttk.Button(rekap_frame, text="Browse", command=self.browse_rekap)
        self.rekap_browse_btn.grid(row=0, column=2)
        
        # Frame untuk progress
        progress_frame = ttk.LabelFrame(main_frame, text="Progress", padding="10")
        progress_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        progress_frame.columnconfigure(0, weight=1)
        progress_frame.rowconfigure(1, weight=1)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, 
                                           maximum=100, mode='determinate')
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=5)
        
        self.progress_text = tk.Text(progress_frame, height=15, wrap=tk.WORD)
        self.progress_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        scrollbar = ttk.Scrollbar(progress_frame, orient=tk.VERTICAL, command=self.progress_text.yview)
        scrollbar.grid(row=1, column=1, sticky=(tk.N, tk.S))
        self.progress_text.configure(yscrollcommand=scrollbar.set)
        
        # Frame untuk tombol
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=3, pady=20)
        
        self.process_btn = ttk.Button(button_frame, text="Proses Data", command=self.start_processing)
        self.process_btn.grid(row=0, column=0, padx=10)
        
        self.clear_btn = ttk.Button(button_frame, text="Clear Log", command=self.clear_log)
        self.clear_btn.grid(row=0, column=1, padx=10)
        
        # Info label
        info_frame = ttk.LabelFrame(main_frame, text="Informasi", padding="10")
        info_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        info_text = """
        Cara Penggunaan:
        1. Pilih folder sumber yang berisi file _pl (packing list) dan _iv (invoice)
        2. Tentukan lokasi file rekap output
        3. Klik "Proses Data" untuk memulai
        
        Format File:
        - Packing List (_pl): 1 sheet, row 2 = supplier, cari keyword: Ocean Vessel, INVOICE NO., B/L NO., Date:, Net Weight
        - Invoice (_iv): 2 sheet (CI & attachment), sheet CI cari keyword: B/L NO., INVOICE NO., Date:
        - Sheet attachment: detail barang mulai dari row dengan angka 1 di kolom A
        """
        
        ttk.Label(info_frame, text=info_text, justify=tk.LEFT).grid(row=0, column=0, sticky=tk.W)
    
    def browse_source(self):
        folder = filedialog.askdirectory(title="Pilih Folder Sumber")
        if folder:
            self.source_var.set(folder)
            self.processor.source_folder = folder
            self.log_message(f"Folder sumber dipilih: {folder}")
    
    def browse_rekap(self):
        file = filedialog.asksaveasfilename(
            title="Simpan File Rekap",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file:
            self.rekap_var.set(file)
            self.processor.rekap_file = file
            self.log_message(f"File rekap ditentukan: {file}")
    
    def log_message(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.progress_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.progress_text.see(tk.END)
    
    def clear_log(self):
        self.progress_text.delete(1.0, tk.END)
    
    def update_progress(self, message, value=None):
        self.log_message(message)
        if value is not None:
            self.progress_var.set(value)
    
    def start_processing(self):
        if not self.processor.source_folder:
            messagebox.showerror("Error", "Pilih folder sumber terlebih dahulu!")
            return
        
        if not self.processor.rekap_file:
            messagebox.showerror("Error", "Tentukan file rekap output terlebih dahulu!")
            return
        
        # Disable tombol saat proses
        self.process_btn.config(state='disabled')
        
        # Jalankan proses di thread terpisah
        thread = threading.Thread(target=self.run_processing)
        thread.daemon = True
        thread.start()
    
    def run_processing(self):
        try:
            success = self.processor.process_files()
            if success:
                self.after(0, lambda: messagebox.showinfo("Sukses", "Proses selesai dengan berhasil!"))
            else:
                self.after(0, lambda: messagebox.showerror("Error", "Proses gagal. Lihat log untuk detail."))
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("Error", f"Terjadi kesalahan: {str(e)}"))
        finally:
            self.after(0, lambda: self.process_btn.config(state='normal'))


if __name__ == "__main__":
    app = Application()
    app.mainloop()
