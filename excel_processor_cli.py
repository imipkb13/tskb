import pandas as pd
import os
import re
from datetime import datetime


class ExcelDataProcessor:
    def __init__(self, source_folder="", rekap_file=""):
        self.source_folder = source_folder
        self.rekap_file = rekap_file
        
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
    
    def find_keyword_value(self, df, keyword):
        """Mencari keyword dan mengambil nilai di kolom setelahnya"""
        try:
            for idx, row in df.iterrows():
                for col_idx, cell in enumerate(row):
                    if isinstance(cell, str) and keyword.upper() in cell.upper():
                        # Ambil nilai di kolom setelahnya
                        if col_idx + 1 < len(row):
                            value = row.iloc[col_idx + 1]
                            if pd.notna(value):
                                return self.clean_value(str(value))
                        # Jika keyword ada di akhir baris, cek baris berikutnya di kolom yang sama
                        if idx + 1 < len(df):
                            value = df.iloc[idx + 1, col_idx]
                            if pd.notna(value):
                                return self.clean_value(str(value))
            return ""
        except Exception as e:
            print(f"Error finding keyword {keyword}: {e}")
            return ""
    
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
            print(f"Error processing packing list {filepath}: {str(e)}")
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
                print(f"Sheet CI not found or error: {str(e)}")
            
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
                print(f"Sheet attachment not found or error: {str(e)}")
                result['details'] = []
            
            return result
        except Exception as e:
            print(f"Error processing invoice {filepath}: {str(e)}")
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
            
            print(f"Ditemukan {len(pl_files)} file packing list dan {len(iv_files)} file invoice")
            
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
                
                print(f"Memproses packing list: {pl_file}")
            
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
                
                print(f"Memproses invoice: {iv_file}")
            
            print("Menyimpan data ke file rekap...")
            self.save_to_rekap(ship_data)
            
            print("Selesai!")
            return True
            
        except Exception as e:
            print(f"Error: {str(e)}")
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
            
            print(f"Data berhasil disimpan ke {self.rekap_file}")
            
        except Exception as e:
            raise Exception(f"Error saving to rekap: {str(e)}")


def main():
    print("=" * 60)
    print("Aplikasi Olah Data Excel Impor Barang")
    print("=" * 60)
    
    source_folder = input("\nMasukkan path folder sumber: ").strip()
    if not os.path.isdir(source_folder):
        print("Error: Folder tidak ditemukan!")
        return
    
    rekap_file = input("Masukkan path file rekap output (.xlsx): ").strip()
    if not rekap_file.endswith('.xlsx'):
        rekap_file += '.xlsx'
    
    processor = ExcelDataProcessor(source_folder, rekap_file)
    processor.process_files()
    
    print("\n" + "=" * 60)
    print("Proses selesai!")
    print("=" * 60)


if __name__ == "__main__":
    main()
