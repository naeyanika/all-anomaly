import streamlit as st
import pandas as pd
import numpy as np
import io

st.title('Aplikasi Analisa THC Pinjaman dan Simpanan')
st.markdown("""
## File yang dibutuhkan
1. **Anomali Simpanan.xlsx**
2. **Anomali Pinjaman.xlsx**
3. **DbSimpanan.xlsx**
    - Hapus bagian header terlebih dahulu
    - Nama File harus DbSimpanan.xlsx dan sheet atau lembar nya IA_SimpananDB
    """)

def format_no(no):
    try:
        if pd.notna(no):
            return f'{int(no):02d}.'
        else:
            return ''
    except (ValueError, TypeError):
        return str(no)

def format_center(center):
    try:
        if pd.notna(center):
            return f'{int(center):03d}'
        else:
            return ''
    except (ValueError, TypeError):
        return str(center)

def format_kelompok(kelompok):
    try:
        if pd.notna(kelompok):
            return f'{int(kelompok):02d}'
        else:
            return ''
    except (ValueError, TypeError):
        return str(kelompok)

def load_data(uploaded_files):
    dfs = {}
    for file in uploaded_files:
        try:
            excel_file = pd.ExcelFile(file, engine='openpyxl')
            
            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(excel_file, sheet_name=sheet_name)
                key = f"{file.name}_{sheet_name}"
                dfs[key] = df
            
            st.success(f"File {file.name} berhasil diunggah dan diproses.")
        except Exception as e:
            st.error(f"Terjadi kesalahan saat memproses file {file.name}: {str(e)}")
    return dfs

def process_data(dfs):
    try:
        # Pinjaman
        df_pu = dfs['Anomali Pinjaman.xlsx_Anomali PU'][['ID','CEK KRITERIA']].rename(columns={'CEK KRITERIA':'PU'})
        df_pu['PU'] = df_pu['PU'].replace({True: 0, False: 1})

        df_pmb = dfs['Anomali Pinjaman.xlsx_Anomali PMB'][['ID','CEK KRITERIA']].rename(columns={'CEK KRITERIA':'PMB'})
        df_pmb['PMB'] = df_pmb['PMB'].replace({True: 0, False: 1})

        df_psa = dfs['Anomali Pinjaman.xlsx_Anomali PSA'][['ID','CEK KRITERIA','Sukarela Sesuai','Wajib Sesuai','Pensiun Sesuai']]
        df_psa[['CEK KRITERIA', 'Sukarela Sesuai', 'Wajib Sesuai', 'Pensiun Sesuai']] = df_psa[['CEK KRITERIA', 'Sukarela Sesuai', 'Wajib Sesuai', 'Pensiun Sesuai']].replace({True: 0, False: 1})
        df_psa['PSA'] = df_psa[['CEK KRITERIA', 'Sukarela Sesuai', 'Wajib Sesuai', 'Pensiun Sesuai']].sum(axis=1)
        df_psa = df_psa[['ID','PSA']]

        df_prr = dfs['Anomali Pinjaman.xlsx_Anomali PRR'][['ID','CEK KRITERIA','Sukarela Sesuai','Wajib Sesuai','Pensiun Sesuai']]
        df_prr[['CEK KRITERIA', 'Sukarela Sesuai', 'Wajib Sesuai', 'Pensiun Sesuai']] = df_prr[['CEK KRITERIA', 'Sukarela Sesuai', 'Wajib Sesuai', 'Pensiun Sesuai']].replace({True: 0, False: 1})
        df_prr['PRR'] = df_prr[['CEK KRITERIA', 'Sukarela Sesuai', 'Wajib Sesuai', 'Pensiun Sesuai']].sum(axis=1)
        df_prr = df_prr[['ID','PRR']]

        df_ptn = dfs['Anomali Pinjaman.xlsx_Anomali PTN'][['ID','SEMUA KRITERIA TERPENUHI']].rename(columns={'SEMUA KRITERIA TERPENUHI':'PTN'})
        df_ptn['PTN'] = df_ptn['PTN'].replace({True: 0, False: 1})

        df_arta = dfs['Anomali Pinjaman.xlsx_Anomali ARTA'][['ID','CEK KRITERIA']].rename(columns={'CEK KRITERIA':'ARTA'})
        df_arta['ARTA'] = df_arta['ARTA'].replace({True: 0, False: 1})

        df_dtp = dfs['Anomali Pinjaman.xlsx_Anomali DTP'][['ID','CEK KRITERIA']].rename(columns={'CEK KRITERIA':'DTP'})
        df_dtp['DTP'] = df_dtp['DTP'].replace({True: 0, False: 1})

        # Simpanan
        df_sukarela = dfs['Anomali Simpanan.xlsx_Sukarela'][['ID','Transaksi > 0 & ≠ Modus Sukarela']].rename(columns={'Transaksi > 0 & ≠ Modus Sukarela':'SUKARELA'})
        df_sihara = dfs['Anomali Simpanan.xlsx_Sihara'][['ID','TRANSAKSI TIDAK SESUAI']].rename(columns={'TRANSAKSI TIDAK SESUAI':'HARI RAYA'})
        df_pensiun = dfs['Anomali Simpanan.xlsx_Pensiun'][['ID','Anomali']].rename(columns={'Anomali':'PENSIUN'})

        # Merge all data
        df_selected_all = dfs['DbSimpanan.xlsx_IA_SimpananDB'][['Client ID','Client Name','Center ID','Group ID','Meeting Day','Officer Name']].rename(columns={
            'Client ID':'ID', 'Client Name': 'NAMA', 'Center ID': 'CENTER', 'Group ID': 'KELOMPOK', 'Meeting Day': 'HARI', 'Officer Name': 'STAF'
        })

        for df in [df_sukarela, df_pensiun, df_sihara, df_pu, df_pmb, df_psa, df_prr, df_ptn, df_arta, df_dtp]:
            df_selected_all = df_selected_all.merge(df, on='ID', how='left')

        df_selected_all = df_selected_all.fillna(0)

        anomali_columns = ['SUKARELA', 'PENSIUN', 'HARI RAYA', 'PU', 'PMB', 'PSA', 'PRR', 'PTN', 'ARTA', 'DTP']
        df_selected_all['TOTAL ANOMALI'] = df_selected_all[anomali_columns].sum(axis=1)

        df_selected_all = df_selected_all[['ID', 'NAMA', 'CENTER', 'KELOMPOK', 'HARI', 'STAF'] + anomali_columns + ['TOTAL ANOMALI']]
        df_selected_all = df_selected_all.drop_duplicates(subset=['ID', 'NAMA'])

        # Pisahkan anomali simpanan dan pinjaman
        anomali_simpanan = ['SUKARELA', 'PENSIUN', 'HARI RAYA']
        anomali_pinjaman = ['PU', 'PMB', 'PSA', 'PRR', 'PTN', 'ARTA', 'DTP']

        df_selected_all['Total Anomali Simpanan'] = df_selected_all[anomali_simpanan].sum(axis=1)
        df_selected_all['Total Anomali Pinjaman'] = df_selected_all[anomali_pinjaman].sum(axis=1)

        # Buat dictionary untuk memetakan hari bahasa Indonesia ke bahasa Inggris
        hari_map = {
        'SENIN': 'Monday',
        'SELASA': 'Tuesday',
        'RABU': 'Wednesday',
        'KAMIS': 'Thursday',
        'JUMAT': 'Friday',
        'SABTU': 'Saturday',
        'MINGGU': 'Sunday'
        }

        # Konversi kolom HARI ke bahasa Inggris
        df_selected_all['HARI_EN'] = df_selected_all['HARI'].map(hari_map)

        # Buat pivot table untuk setiap hari
        pivot_dfs = []
        for day in ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']:
            pivot = df_selected_all[df_selected_all['HARI_EN'] == day].pivot_table(
                values=['Total Anomali Simpanan', 'Total Anomali Pinjaman'],
                index=['STAF', 'CENTER'],
                aggfunc='sum'
            ).reset_index()
    
            pivot.columns = [f'{day}_{col}' if col not in ['STAF', 'CENTER'] else col for col in pivot.columns]
            pivot_dfs.append(pivot)

        # Gabungkan semua pivot table
        result = pivot_dfs[0]
        for df in pivot_dfs[1:]:
            result = result.merge(df, on=['STAF', 'CENTER'], how='outer')

        # Urutkan kolom
        column_order = ['STAF', 'CENTER']
        for day in ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']:
            column_order.extend([f'{day}_Total Anomali Simpanan', f'{day}_Total Anomali Pinjaman'])

        result = result.reindex(columns=column_order)

        # Ganti nama kolom kembali ke bahasa Indonesia
        rename_dict = {
        'Monday': 'SENIN',
        'Tuesday': 'SELASA',
        'Wednesday': 'RABU',
        'Thursday': 'KAMIS',
        'Friday': 'JUMAT',
        'Saturday': 'SABTU',
        'Sunday': 'MINGGU'
        }

        result.columns = [' '.join(col.split('_')).replace(day_en, day_id) 
                          for col in result.columns 
                          for day_en, day_id in rename_dict.items() 
                          if day_en in col]

        # Isi NaN dengan 0
        result = result.fillna(0)

# Kembalikan kedua dataframe
        return df_selected_all, result

    except KeyError as e:
        st.error(f"Terjadi kesalahan saat memproses data: {str(e)}. Pastikan semua file dan sheet yang diperlukan telah diunggah.")
        return None

def main():
    uploaded_files = st.file_uploader("Unggah file Excel", accept_multiple_files=True, type=["xlsx"])

    if uploaded_files:
        dfs = load_data(uploaded_files)
        
        required_files = [
            'Anomali Pinjaman.xlsx_Anomali PU',
            'Anomali Pinjaman.xlsx_Anomali PMB',
            'Anomali Pinjaman.xlsx_Anomali DTP',
            'Anomali Pinjaman.xlsx_Anomali PSA',
            'Anomali Pinjaman.xlsx_Anomali ARTA',
            'Anomali Pinjaman.xlsx_Anomali PRR',
            'Anomali Pinjaman.xlsx_Anomali PTN',
            'Anomali Simpanan.xlsx_Sihara',
            'Anomali Simpanan.xlsx_Pensiun',
            'Anomali Simpanan.xlsx_Sukarela',
            'DbSimpanan.xlsx_IA_SimpananDB'
        ]
        
        if all(file in dfs for file in required_files):
            df_selected_all, result = process_data(dfs)
            
            if df_selected_all is not None and result is not None:
                st.write("Data setelah diproses:")
                st.write(df_selected_all)
                
                st.write("Data Pivot:")
                st.write(result)
                
                # Unduh Data Anomali
                buffer1 = io.BytesIO()
                with pd.ExcelWriter(buffer1, engine='xlsxwriter') as writer:
                    df_selected_all.to_excel(writer, index=False, sheet_name='Sheet1')
                buffer1.seek(0)
                st.download_button(
                    label="Unduh Data Anomali.xlsx",
                    data=buffer1.getvalue(),
                    file_name="Data Anomali.xlsx",
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
                
                # Unduh Data Pivot
                buffer2 = io.BytesIO()
                with pd.ExcelWriter(buffer2, engine='xlsxwriter') as writer:
                    result.to_excel(writer, index=False, sheet_name='Pivot')
                buffer2.seek(0)
                st.download_button(
                    label="Unduh Data Pivot.xlsx",
                    data=buffer2.getvalue(),
                    file_name="Data Pivot.xlsx",
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
            else:
                st.error("Terjadi kesalahan saat memproses data.")
        else:
            missing_files = [file for file in required_files if file not in dfs]
            st.warning(f"Beberapa file atau sheet yang diperlukan belum diunggah: {', '.join(missing_files)}")

if __name__ == "__main__":
    main()
