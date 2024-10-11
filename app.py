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

## FUNGSI FORMAT NOMOR
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

uploaded_files = st.file_uploader("Unggah file Excel", accept_multiple_files=True, type=["xlsx"])

if uploaded_files:
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

    # Pengkategorian dataframe
    if 'Anomali Pinjaman.xlsx_Anomali PU' in dfs:
        df_anomali_pu = dfs['Anomali Pinjaman.xlsx_Anomali PU']
        st.success("Data Anomali PU berhasil dimuat.")
    else:
        st.error("File Anomali Pinjaman.xlsx tidak ditemukan atau tidak memiliki sheet 'Anomali PU'.")
        
    if 'Anomali Pinjaman.xlsx_Anomali PMB' in dfs:
        df_anomali_pmb = dfs['Anomali Pinjaman.xlsx_Anomali PMB']
        st.success("Data Anomali PMB berhasil dimuat.")
    else:
        st.warning("File Anomali Pinjaman.xlsx tidak memiliki sheet 'Anomali PMB' atau belum diunggah.")
        
    if 'Anomali Pinjaman.xlsx_Anomali DTP' in dfs:
        df_anomali_dtp = dfs['Anomali Pinjaman.xlsx_Anomali DTP']
        st.success("Data Anomali DTP berhasil dimuat.")
    else:
        st.warning("File Anomali Pinjaman.xlsx tidak memiliki sheet 'Anomali DTP' atau belum diunggah.")
        
    if 'Anomali Pinjaman.xlsx_Anomali PSA' in dfs:
        df_anomali_psa = dfs['Anomali Pinjaman.xlsx_Anomali PSA']
        st.success("Data Anomali PSA berhasil dimuat.")
    else:
        st.warning("File Anomali Pinjaman.xlsx tidak memiliki sheet 'Anomali PSA' atau belum diunggah.")
        
    if 'Anomali Pinjaman.xlsx_Anomali ARTA' in dfs:
        df_anomali_arta = dfs['Anomali Pinjaman.xlsx_Anomali ARTA']
        st.success("Data Anomali ARTA berhasil dimuat.")
    else:
        st.warning("File Anomali Pinjaman.xlsx tidak memiliki sheet 'Anomali ARTA' atau belum diunggah.")
        
    if 'Anomali Pinjaman.xlsx_Anomali PRR' in dfs:
        df_anomali_prr = dfs['Anomali Pinjaman.xlsx_Anomali PRR']
        st.success("Data Anomali PRR berhasil dimuat.")
    else:
        st.warning("File Anomali Pinjaman.xlsx tidak memiliki sheet 'Anomali PRR' atau belum diunggah.")
        
    if 'Anomali Pinjaman.xlsx_Anomali PTN' in dfs:
        df_anomali_ptn = dfs['Anomali Pinjaman.xlsx_Anomali PTN']
        st.success("Data Anomali PTN berhasil dimuat.")
    else:
        st.warning("File Anomali Pinjaman.xlsx tidak memiliki sheet 'Anomali PTN' atau belum diunggah.")
        
    if 'Anomali Simpanan.xlsx_Sihara' in dfs:
        df_sihara = dfs['Anomali Simpanan.xlsx_Sihara']
        st.success("Data Sihara berhasil dimuat.")
    else:
        st.error("File Anomali Simpanan.xlsx tidak ditemukan atau tidak memiliki sheet 'Sihara'.")
        
    if 'Anomali Simpanan.xlsx_Pensiun' in dfs:
        df_pensiun = dfs['Anomali Simpanan.xlsx_Pensiun']
        st.success("Data Pensiun berhasil dimuat.")
    else:
        st.warning("File Anomali Simpanan.xlsx tidak memiliki sheet 'Pensiun' atau belum diunggah.")
        
    if 'Anomali Simpanan.xlsx_Sukarela' in dfs:
        df_sukarela = dfs['Anomali Simpanan.xlsx_Sukarela']
        st.success("Data Sukarela berhasil dimuat.")
    else:
        st.warning("File Anomali Simpanan.xlsx tidak memiliki sheet 'Sukarela' atau belum diunggah.")

    if 'DbSimpanan.xlsx_IA_SimpananDB' in dfs:
        df_DbSimpanan = dfs['DbSimpanan.xlsx_IA_SimpananDB']
        st.success("Data DbSImpanan berhasil dimuat.")
    else:
        st.warning("File DbSimpanan.xlsx tidak memiliki sheet 'IA_SimpananDB' atau belum diunggah.")


##----Nama Dataframe
# df_anomali_pu
# df_anomali_pmb
# df_anomali_psa
# df_anomali_prr
# df_anomali_ptn
# df_anomali_arta
# df_anomali_dtp
# df_sukarela
# df_pensiun
# df_sihara
# df_DbSimpanan

df_pu = df_anomali_pu[['ID','CEK KRITERIA']]
rename_dict = {
    'CEK KRITERIA':'PU'
}
df_pu = df_pu.rename(columns=rename_dict)
df_pu['PU'] = df_pu['PU'].replace({True: 0, False: 1})
#-------------------------------------------------------------------------------
df_pmb = df_anomali_pmb[['ID','CEK KRITERIA']]
rename_dict = {
    'CEK KRITERIA':'PMB'
}
df_pmb = df_pmb.rename(columns=rename_dict)
df_pmb['PMB'] = df_pmb['PMB'].replace({True: 0, False: 1})
#-------------------------------------------------------------------------------
df_psa = df_anomali_psa[['ID','CEK KRITERIA','Sukarela Sesuai','Wajib Sesuai','Pensiun Sesuai']]

df_psa[['CEK KRITERIA', 'Sukarela Sesuai', 'Wajib Sesuai', 'Pensiun Sesuai']] = df_psa[['CEK KRITERIA', 'Sukarela Sesuai', 'Wajib Sesuai', 'Pensiun Sesuai']].replace({True: 0, False: 1})

df_psa['TOTAL'] = df_psa[['CEK KRITERIA', 'Sukarela Sesuai', 'Wajib Sesuai', 'Pensiun Sesuai']].sum(axis=1)

rename_dict = {
    'TOTAL':'PSA'
}
df_psa = df_psa.rename(columns=rename_dict)

df_psa = df_psa[['ID','PSA']]
df_psa['PSA'] = df_psa['PSA'].replace({True: 0, False: 1})
#-------------------------------------------------------------------------------
df_prr = df_anomali_prr[['ID','CEK KRITERIA','Sukarela Sesuai','Wajib Sesuai','Pensiun Sesuai']]

df_prr[['CEK KRITERIA', 'Sukarela Sesuai', 'Wajib Sesuai', 'Pensiun Sesuai']] = df_prr[['CEK KRITERIA', 'Sukarela Sesuai', 'Wajib Sesuai', 'Pensiun Sesuai']].replace({True: 0, False: 1})

df_prr['TOTAL'] = df_prr[['CEK KRITERIA', 'Sukarela Sesuai', 'Wajib Sesuai', 'Pensiun Sesuai']].sum(axis=1)

rename_dict = {
    'TOTAL':'PRR'
}
df_prr = df_prr.rename(columns=rename_dict)

df_prr = df_prr[['ID','PRR']]
df_prr['PRR'] = df_prr['PRR'].replace({True: 0, False: 1})
#-------------------------------------------------------------------------------
df_ptn = df_anomali_ptn[['ID','SEMUA KRITERIA TERPENUHI']]

rename_dict = {
    'SEMUA KRITERIA TERPENUHI':'PTN'
    }
df_ptn = df_ptn.rename(columns=rename_dict)
df_ptn['PTN'] = df_ptn['PTN'].replace({True: 0, False: 1})
#-------------------------------------------------------------------------------
df_arta = df_anomali_arta[['ID','CEK KRITERIA']]

rename_dict = {
    'CEK KRITERIA':'ARTA'
    }
df_arta = df_arta.rename(columns=rename_dict)
df_arta['ARTA'] = df_arta['ARTA'].replace({True: 0, False: 1})
#-------------------------------------------------------------------------------
df_dtp = df_anomali_dtp[['ID','CEK KRITERIA']]

rename_dict = {
    'CEK KRITERIA':'DTP'
    }
df_dtp = df_dtp.rename(columns=rename_dict)
df_dtp['DTP'] = df_dtp['DTP'].replace({True: 0, False: 1})
################################################################################
# Sukarela
df_sukarela = df_sukarela[['ID','Transaksi > 0 & ≠ Modus Sukarela']]
rename_dict = {
    'Transaksi > 0 & ≠ Modus Sukarela':'SUKARELA'
    }
df_sukarela = df_sukarela.rename(columns=rename_dict)

# Sihara
df_sihara = df_sihara[['ID','TRANSAKSI TIDAK SESUAI']]
rename_dict = {
    'TRANSAKSI TIDAK SESUAI':'HARI RAYA'
    }
df_sihara = df_sihara.rename(columns=rename_dict)

# Pensiun
df_pensiun = df_pensiun[['ID','Anomali']]
rename_dict = {
    'Anomali':'PENSIUN'
    }
df_pensiun = df_pensiun.rename(columns=rename_dict)
#-------------------------------------------------------------------------------
df_selected_all = df_DbSimpanan[['Client ID','Client Name','Center ID','Group ID']]
rename_dict = {
    'Client ID':'ID',
    'Client Name': 'Nama',
    'Center ID': 'Center',
    'Group ID': 'Kelompok'
    }
df_selected_all = df_selected_all.rename(columns=rename_dict)


df_selected_all = df_selected_all.merge(df_sukarela[['ID', 'SUKARELA']], on='ID', how='left')
df_selected_all = df_selected_all.merge(df_pensiun[['ID', 'PENSIUN']], on='ID', how='left')
df_selected_all = df_selected_all.merge(df_sihara[['ID', 'HARI RAYA']], on='ID', how='left')
df_selected_all = df_selected_all.merge(df_pu[['ID', 'PU']], on='ID', how='left')
df_selected_all = df_selected_all.merge(df_pmb[['ID', 'PMB']], on='ID', how='left')
df_selected_all = df_selected_all.merge(df_psa[['ID', 'PSA']], on='ID', how='left')
df_selected_all = df_selected_all.merge(df_prr[['ID', 'PRR']], on='ID', how='left')
df_selected_all = df_selected_all.merge(df_ptn[['ID', 'PTN']], on='ID', how='left')
df_selected_all = df_selected_all.merge(df_arta[['ID', 'ARTA']], on='ID', how='left')
df_selected_all = df_selected_all.merge(df_dtp[['ID', 'DTP']], on='ID', how='left')


df_selected_all = df_selected_all.fillna(0)


anomali_columns = ['SUKARELA', 'PENSIUN', 'HARI RAYA', 'PU', 'PMB', 'PSA', 'PRR', 'PTN', 'ARTA', 'DTP']
df_selected_all['TOTAL ANOMALI'] = df_selected_all[anomali_columns].sum(axis=1)

df_selected_all = df_selected_all[['ID', 'Nama', 'Center', 'Kelompok', 'SUKARELA', 'PENSIUN', 'HARI RAYA', 'PU', 'PMB', 'PSA', 'PRR', 'PTN', 'ARTA', 'DTP', 'TOTAL ANOMALI']]
df_selected_all = df_selected_all.drop_duplicates(subset=['ID', 'Nama'])

st.write("Data setelah diproses:")
st.write(df_selected_all)
