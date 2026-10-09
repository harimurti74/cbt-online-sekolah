# Tambahkan ini di jajaran import paling atas
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

import streamlit as st
st.markdown("""
    <style>
    /* Sembunyikan Header Atas & Menu Utama */
    header, footer, #MainMenu, [data-testid="stHeader"] {
        display: none !important;
        visibility: hidden !important;
    }
    
    /* Sembunyikan Badge, Tombol Deploy, & Floating Menu Kelola di Pojok Bawah */
    div[data-testid="stAppDeployButton"],
    div[data-testid="stStatusWidget"],
    div[class*="viewerBadge"],
    div[class*="styles_viewerBadge"],
    .stAppDeployButton,
    a[href*="streamlit.io"],
    a[href*="streamlit.app"] {
        display: none !important;
        visibility: hidden !important;
    }
    </style>
""", unsafe_allow_html=True)
import pandas as pd
import numpy as np
import json
import time
import re
import io
import difflib
from datetime import datetime

# Set page config
st.set_page_config(
    page_title="Sistem Ujian Online CBT - TKA",
    page_icon="📝",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ---------------------------------------------------------
# INITIALIZE SESSION STATE DATA (DATABASE SIMULATION)
# ---------------------------------------------------------
if 'db' not in st.session_state:
    st.session_state.db = {
        'identitas': {
            'nama_sekolah': 'SMA Negeri 1 Nusantara',
            'nama_guru': 'Bapak/Ibu Guru',
            'kelas': 'XII MIPA 1',
            'mapel': 'Biologi / Matematika',
            'judul_ujian': 'Asesmen Sumatif / Ujian TKA',
            'durasi_menit': 60,
            'token_ujian': 'TKA2026',
            'folder_id_gdrive': '',
            'wajib_pilih_daftar_siswa': False,
            'max_pelanggaran': 3
        },
        'soal_list': [],
        'daftar_siswa': [],  # List of dicts: {'no_absen': '01', 'nama': 'Ahmad'}
        'sesi_siswa': {},    # key: id_siswa -> {'status': 'Belum'/'Sedang'/'Selesai', 'jawaban': {}, 'ragu': {}, 'pelanggaran': 0, 'waktu_mulai': None, 'nilai_total': 0, 'detail_skor': {}}
        'log_pelanggaran': []
    }

if 'role' not in st.session_state:
    st.session_state.role = 'Admin'  # 'Admin' or 'Siswa'

if 'current_siswa' not in st.session_state:
    st.session_state.current_siswa = None

if 'no_soal_aktif' not in st.session_state:
    st.session_state.no_soal_aktif = 0

# ---------------------------------------------------------
# HELPER FUNCTIONS & SCORING LOGIC
# ---------------------------------------------------------
def hitung_nilai_uraian(jawaban_siswa, kunci_jawaban):
    if not jawaban_siswa or not kunci_jawaban:
        return 0.0
    teks_siswa = str(jawaban_siswa).strip().lower()
    teks_kunci = str(kunci_jawaban).strip().lower()
    
    # Check exact keywords or similarity ratio
    ratio = difflib.SequenceMatcher(None, teks_siswa, teks_kunci).ratio()
    persentase = ratio * 100.0
    
    if persentase < 80.0:
        return 0.0
    else:
        # Formulasi: 80% -> 0, 81% -> 1.0, 81.5% -> 1.5, 82% -> 2.0, dst.
        nilai = persentase - 80.0
        return min(10.0, round(nilai, 2))

def parse_kunci_tepat_tidak_tepat(kunci_str):
    # Format Kunci: "A:Tidak Tepat; B:Tepat; C:Tidak Tepat"
    kunci_dict = {}
    items = str(kunci_str).split(';')
    for it in items:
        if ':' in it:
            k, v = it.split(':', 1)
            kunci_dict[k.strip().upper()] = v.strip().title()
    return kunci_dict

def hitung_skor_siswa(id_siswa):
    siswa_data = st.session_state.db['sesi_siswa'].get(id_siswa, {})
    jawaban_siswa = siswa_data.get('jawaban', {})
    soal_list = st.session_state.db['soal_list']
    
    skor_total = 0.0
    detail = {}
    
    for idx, s in enumerate(soal_list):
        tipe = s.get('Tipe_Soal', 'PG_STANDAR')
        kunci = s.get('Kunci_Jawaban', '')
        j_siswa = jawaban_siswa.get(idx, None)
        bobot = float(s.get('Bobot_Per_Pernyataan', 1))
        
        skor_soal = 0.0
        
        if tipe in ['PG_STANDAR', 'PG_KOMPLEKS_ASOSIASI', 'SEBAB_AKIBAT']:
            if j_siswa and str(j_siswa).strip().upper() == str(kunci).strip().upper():
                skor_soal = bobot
        elif tipe == 'TEPAT_TIDAK_TEPAT':
            # j_siswa is dict: {'A': 'Tepat', 'B': 'Tidak Tepat'}
            kunci_dict = parse_kunci_tepat_tidak_tepat(kunci)
            if isinstance(j_siswa, dict):
                for k_item, v_kunci in kunci_dict.items():
                    v_siswa = j_siswa.get(k_item, '')
                    if v_siswa and v_siswa.strip().lower() == v_kunci.strip().lower():
                        skor_soal += bobot
        elif tipe == 'URAIAN_SINGKAT':
            skor_soal = hitung_nilai_uraian(j_siswa, kunci)
            
        skor_total += skor_soal
        detail[f"Soal_{idx+1}"] = skor_soal
        
    st.session_state.db['sesi_siswa'][id_siswa]['nilai_total'] = round(skor_total, 2)
    st.session_state.db['sesi_siswa'][id_siswa]['detail_skor'] = detail
    return skor_total

def upload_ke_gdrive_guru(file_buffer, file_name, folder_id):
    """Mengunggah file Excel ke folder Google Drive pribadi milik Guru"""
    try:
        # Mengambil kunci Service Account dari Secrets Streamlit
        creds = service_account.Credentials.from_service_account_info(
            st.secrets["gcp_service_account"],
            scopes=['https://www.googleapis.com/auth/drive']
        )
        service = build('drive', 'v3', credentials=creds)
        
        file_metadata = {
            'name': file_name,
            'parents': [folder_id]  # Folder ID milik guru
        }
        
        media = MediaIoBaseUpload(
            file_buffer, 
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
            resumable=True
        )
        
        uploaded_file = service.files().create(
            body=file_metadata, 
            media_body=media, 
            fields='id'
        ).execute()
        
        return True, uploaded_file.get('id')
    except Exception as e:
        return False, str(e)

# ---------------------------------------------------------
# NAVIGATION & SIDEBAR
# ---------------------------------------------------------
st.sidebar.title("📌 Menu Navigasi")
mode = st.sidebar.radio("Pilih Mode Akses:", ["Panel Guru (Admin CBT)", "Halaman Ujian Siswa"])

# Query param check for student direct token link
query_params = st.query_params
if 'token' in query_params:
    st.session_state.db['identitas']['token_ujian'] = query_params['token']

# ---------------------------------------------------------
# MODE 1: PANEL GURU / ADMIN CBT
# ---------------------------------------------------------
if mode == "Panel Guru (Admin CBT)":
    st.title("⚙️ Panel Kelola & Pemantauan CBT Guru")
    st.caption("Kelola Identitas Ujian, Upload Soal, Peserta, Pemantauan Real-time, dan Reset Ujian")
    
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "1. Setting Ujian & Link", 
        "2. Upload Soal & Peserta", 
        "3. Pemantauan Real-Time", 
        "4. Reset Ujian", 
        "5. Rekapitulasi Hasil"
    ])
    
    # TAB 1: SETTING UJIAN & LINK UNIK
    with tab1:
        st.subheader("📋 Identitas Pelaksanaan Ujian")
        col1, col2 = st.columns(2)
        with col1:
            st.session_state.db['identitas']['nama_sekolah'] = st.text_input("Nama Sekolah / Instansi:", st.session_state.db['identitas']['nama_sekolah'])
            st.session_state.db['identitas']['nama_guru'] = st.text_input("Nama Guru / Pengampu:", st.session_state.db['identitas']['nama_guru'])
            st.session_state.db['identitas']['kelas'] = st.text_input("Kelas / Rombel:", st.session_state.db['identitas']['kelas'])
        with col2:
            st.session_state.db['identitas']['mapel'] = st.text_input("Mata Pelajaran:", st.session_state.db['identitas']['mapel'])
            st.session_state.db['identitas']['judul_ujian'] = st.text_input("Judul Asesmen / Ujian:", st.session_state.db['identitas']['judul_ujian'])
            st.session_state.db['identitas']['durasi_menit'] = st.number_input("Durasi Ujian (Menit):", min_value=1, value=int(st.session_state.db['identitas']['durasi_menit']))
            
        st.markdown("---")
        st.subheader("🔐 Pengaturan Akses & Keamanan Siswa")
        col_a, col_b = st.columns(2)
        with col_a:
            st.session_state.db['identitas']['token_ujian'] = st.text_input("Token Akses Ujian:", st.session_state.db['identitas']['token_ujian'])
            st.session_state.db['identitas']['max_pelanggaran'] = st.slider("Maksimal Toleransi Kecurangan (Minimize/Tab Switch):", 1, 10, int(st.session_state.db['identitas']['max_pelanggaran']))
        with col_b:
            st.session_state.db['identitas']['wajib_pilih_daftar_siswa'] = st.checkbox(
                "Siswa Wajib Pilih dari Daftar Peserta (Tidak Ketik Manual)", 
                value=st.session_state.db['identitas']['wajib_pilih_daftar_siswa']
            )
         # Masukkan kode ini di dalam "with tab1:"
        st.subheader("📁 Pengaturan Penyimpanan Google Drive")
        st.session_state.db['identitas']['folder_id_gdrive'] = st.text_input(
        "Folder ID Google Drive Guru:",
        value=st.session_state.db['identitas'].get('folder_id_gdrive', ''),
        placeholder="Contoh: 1a2B3c4D5e6F7g8H9i0J_kLmnOpQrStUv",
        help="Masukkan kode unik folder Google Drive Anda. Hasil ujian siswa akan otomatis diunggah ke folder ini."
)   
        st.success(f"🔗 **Link Unik Ujian Siswa**: `https://cbt-ujian.sekolah.sch.id/?token={st.session_state.db['identitas']['token_ujian']}`")
        st.info("Bagikan link atau Token di atas kepada siswa untuk memasuki ruang ujian.")

    # TAB 2: UPLOAD SOAL & DAFTAR SISWA
    with tab2:
        st.subheader("📤 Upload File Bank Soal (Excel / CSV)")
        file_soal = st.file_uploader("Unggah File Template Soal (Format XLSX / CSV):", type=['xlsx', 'xls', 'csv'])
        if file_soal is not None:
            try:
                if file_soal.name.endswith('.csv'):
                    df_soal = pd.read_csv(file_soal)
                else:
                    df_soal = pd.read_excel(file_soal)
                st.session_state.db['soal_list'] = df_soal.to_dict('records')
                st.success(f"✅ Berhasil mengimpor {len(st.session_state.db['soal_list'])} soal ke dalam sistem!")
                st.dataframe(df_soal.head(5))
            except Exception as e:
                st.error(f"Gagal membaca file soal: {e}")
                
        st.markdown("---")
        st.subheader("👥 Upload / Isil Daftar Peserta Ujian (Opsional)")
        file_siswa = st.file_uploader("Unggah File Daftar Siswa (Kolom: No_Absen, Nama):", type=['xlsx', 'csv'])
        if file_siswa is not None:
            try:
                if file_siswa.name.endswith('.csv'):
                    df_siswa = pd.read_csv(file_siswa)
                else:
                    df_siswa = pd.read_excel(file_siswa)
                st.session_state.db['daftar_siswa'] = df_siswa.to_dict('records')
                st.success(f"✅ Berhasil memuat {len(st.session_state.db['daftar_siswa'])} nama siswa!")
                st.dataframe(df_siswa)
            except Exception as e:
                st.error(f"Gagal membaca daftar siswa: {e}")

    # TAB 3: PEMANTAUAN REAL-TIME (CBT MONITOR)
    with tab3:
        st.subheader("📊 Monitoring Kondisi Ujian Siswa")
        
        # Calculate stats
        total_peserta = len(st.session_state.db['daftar_siswa']) if st.session_state.db['daftar_siswa'] else len(st.session_state.db['sesi_siswa'])
        login_count = len(st.session_state.db['sesi_siswa'])
        
        sedang = sum(1 for s in st.session_state.db['sesi_siswa'].values() if s.get('status') == 'Sedang Mengerjakan')
        selesai = sum(1 for s in st.session_state.db['sesi_siswa'].values() if s.get('status') == 'Selesai')
        belum = total_peserta - login_count if total_peserta >= login_count else 0
        
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Siswa Login", f"{login_count} Siswa")
        m2.metric("Sedang Mengerjakan", f"{sedang} Siswa", delta_color="normal")
        m3.metric("Selesai Ujian", f"{selesai} Siswa")
        m4.metric("Belum Mengerjakan/Login", f"{belum} Siswa")
        
        st.markdown("---")
        st.write("### 🔍 Tabel Status Detail Per Siswa")
        if st.session_state.db['sesi_siswa']:
            monitor_data = []
            for id_s, data in st.session_state.db['sesi_siswa'].items():
                soal_terjawab = len([j for j in data.get('jawaban', {}).values() if j is not None and j != ''])
                monitor_data.append({
                    'ID Siswa / No Absen': id_s,
                    'Nama Siswa': data.get('nama', 'Siswa'),
                    'Status Ujian': data.get('status', 'Belum'),
                    'Terjawab': f"{soal_terjawab} / {len(st.session_state.db['soal_list'])} Soal",
                    'Indikasi Kecurangan': f"{data.get('pelanggaran', 0)}x Warning",
                    'Waktu Mulai': data.get('waktu_mulai', '-'),
                    'Nilai Sementara': data.get('nilai_total', 0)
                })
            st.dataframe(pd.DataFrame(monitor_data), use_container_width=True)
        else:
            st.info("Belum ada siswa yang login ke ruang ujian.")

    # TAB 4: RESET UJIAN
    with tab4:
        st.subheader("🔄 Menu Reset Ujian (Per Siswa / Per Kelas)")
        st.warning("Gunakan fitur ini jika siswa mengalami kendala PC / mati listrik / eror jaringan saat ujian.")
        
        col_r1, col_r2 = st.columns(2)
        with col_r1:
            st.markdown("#### 👤 Reset Ujian Per Siswa")
            if st.session_state.db['sesi_siswa']:
                pilih_siswa_reset = st.selectbox("Pilih Siswa yang Ingin Direset:", list(st.session_state.db['sesi_siswa'].keys()))
                if st.button("🔴 Reset Ujian Siswa Ini", type="secondary"):
                    del st.session_state.db['sesi_siswa'][pilih_siswa_reset]
                    st.success(f"Sesi ujian siswa {pilih_siswa_reset} telah di-reset. Siswa dapat login kembali!")
                    st.rerun()
            else:
                st.write("Tidak ada sesi siswa aktif.")
                
        with col_r2:
            st.markdown("#### 🏫 Reset Seluruh Ujian (Per Kelas / Sesi)")
            if st.button("🚨 RESET SEMUA SISWA & SESI", type="primary"):
                st.session_state.db['sesi_siswa'] = {}
                st.success("Semua data sesi ujian siswa berhasil di-reset bersih!")
                st.rerun()

    # TAB 5: REKAPITULASI HASIL
    with tab5:
        st.subheader("📥 Unduh Rekapitulasi Nilai Ujian (Excel)")
        if st.session_state.db['sesi_siswa']:
            rekap_list = []
            for id_s, d in st.session_state.db['sesi_siswa'].items():
                row = {
                    'No_Absen': id_s,
                    'Nama_Siswa': d.get('nama', ''),
                    'Status': d.get('status', ''),
                    'Jumlah_Pelanggaran': d.get('pelanggaran', 0),
                    'Nilai_Total': d.get('nilai_total', 0)
                }
                # Add detail scores per question
                for q_k, q_v in d.get('detail_skor', {}).items():
                    row[q_k] = q_v
                rekap_list.append(row)
                
            df_rekap = pd.DataFrame(rekap_list)
            st.dataframe(df_rekap, use_container_width=True)
            
            # Export to Excel buffer
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_rekap.to_excel(writer, index=False, sheet_name='Hasil_Ujian')
            excel_data = output.getvalue()
            
            st.download_button(
                label="📥 Download Hasil Ujian (.xlsx)",
                data=excel_data,
                file_name=f"Hasil_Ujian_{st.session_state.db['identitas']['mapel']}_{st.session_state.db['identitas']['kelas']}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            if st.button("☁️️ Simpan Hasil Ujian Langsung ke Google Drive Saya"):
                folder_id = st.session_state.db['identitas'].get('folder_id_gdrive', '').strip()
                
                if not folder_id:
                    st.error("❌ Folder ID Google Drive belum diisi pada Tab 1!")
                else:
                    output.seek(0)  # Reset pointer buffer file
                    nama_file_excel = f"Hasil_Ujian_{st.session_state.db['identitas']['mapel']}_{st.session_state.db['identitas']['kelas']}.xlsx"
                    
                    with st.spinner("Sedang mengunggah hasil ujian ke Google Drive Anda..."):
                        berhasil, pesan = upload_ke_gdrive_guru(output, nama_file_excel, folder_id)
                        if berhasil:
                            st.success(f"✅ Berhasil disimpan ke Google Drive Anda! (File ID: {pesan})")
                        else:
                            st.error(f"❌ Gagal mengunggah: {pesan}")
                        
        else:
            st.info("Belum ada data nilai hasil ujian yang dapat diunduh.")


# ---------------------------------------------------------
# MODE 2: HALAMAN UJIAN SISWA (CBT INTERFACE ALA TKA)
# ---------------------------------------------------------
else:
    # Check if questions exist
    if not st.session_state.db['soal_list']:
        st.error("⚠️ Ujian belum siap! Bank soal belum diunggah oleh Guru/Admin CBT.")
        st.stop()
        
    identitas = st.session_state.db['identitas']
    
    # LOGIN SISWA
    if not st.session_state.current_siswa:
        st.title("📝 Portal Ujian Online CBT")
        st.markdown(f"### {identitas['nama_sekolah']}")
        st.write(f"**Mata Pelajaran**: {identitas['mapel']} | **Kelas**: {identitas['kelas']}")
        st.write(f"**Ujian**: {identitas['judul_ujian']} | **Durasi**: {identitas['durasi_menit']} Menit")
        st.markdown("---")
        
        with st.form("form_login_siswa"):
            st.subheader("🔑 Konfirmasi Identitas Peserta Ujian")
            
            token_input = st.text_input("Masukkan Token Ujian:", value=identitas['token_ujian'])
            
            if identitas['wajib_pilih_daftar_siswa'] and st.session_state.db['daftar_siswa']:
                opsi_siswa = [f"{s.get('No_Absen', idx+1)} - {s.get('Nama', '')}" for idx, s in enumerate(st.session_state.db['daftar_siswa'])]
                pilihan_s = st.selectbox("Pilih Nama Anda dari Daftar Peserta:", opsi_siswa)
                nama_siswa = pilihan_s.split(' - ', 1)[1] if ' - ' in pilihan_s else pilihan_s
                no_absen = pilihan_s.split(' - ', 1)[0]
            else:
                nama_siswa = st.text_input("Nama Lengkap Siswa:")
                no_absen = st.text_input("Nomor Absen / ID Peserta:")
                
            btn_login = st.form_submit_button("🚀 MASUK RUANG UJIAN")
            
            if btn_login:
                if token_input.strip() != identitas['token_ujian'].strip():
                    st.error("❌ Token Ujian Salah!")
                elif not nama_siswa or not no_absen:
                    st.error("❌ Harap isi Nama dan No Absen!")
                else:
                    id_siswa = str(no_absen).strip()
                    st.session_state.current_siswa = id_siswa
                    
                    if id_siswa not in st.session_state.db['sesi_siswa']:
                        st.session_state.db['sesi_siswa'][id_siswa] = {
                            'nama': nama_siswa,
                            'status': 'Sedang Mengerjakan',
                            'jawaban': {},
                            'ragu': {},
                            'pelanggaran': 0,
                            'waktu_mulai': datetime.now().strftime("%H:%M:%S"),
                            'nilai_total': 0,
                            'detail_skor': {}
                        }
                    st.success("Login Berhasil! Mengalihkan ke ruang ujian...")
                    st.rerun()

    # AREA KERJA UJIAN SISWA
    else:
        id_siswa = st.session_state.current_siswa
        sesi = st.session_state.db['sesi_siswa'].get(id_siswa)
        
        if not sesi:
            st.error("Sesi ujian tidak ditemukan atau telah di-reset oleh Admin.")
            st.session_state.current_siswa = None
            st.rerun()
            
        if sesi.get('status') == 'Selesai':
            st.balloons()
            st.title("🎉 Ujian Telah Selesai!")
            st.success(f"Terima kasih **{sesi['nama']}**! Jawaban Anda telah tersimpan dengan aman.")
            st.write(f"**Nilai Akhir Anda**: {sesi.get('nilai_total', 0)}")
            if st.button("🚪 Keluar Ruang Ujian"):
                st.session_state.current_siswa = None
                st.rerun()
            st.stop()
            
        # ---------------------------------------------------------
        # ANTI-CHEAT JAVASCRIPT & FULLSCREEN DETECTOR (TKA SYSTEM)
        # ---------------------------------------------------------
        js_anti_cheat = f"""
        <script>
        // Disable Right Click & Copy Paste
        document.addEventListener('contextmenu', event => event.preventDefault());
        document.addEventListener('keydown', function(e) {{
            if (e.ctrlKey && (e.key === 'c' || e.key === 'v' || e.key === 'u' || e.key === 'a')) {{
                e.preventDefault();
            }}
        }});
        
        // Tab visibility switch detection
        document.addEventListener("visibilitychange", function() {{
            if (document.hidden) {{
                alert("⚠️ PERINGATAN KECURANGAN: Anda terdeteksi meninggalkan layar ujian! Kejadian ini dicatat oleh pengawas.");
            }}
        }});
        </script>
        """
        st.components.v1.html(js_anti_cheat, height=0)
        
        # CBT HEADER BAR
        st.markdown(f"""
        <div style="background-color: #1F4E78; padding: 12px 20px; border-radius: 8px; color: white; margin-bottom: 15px;">
            <div style="display: flex; justify-content: space-between; align-items: center;">
                <div>
                    <h3 style="margin: 0; color: white;">{identitas['nama_sekolah']} - {identitas['mapel']}</h3>
                    <small>{identitas['judul_ujian']} | Siswa: <b>{sesi['nama']} ({id_siswa})</b></small>
                </div>
                <div style="text-align: right; background-color: #d9534f; padding: 6px 15px; border-radius: 5px;">
                    <span style="font-size: 12px; display:block;">SISA WAKTU</span>
                    <strong style="font-size: 18px;">⏱️ {identitas['durasi_menit']}:00 Min</strong>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        # MAIN LAYOUT (SOAL vs NAVIGASI KISI-KISI)
        col_left, col_right = st.columns([3, 1])
        
        soal_list = st.session_state.db['soal_list']
        total_soal = len(soal_list)
        curr_idx = st.session_state.no_soal_aktif
        soal = soal_list[curr_idx]
        tipe = soal.get('Tipe_Soal', 'PG_STANDAR')
        
        with col_left:
            st.markdown(f"#### Soal Nomor {curr_idx + 1} / {total_soal}  `[{tipe}]`")
            
            # STIMULUS TEKS & MEDIA GAMBAR
            if pd.notna(soal.get('Media_Gambar')) and str(soal.get('Media_Gambar')).strip() != '':
                st.info(f"🖼️ Media Gambar / Grafik Attached: `{soal.get('Media_Gambar')}`")
                
            if pd.notna(soal.get('Stimulus_Teks')) and str(soal.get('Stimulus_Teks')).strip() != '':
                st.markdown(f"""
                <div style="background-color: #F8F9FA; padding: 15px; border-left: 4px solid #1F4E78; margin-bottom: 15px; border-radius: 4px;">
                    {soal.get('Stimulus_Teks')}
                </div>
                """, unsafe_allow_html=True)
                
            # SOAL UTAMA
            st.write(f"**{soal.get('Soal_Utama', '')}**")
            
            # INPUT JAWABAN SESUAI TIPE SOAL
            curr_ans = sesi['jawaban'].get(curr_idx, None)
            
            # TIPE 1: PG STANDAR / PG KOMPLEKS ASOSIASI / SEBAB AKIBAT
            if tipe in ['PG_STANDAR', 'PG_KOMPLEKS_ASOSIASI', 'SEBAB_AKIBAT']:
                # Print Pernyataan if any
                if tipe == 'PG_KOMPLEKS_ASOSIASI':
                    st.write("---")
                    st.write("Daftar Pernyataan:")
                    for i in range(1, 5):
                        p_val = soal.get(f'Pernyataan_{i}') or soal.get(f'Pernyataan_{i}_atau_Sebab') or soal.get(f'Pernyataan_{i}_atau_Akibat')
                        if pd.notna(p_val) and str(p_val).strip() != '':
                            st.write(f"• {p_val}")
                elif tipe == 'SEBAB_AKIBAT':
                    st.write("---")
                    st.write(f"**Pernyataan**: {soal.get('Pernyataan_1_atau_Sebab', '')}")
                    st.write(f"**Alasan (SEBAB)**: {soal.get('Pernyataan_2_atau_Akibat', '')}")
                    
                opsi_choices = ['A', 'B', 'C', 'D', 'E']
                opsi_labels = {
                    'A': f"A. {soal.get('Opsi_A', '')}",
                    'B': f"B. {soal.get('Opsi_B', '')}",
                    'C': f"C. {soal.get('Opsi_C', '')}",
                    'D': f"D. {soal.get('Opsi_D', '')}",
                    'E': f"E. {soal.get('Opsi_E', '')}"
                }
                
                def_index = opsi_choices.index(curr_ans) if curr_ans in opsi_choices else None
                jawaban_user = st.radio(
                    "Pilih Jawaban Anda (Bisa Klik Opsi / Tekan Keyboard A, B, C, D, E):",
                    opsi_choices,
                    format_func=lambda x: opsi_labels[x],
                    index=def_index,
                    key=f"radio_soal_{curr_idx}"
                )
                if jawaban_user:
                    sesi['jawaban'][curr_idx] = jawaban_user

            # TIPE 2: TEPAT / TIDAK TEPAT
            elif tipe == 'TEPAT_TIDAK_TEPAT':
                st.write("Tentukan Tepat / Tidak Tepat untuk setiap pernyataan berikut:")
                list_items = ['A', 'B', 'C', 'D']
                if not isinstance(curr_ans, dict):
                    curr_ans = {}
                    
                ans_dict = {}
                for it in list_items:
                    p_text = soal.get(f'Pernyataan_{list_items.index(it)+1}') or soal.get(f'Pernyataan_{list_items.index(it)+1}_atau_Sebab')
                    if pd.notna(p_text) and str(p_text).strip() != '':
                        c_val = curr_ans.get(it, None)
                        idx_c = ['Tepat', 'Tidak Tepat'].index(c_val) if c_val in ['Tepat', 'Tidak Tepat'] else None
                        ans_dict[it] = st.radio(
                            f"**{p_text}**",
                            ['Tepat', 'Tidak Tepat'],
                            index=idx_c,
                            key=f"radio_tepat_{curr_idx}_{it}"
                        )
                sesi['jawaban'][curr_idx] = ans_dict

            # TIPE 3: URAIAN SINGKAT
            elif tipe == 'URAIAN_SINGKAT':
                val_uraian = str(curr_ans) if curr_ans is not None else ""
                ans_uraian = st.text_area("Ketik Jawaban Singkat Anda di Sini:", value=val_uraian, key=f"uraian_{curr_idx}")
                sesi['jawaban'][curr_idx] = ans_uraian

            # TOMBOL NAVIGASI SOAL & RAGU-RAGU
            st.markdown("---")
            nav_col1, nav_col2, nav_col3 = st.columns([1, 1, 1])
            with nav_col1:
                if st.button("⬅️ Soal Sebelumnya", disabled=(curr_idx == 0)):
                    st.session_state.no_soal_aktif -= 1
                    st.rerun()
            with nav_col2:
                is_ragu = sesi['ragu'].get(curr_idx, False)
                if st.checkbox("🟨 Ragu-Ragu", value=is_ragu, key=f"ragu_{curr_idx}"):
                    sesi['ragu'][curr_idx] = True
                else:
                    sesi['ragu'][curr_idx] = False
            with nav_col3:
                if curr_idx < total_soal - 1:
                    if st.button("Soal Berikutnya ➡️", type="primary"):
                        st.session_state.no_soal_aktif += 1
                        st.rerun()
                else:
                    if st.button("🏁 SELESAI & SIMPAN UJIAN", type="primary"):
                        sesi['status'] = 'Selesai'
                        hitung_skor_siswa(id_siswa)
                        st.rerun()

        # KISI-KISI NOMOR SOAL (SIDEBAR KANAN)
        with col_right:
            st.markdown("### 🔲 Kisi-Kisi Soal")
            cols_grid = st.columns(4)
            for i in range(total_soal):
                # Determine button style / color
                is_answered = (sesi['jawaban'].get(i) is not None and sesi['jawaban'].get(i) != '')
                is_ragu = sesi['ragu'].get(i, False)
                is_current = (i == curr_idx)
                
                label_num = f"{i+1}"
                if is_ragu:
                    btn_type = "secondary" # Yellow style
                    label_num += " 🟨"
                elif is_answered:
                    btn_type = "primary"   # Green / Primary style
                    label_num += " 🟢"
                else:
                    btn_type = "secondary" # Default
                    
                with cols_grid[i % 4]:
                    if st.button(label_num, key=f"grid_btn_{i}"):
                        st.session_state.no_soal_aktif = i
                        st.rerun()
                        
            st.markdown("""
            <div style="font-size: 11px; margin-top: 15px;">
                🟢 : Sudah Dijawab<br>
                🟨 : Ragu-Ragu<br>
                ⚪ : Belum Dijawab
            </div>
            """, unsafe_allow_html=True)
