# Tambahkan ini di jajaran import paling atas
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

import streamlit as st
import pandas as pd
import numpy as np
import json
import time
import re
import io
import os
import difflib
import hashlib
from datetime import datetime

# Safe import for python-docx
try:
    import docx
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

# Streamlit Page Config - Wide Layout & Collapsed Sidebar for Fullscreen feel
st.set_page_config(
    page_title="Sistem Ujian Online CBT - TKA",
    page_icon="📝",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ---------------------------------------------------------
# PERSISTENT LOCAL FILE DATABASE
# ---------------------------------------------------------
DB_FILE = "cbt_persistent_db.json"

def load_persistent_db():
    if os.path.exists(DB_FILE):
        try:
            with open(DB_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def save_persistent_db(db_data):
    try:
        with open(DB_FILE, "w", encoding="utf-8") as f:
            json.dump(db_data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print("Save error:", e)

@st.cache_resource
def get_global_store():
    data = load_persistent_db()
    if 'sessions' not in data:
        data['sessions'] = {}
    return data

global_store = get_global_store()

def save_current_state():
    save_persistent_db(global_store)

def get_or_create_exam_session(session_code):
    session_code = str(session_code).strip().upper()
    if not session_code:
        session_code = "CBT-UTAMA"
        
    if session_code not in global_store['sessions']:
        global_store['sessions'][session_code] = {
            'identitas': {
                'kode_sesi': session_code,
                'nama_sekolah': 'SMA Negeri 1 Nusantara',
                'nama_guru': 'Bapak/Ibu Guru',
                'kelas': 'XII MIPA 1',
                'judul_ujian': 'Asesmen Sumatif / Ujian TKA',
                'durasi_menit': 60,
                'token_mode': 'Dinamis (15 Menit)',
                'token_statis': 'TKA2026',
                'token_secret_salt': f'SALT_{session_code}',
                'base_url_app': 'https://cbt-sekolah.streamlit.app',
                'folder_id_gdrive': '',
                'wajib_kroscek_peserta': False,
                'max_pelanggaran': 3
            },
            'bank_soal_mapel': {},
            'daftar_siswa': [],
            'sesi_siswa': {},
            'log_pelanggaran': []
        }
        save_current_state()
    return session_code, global_store['sessions'][session_code]

# Query parameters handling
query_params = st.query_params
url_session = query_params.get('session', query_params.get('kode', 'CBT-UTAMA')).strip().upper()
url_token = query_params.get('token', '').strip().upper()
url_mode = query_params.get('mode', '').strip().lower()

is_student_link = (url_mode == 'siswa') or (bool(url_token) and url_mode != 'admin')

# ---------------------------------------------------------
# GLOBAL CSS STYLING (Aplikasi TKA, Font Arial, Fullscreen)
# ---------------------------------------------------------
custom_css = """
<style>
/* Font Global Arial & Fullscreen Layout */
html, body, [class*="css"], .stApp {
    font-family: Arial, Helvetica, sans-serif !important;
    background-color: #F4F6F9 !important;
}

/* Hilangkan Padding Atas/Samping Bawaan Streamlit agar Layar Proporsional */
.block-container {
    padding-top: 1rem !important;
    padding-bottom: 1rem !important;
    padding-left: 2rem !important;
    padding-right: 2rem !important;
    max-width: 100% !important;
}

/* Sembunyikan Branding & Top Header */
header, footer, #MainMenu, [data-testid="stHeader"] {
    display: none !important;
    visibility: hidden !important;
}

/* Sembunyikan Toolbar & Badge Streamlit */
[data-testid="stToolbar"],
[data-testid="stDecoration"],
[data-testid="stStatusWidget"],
[data-testid="stAppDeployButton"],
[data-testid="stViewerBadge"],
div[class*="viewerBadge"],
div[class*="styles_viewerBadge"],
a[href*="streamlit.io"],
a[href*="streamlit.app"] {
    display: none !important;
    visibility: hidden !important;
}

/* Ukuran Font Khusus TKA */
.tka-instansi-title {
    font-family: Arial, sans-serif !important;
    font-size: 18pt !important;
    font-weight: bold !important;
    color: #FFFFFF !important;
    margin: 0 !important;
    padding: 0 !important;
}

.tka-exam-sub {
    font-family: Arial, sans-serif !important;
    font-size: 11pt !important;
    color: #E2E8F0 !important;
}

.tka-stimulus-box {
    font-family: Arial, sans-serif !important;
    font-size: 12pt !important;
    line-height: 1.6 !important;
    background-color: #FFFFFF !important;
    border: 1px solid #CBD5E1 !important;
    border-left: 5px solid #1E3A8A !important;
    padding: 15px !important;
    border-radius: 6px !important;
    margin-bottom: 15px !important;
    color: #1E293B !important;
}

.tka-soal-text {
    font-family: Arial, sans-serif !important;
    font-size: 12pt !important;
    font-weight: bold !important;
    line-height: 1.6 !important;
    color: #0F172A !important;
    margin-bottom: 15px !important;
}

/* Label Radio / Jawaban Arial 12pt */
[data-testid="stMarkdownContainer"] p {
    font-family: Arial, sans-serif !important;
    font-size: 12pt !important;
    color: #1E293B !important;
}

div[class*="stRadio"] label p {
    font-family: Arial, sans-serif !important;
    font-size: 12pt !important;
    font-weight: normal !important;
}

/* Tombol Navigasi Bawah (Arial 10pt) */
.stButton > button {
    font-family: Arial, sans-serif !important;
    font-size: 10pt !important;
    font-weight: bold !important;
    border-radius: 4px !important;
}

/* Kotak Kisi-Kisi Nomor Soal sebelah Kanan (Arial 10pt) */
div[data-testid="column"] button {
    font-family: Arial, sans-serif !important;
    font-size: 10pt !important;
    padding: 4px 8px !important;
}
</style>
"""

st.markdown(custom_css, unsafe_allow_html=True)

if is_student_link:
    # Sembunyikan total sidebar untuk siswa
    st.markdown("""
    <style>
    [data-testid="stSidebar"], section[data-testid="stSidebar"] {
        display: none !important;
        width: 0px !important;
    }
    </style>
    """, unsafe_allow_html=True)

# Tentukan mode akses
if is_student_link:
    mode = "Halaman Ujian Siswa"
else:
    st.sidebar.title("📌 Navigasi CBT")
    mode = st.sidebar.radio(
        "Pilih Mode Akses:", 
        ["Panel Guru (Admin CBT)", "Halaman Ujian Siswa"],
        index=0
    )

if not is_student_link:
    st.sidebar.markdown("---")
    st.sidebar.subheader("🔑 Kode Sesi Ujian Guru")
    existing_sessions = list(global_store['sessions'].keys()) if global_store['sessions'] else ["CBT-UTAMA"]
    
    selected_session = st.sidebar.text_input(
        "Kode Unik Ujian Anda:",
        value=url_session if url_session else "CBT-UTAMA",
        help="Setiap guru memiliki Kode Unik Ujian sendiri (misal: MAT-AHMAD, IPA-RINA)."
    ).strip().upper()
    
    if not selected_session:
        selected_session = "CBT-UTAMA"
        
    active_session_code, current_db = get_or_create_exam_session(selected_session)
else:
    active_session_code, current_db = get_or_create_exam_session(url_session)

# ---------------------------------------------------------
# HELPER FUNCTIONS
# ---------------------------------------------------------
def generate_dynamic_token(secret_salt="TKA_SEKOLAH_SECRET"):
    interval_15m = int(time.time() // 900)
    hash_obj = hashlib.md5(f"{secret_salt}_{interval_15m}".encode())
    token = hash_obj.hexdigest()[:6].upper()
    
    seconds_remaining = 900 - (int(time.time()) % 900)
    mins = seconds_remaining // 60
    secs = seconds_remaining % 60
    return token, f"{mins:02d}:{secs:02d}"

def get_active_token(db_identitas):
    if db_identitas.get('token_mode') == 'Dinamis (15 Menit)':
        token, _ = generate_dynamic_token(db_identitas.get('token_secret_salt', 'TKA_SEKOLAH_SECRET'))
        return token
    else:
        return db_identitas.get('token_statis', 'TKA2026').strip().upper()

@st.cache_data(show_spinner=False)
def parse_word_soal_bytes(file_bytes):
    if not HAS_DOCX:
        st.error("❌ Library `python-docx` belum terinstall. Pastikan `python-docx` ada di `requirements.txt`!")
        return []
        
    doc = docx.Document(io.BytesIO(file_bytes))
    soal_list = []
    
    if len(doc.tables) > 0:
        for tbl in doc.tables:
            row_dict = {}
            for row in tbl.rows:
                if len(row.cells) >= 2:
                    key = row.cells[0].text.strip().lower()
                    val = row.cells[1].text.strip()
                    
                    if 'tipe' in key:
                        row_dict['Tipe_Soal'] = val
                    elif 'media' in key or 'gambar' in key:
                        row_dict['Media_Gambar'] = val if val != '-' else ''
                    elif 'stimulus' in key or 'petunjuk' in key:
                        row_dict['Stimulus_Teks'] = val
                    elif 'soal' in key:
                        row_dict['Soal_Utama'] = val
                    elif 'pernyataan' in key or 'sebab' in key:
                        lines = [l.strip() for l in val.split('\n') if l.strip()]
                        for idx_l, line in enumerate(lines[:4]):
                            row_dict[f'Pernyataan_{idx_l+1}'] = line
                            if idx_l == 0:
                                row_dict['Pernyataan_1_atau_Sebab'] = line
                            elif idx_l == 1:
                                row_dict['Pernyataan_2_atau_Akibat'] = line
                    elif 'opsi' in key:
                        lines = [l.strip() for l in val.split('\n') if l.strip()]
                        for line in lines:
                            if line.startswith('A.') or line.startswith('A '):
                                row_dict['Opsi_A'] = line.split('.', 1)[-1].strip()
                            elif line.startswith('B.') or line.startswith('B '):
                                row_dict['Opsi_B'] = line.split('.', 1)[-1].strip()
                            elif line.startswith('C.') or line.startswith('C '):
                                row_dict['Opsi_C'] = line.split('.', 1)[-1].strip()
                            elif line.startswith('D.') or line.startswith('D '):
                                row_dict['Opsi_D'] = line.split('.', 1)[-1].strip()
                            elif line.startswith('E.') or line.startswith('E '):
                                row_dict['Opsi_E'] = line.split('.', 1)[-1].strip()
                    elif 'kunci' in key:
                        row_dict['Kunci_Jawaban'] = val
                    elif 'bobot' in key:
                        row_dict['Bobot_Per_Pernyataan'] = val
                        
            if 'Soal_Utama' in row_dict or 'Tipe_Soal' in row_dict:
                for k in ['Tipe_Soal', 'Media_Gambar', 'Stimulus_Teks', 'Soal_Utama', 
                          'Pernyataan_1_atau_Sebab', 'Pernyataan_2_atau_Akibat', 'Pernyataan_3', 'Pernyataan_4',
                          'Opsi_A', 'Opsi_B', 'Opsi_C', 'Opsi_D', 'Opsi_E', 'Kunci_Jawaban', 'Bobot_Per_Pernyataan']:
                    if k not in row_dict:
                        row_dict[k] = ''
                soal_list.append(row_dict)
    return soal_list

def is_jawaban_terjawab(j):
    if j is None or j == '':
        return False
    if isinstance(j, dict):
        return any(v is not None and str(v).strip() != '' for v in j.values())
    return True

def hitung_nilai_uraian(jawaban_siswa, kunci_jawaban):
    if not jawaban_siswa or not kunci_jawaban:
        return 0.0
    teks_siswa = str(jawaban_siswa).strip().lower()
    teks_kunci = str(kunci_jawaban).strip().lower()
    
    ratio = difflib.SequenceMatcher(None, teks_siswa, teks_kunci).ratio()
    persentase = ratio * 100.0
    
    if persentase < 80.0:
        return 0.0
    else:
        nilai = persentase - 80.0
        return min(10.0, round(nilai, 2))

def parse_kunci_tepat_tidak_tepat(kunci_str):
    kunci_dict = {}
    items = str(kunci_str).split(';')
    for it in items:
        if ':' in it:
            k, v = it.split(':', 1)
            kunci_dict[k.strip().upper()] = v.strip().title()
    return kunci_dict

def hitung_skor_siswa(id_siswa, db_target):
    siswa_data = db_target['sesi_siswa'].get(id_siswa, {})
    jawaban_siswa = siswa_data.get('jawaban', {})
    mapel_siswa = siswa_data.get('mapel', '')
    soal_list = db_target['bank_soal_mapel'].get(mapel_siswa, [])
    
    skor_total = 0.0
    detail = {}
    
    for idx, s in enumerate(soal_list):
        tipe = s.get('Tipe_Soal', 'PG_STANDAR')
        kunci = s.get('Kunci_Jawaban', '')
        j_siswa = jawaban_siswa.get(idx, None)
        
        bobot_raw = s.get('Bobot_Per_Pernyataan', 1)
        bobot = float(bobot_raw) if pd.notna(bobot_raw) and str(bobot_raw).strip() != '' else 1.0
        
        skor_soal = 0.0
        
        if tipe in ['PG_STANDAR', 'PG_KOMPLEKS_ASOSIASI', 'SEBAB_AKIBAT']:
            if j_siswa and str(j_siswa).strip().upper() == str(kunci).strip().upper():
                skor_soal = bobot
        elif tipe == 'TEPAT_TIDAK_TEPAT':
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
        
    db_target['sesi_siswa'][id_siswa]['nilai_total'] = round(skor_total, 2)
    db_target['sesi_siswa'][id_siswa]['detail_skor'] = detail
    save_current_state()
    return skor_total

def upload_ke_gdrive_guru(file_buffer, file_name, folder_id):
    try:
        creds = service_account.Credentials.from_service_account_info(
            st.secrets["gcp_service_account"],
            scopes=['https://www.googleapis.com/auth/drive']
        )
        service = build('drive', 'v3', credentials=creds)
        file_metadata = {'name': file_name, 'parents': [folder_id]}
        media = MediaIoBaseUpload(
            file_buffer, 
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
            resumable=True
        )
        uploaded_file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        return True, uploaded_file.get('id')
    except Exception as e:
        return False, str(e)

if 'current_siswa' not in st.session_state:
    st.session_state.current_siswa = None

if 'no_soal_aktif' not in st.session_state:
    st.session_state.no_soal_aktif = 0

# ---------------------------------------------------------
# MODE 1: PANEL GURU / ADMIN CBT
# ---------------------------------------------------------
if mode == "Panel Guru (Admin CBT)":
    st.title("⚙️ Panel Kelola & Pemantauan CBT Guru")
    st.info(f"📍 **Kode Unik Ujian Aktif**: `{active_session_code}` | Data tersimpan otomatis & aman selamanya.")
    
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "1. Setting Ujian & Link", 
        "2. Upload Soal (Word/Excel) & Peserta", 
        "3. Pemantauan Real-Time", 
        "4. Reset Ujian", 
        "5. Rekapitulasi Hasil"
    ])
    
    with tab1:
        st.subheader("📋 Identitas Pelaksanaan Ujian")
        col1, col2 = st.columns(2)
        with col1:
            current_db['identitas']['nama_sekolah'] = st.text_input("Nama Sekolah / Instansi:", current_db['identitas']['nama_sekolah'])
            current_db['identitas']['nama_guru'] = st.text_input("Nama Guru / Pengampu:", current_db['identitas']['nama_guru'])
            current_db['identitas']['kelas'] = st.text_input("Kelas / Rombel Target:", current_db['identitas']['kelas'])
        with col2:
            current_db['identitas']['judul_ujian'] = st.text_input("Judul Asesmen / Ujian:", current_db['identitas']['judul_ujian'])
            current_db['identitas']['durasi_menit'] = st.number_input("Durasi Ujian (Menit):", min_value=1, value=int(current_db['identitas']['durasi_menit']))
            current_db['identitas']['base_url_app'] = st.text_input("Domain / URL Aplikasi CBT Anda:", current_db['identitas']['base_url_app'], help="Masukkan URL Streamlit/Custom Domain sekolah Anda agar Link Ujian aktif")

        st.markdown("---")
        st.subheader("🔐 Pengaturan Token 15 Menit & Keamanan TKA")
        col_a, col_b = st.columns(2)
        with col_a:
            current_db['identitas']['token_mode'] = st.radio(
                "Sistem Token Ujian:", 
                ["Dinamis (15 Menit)", "Statis Manual"],
                index=0 if current_db['identitas'].get('token_mode') == 'Dinamis (15 Menit)' else 1
            )
            if current_db['identitas']['token_mode'] == 'Statis Manual':
                current_db['identitas']['token_statis'] = st.text_input("Token Statis:", current_db['identitas']['token_statis'])
            else:
                active_tok, timer_tok = generate_dynamic_token(current_db['identitas'].get('token_secret_salt', f'SALT_{active_session_code}'))
                st.info(f"🔑 **TOKEN AKTIF SAAT INI**: `{active_tok}` (Berganti dalam: **{timer_tok}**)")
                
            current_db['identitas']['max_pelanggaran'] = st.slider("Maksimal Toleransi Kecurangan (Minimize/Tab Switch):", 1, 10, int(current_db['identitas']['max_pelanggaran']))
            
        with col_b:
            current_db['identitas']['wajib_kroscek_peserta'] = st.checkbox(
                "Wajibkan Kroscek Nomor Peserta dengan Daftar Siswa (DPT)", 
                value=current_db['identitas'].get('wajib_kroscek_peserta', False)
            )

        st.subheader("📁 Pengaturan Penyimpanan Google Drive Guru")
        current_db['identitas']['folder_id_gdrive'] = st.text_input(
            "Folder ID Google Drive Guru:",
            value=current_db['identitas'].get('folder_id_gdrive', ''),
            placeholder="Contoh: 1a2B3c4D5e6F7g8H9i0J_kLmnOpQrStUv"
        )

        save_current_state()

        st.markdown("---")
        st.subheader("🚀 Status Kesiapan & Generator Link Ujian Siswa")
        
        has_soal = len(current_db['bank_soal_mapel']) > 0
        has_siswa = len(current_db['daftar_siswa']) > 0 or not current_db['identitas']['wajib_kroscek_peserta']
        has_url = bool(current_db['identitas']['base_url_app'].strip())
        
        col_chk1, col_chk2, col_chk3 = st.columns(3)
        col_chk1.metric("1. Pengaturan Identitas & URL", "READY ✅" if has_url else "BELUM ❌")
        col_chk2.metric("2. Upload Bank Soal", f"READY ({len(current_db['bank_soal_mapel'])} Mapel) ✅" if has_soal else "BELUM ❌")
        col_chk3.metric("3. Kroscek DPT Peserta", f"READY ({len(current_db['daftar_siswa'])} Siswa) ✅" if has_siswa else "OPSIONAL / BELUM ⚠️")
        
        if has_soal and has_url:
            cur_token = get_active_token(current_db['identitas'])
            base_u = current_db['identitas']['base_url_app'].rstrip('/')
            
            link_aktif_siswa = f"{base_u}/?session={active_session_code}&token={cur_token}&mode=siswa"
            
            st.success("🎉 **UJIAN SIAP DILAKSANAKAN!** Bagikan link khusus siswa di bawah ini:")
            st.code(link_aktif_siswa, language="text")
            st.caption("🔒 Link di atas mengunci total menu Administrasi Guru sehingga siswa hanya dapat membuka ruang ujian.")
        else:
            st.warning("⚠️ **Link Ujian Belum Aktif**: Harap unggah minimal 1 Bank Soal pada Tab 2 terlebih dahulu.")

    with tab2:
        st.subheader("📤 Upload Bank Soal Multi-Mapel (Mendukung Word .docx & Excel .xlsx)")
        
        col_u1, col_u2 = st.columns([1, 2])
        with col_u1:
            input_mapel_name = st.text_input("Nama Mata Pelajaran (Contoh: Matematika / Fisika / Biologi):", "Matematika")
        with col_u2:
            file_soal_up = st.file_uploader(f"Unggah File Soal untuk Mapel '{input_mapel_name}' (Format .docx, .xlsx, atau .csv):", type=['docx', 'xlsx', 'xls', 'csv'])
            
        if file_soal_up is not None and input_mapel_name.strip():
            if st.button(f"📥 Impor Bank Soal {input_mapel_name}", type="primary"):
                try:
                    mapel_clean = input_mapel_name.strip()
                    if file_soal_up.name.endswith('.docx'):
                        parsed_soal = parse_word_soal_bytes(file_soal_up.getvalue())
                    elif file_soal_up.name.endswith('.csv'):
                        df_s = pd.read_csv(file_soal_up)
                        parsed_soal = df_s.to_dict('records')
                    else:
                        df_s = pd.read_excel(file_soal_up)
                        parsed_soal = df_s.to_dict('records')
                        
                    if parsed_soal:
                        current_db['bank_soal_mapel'][mapel_clean] = parsed_soal
                        save_current_state()
                        st.success(f"✅ Berhasil mengimpor {len(parsed_soal)} soal untuk Mata Pelajaran **{mapel_clean}**!")
                    else:
                        st.error("Gagal mengekstrak soal dari file. Pastikan tabel format soal sesuai template.")
                except Exception as e:
                    st.error(f"Gagal membaca file: {e}")
                    
        st.markdown("---")
        st.write("### 📚 Daftar Bank Soal Aktif per Mapel:")
        if current_db['bank_soal_mapel']:
            for m_name, q_list in current_db['bank_soal_mapel'].items():
                st.write(f"• **{m_name}**: {len(q_list)} Soal Terunggah")
        else:
            st.info("Belum ada bank soal terunggah.")
            
        st.markdown("---")
        st.subheader("👥 Upload / Isil Daftar Peserta Ujian (DPT Kroscek Nomor Peserta)")
        file_siswa = st.file_uploader("Unggah File Daftar Siswa DPT:", type=['xlsx', 'csv'])
        if file_siswa is not None:
            try:
                if file_siswa.name.endswith('.csv'):
                    df_siswa = pd.read_csv(file_siswa)
                else:
                    df_siswa = pd.read_excel(file_siswa)
                current_db['daftar_siswa'] = df_siswa.to_dict('records')
                save_current_state()
                st.success(f"✅ Berhasil memuat {len(current_db['daftar_siswa'])} data peserta DPT!")
                st.dataframe(df_siswa.head(10))
            except Exception as e:
                st.error(f"Gagal membaca daftar siswa: {e}")

    with tab3:
        st.subheader("📊 Monitoring Real-Time Peserta Ujian")
        total_peserta = len(current_db['daftar_siswa']) if current_db['daftar_siswa'] else len(current_db['sesi_siswa'])
        login_count = len(current_db['sesi_siswa'])
        
        sedang = sum(1 for s in current_db['sesi_siswa'].values() if s.get('status') == 'Sedang Mengerjakan')
        selesai = sum(1 for s in current_db['sesi_siswa'].values() if s.get('status') == 'Selesai')
        belum = total_peserta - login_count if total_peserta >= login_count else 0
        
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Siswa Login", f"{login_count} Siswa")
        m2.metric("Sedang Mengerjakan", f"{sedang} Siswa")
        m3.metric("Selesai Ujian", f"{selesai} Siswa")
        m4.metric("Belum Login/Mengerjakan", f"{belum} Siswa")
        
        st.markdown("---")
        st.write("### 🔍 Tabel Status Detail Per Peserta")
        if current_db['sesi_siswa']:
            monitor_data = []
            for id_s, data in current_db['sesi_siswa'].items():
                m_siswa = data.get('mapel', '')
                total_s_mapel = len(current_db['bank_soal_mapel'].get(m_siswa, []))
                soal_terjawab = len([j for j in data.get('jawaban', {}).values() if is_jawaban_terjawab(j)])
                
                monitor_data.append({
                    'No Peserta': data.get('no_peserta', id_s),
                    'Nama Siswa': data.get('nama', 'Siswa'),
                    'Mata Pelajaran': m_siswa,
                    'Status Ujian': data.get('status', 'Belum'),
                    'Terjawab': f"{soal_terjawab} / {total_s_mapel} Soal",
                    'Indikasi Kecurangan': f"{data.get('pelanggaran', 0)}x Warning",
                    'Waktu Mulai': data.get('waktu_mulai', '-'),
                    'Nilai Sementara': data.get('nilai_total', 0)
                })
            st.dataframe(pd.DataFrame(monitor_data), use_container_width=True)
        else:
            st.info("Belum ada siswa yang login ke ruang ujian.")

    with tab4:
        st.subheader("🔄 Menu Reset Ujian (Per Siswa / Per Kelas)")
        col_r1, col_r2 = st.columns(2)
        with col_r1:
            st.markdown("#### 👤 Reset Ujian Per Siswa")
            if current_db['sesi_siswa']:
                pilih_siswa_reset = st.selectbox("Pilih Siswa yang Ingin Direset:", list(current_db['sesi_siswa'].keys()))
                if st.button("🔴 Reset Ujian Siswa Ini", type="secondary"):
                    del current_db['sesi_siswa'][pilih_siswa_reset]
                    save_current_state()
                    st.success(f"Sesi ujian siswa {pilih_siswa_reset} telah di-reset.")
                    st.rerun()
            else:
                st.write("Tidak ada sesi siswa aktif.")
                
        with col_r2:
            st.markdown("#### 🏫 Reset Seluruh Ujian")
            if st.button("🚨 RESET SEMUA SISWA & SESI", type="primary"):
                current_db['sesi_siswa'] = {}
                save_current_state()
                st.success("Semua data sesi ujian siswa berhasil di-reset bersih!")
                st.rerun()

    with tab5:
        st.subheader("📥 Unduh Rekapitulasi Nilai Ujian (Excel)")
        if current_db['sesi_siswa']:
            rekap_list = []
            for id_s, d in current_db['sesi_siswa'].items():
                row = {
                    'No_Peserta': d.get('no_peserta', id_s),
                    'Nama_Siswa': d.get('nama', ''),
                    'Mata_Pelajaran': d.get('mapel', ''),
                    'Status': d.get('status', ''),
                    'Jumlah_Pelanggaran': d.get('pelanggaran', 0),
                    'Nilai_Total': d.get('nilai_total', 0)
                }
                for q_k, q_v in d.get('detail_skor', {}).items():
                    row[q_k] = q_v
                rekap_list.append(row)
                
            df_rekap = pd.DataFrame(rekap_list)
            st.dataframe(df_rekap, use_container_width=True)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_rekap.to_excel(writer, index=False, sheet_name='Hasil_Ujian')
            excel_data = output.getvalue()
            
            st.download_button(
                label="📥 Download Hasil Ujian (.xlsx)",
                data=excel_data,
                file_name=f"Hasil_Ujian_{current_db['identitas']['judul_ujian']}_{active_session_code}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            if st.button("☁ Simpan Hasil Ujian Langsung ke Google Drive Saya"):
                folder_id = current_db['identitas'].get('folder_id_gdrive', '').strip()
                if not folder_id:
                    st.error("❌ Folder ID Google Drive belum diisi pada Tab 1!")
                else:
                    output.seek(0)
                    nama_file_excel = f"Hasil_Ujian_{current_db['identitas']['judul_ujian']}_{active_session_code}.xlsx"
                    with st.spinner("Sedang mengunggah hasil ujian ke Google Drive..."):
                        berhasil, pesan = upload_ke_gdrive_guru(output, nama_file_excel, folder_id)
                        if berhasil:
                            st.success(f"✅ Berhasil disimpan ke Google Drive! (File ID: {pesan})")
                        else:
                            st.error(f"❌ Gagal mengunggah: {pesan}")
        else:
            st.info("Belum ada data nilai hasil ujian yang dapat diunduh.")

# ---------------------------------------------------------
# MODE 2: HALAMAN UJIAN SISWA (TAMPILAN EXCLUSIF TKA)
# ---------------------------------------------------------
else:
    if not current_db['bank_soal_mapel']:
        st.error(f"⚠️ Ujian belum siap! Bank soal untuk sesi `{active_session_code}` belum diunggah oleh Guru.")
        st.stop()
        
    identitas = current_db['identitas']
    
    # LOGIN SISWA
    if not st.session_state.current_siswa:
        # Header TKA Asli
        st.markdown(f"""
        <div style="background-color: #1E3A8A; padding: 18px 25px; border-radius: 8px; margin-bottom: 20px;">
            <div class="tka-instansi-title">{identitas['nama_sekolah']}</div>
            <div class="tka-exam-sub">{identitas['judul_ujian']} | Durasi: {identitas['durasi_menit']} Menit</div>
        </div>
        """, unsafe_allow_html=True)
        
        with st.form("form_login_siswa"):
            st.subheader("🔑 Konfirmasi Identitas Peserta Ujian")
            
            default_token = url_token if url_token else get_active_token(identitas)
            token_input = st.text_input("Masukkan Token Ujian:", value=default_token)
            
            no_peserta_input = st.text_input("Nomor Peserta Ujian (Contoh: 2026-001):")
            nama_siswa_input = st.text_input("Nama Lengkap Siswa:")
            
            mapel_options = list(current_db['bank_soal_mapel'].keys())
            mapel_pilihan = st.selectbox("Pilih Mata Pelajaran yang Diikuti:", mapel_options)
            
            btn_login = st.form_submit_button("🚀 MASUK RUANG UJIAN")
            
            if btn_login:
                active_token = get_active_token(identitas)
                if token_input.strip().upper() != active_token:
                    st.error(f"❌ Token Ujian Salah atau Telah Kadaluarsa! Token saat ini: `{active_token}`")
                elif not nama_siswa_input or not no_peserta_input:
                    st.error("❌ Harap isi Nomor Peserta dan Nama Lengkap!")
                else:
                    is_valid_dpt = True
                    dpt = current_db['daftar_siswa']
                    if identitas.get('wajib_kroscek_peserta') and dpt:
                        found = [s for s in dpt if str(s.get('No_Peserta', s.get('no_peserta', ''))).strip().lower() == no_peserta_input.strip().lower()]
                        if not found:
                            is_valid_dpt = False
                            st.error("❌ Nomor Peserta tidak terdaftar pada DPT Peserta Ujian!")
                        else:
                            nama_dpt = str(found[0].get('Nama', found[0].get('nama', ''))).strip().lower()
                            if difflib.SequenceMatcher(None, nama_siswa_input.strip().lower(), nama_dpt).ratio() < 0.6:
                                is_valid_dpt = False
                                st.error("❌ Nama Siswa tidak sesuai dengan Nomor Peserta pada DPT!")
                                
                    if is_valid_dpt:
                        id_siswa = str(no_peserta_input).strip()
                        st.session_state.current_siswa = id_siswa
                        
                        if id_siswa not in current_db['sesi_siswa']:
                            current_db['sesi_siswa'][id_siswa] = {
                                'no_peserta': id_siswa,
                                'nama': nama_siswa_input.strip(),
                                'mapel': mapel_pilihan,
                                'status': 'Sedang Mengerjakan',
                                'jawaban': {},
                                'ragu': {},
                                'pelanggaran': 0,
                                'waktu_mulai': datetime.now().strftime("%H:%M:%S"),
                                'nilai_total': 0,
                                'detail_skor': {}
                            }
                            save_current_state()
                        st.success("Login Berhasil! Mengalihkan ke ruang ujian...")
                        st.rerun()

    # AREA KERJA UJIAN SISWA (RUANG UJIAN TKA HASIL REFINEMENT)
    else:
        id_siswa = st.session_state.current_siswa
        sesi = current_db['sesi_siswa'].get(id_siswa)
        
        if not sesi:
            st.error("Sesi ujian tidak ditemukan atau telah di-reset oleh Admin.")
            st.session_state.current_siswa = None
            st.rerun()
            
        if sesi.get('status') == 'Selesai':
            st.balloons()
            st.title("🎉 Ujian Telah Selesai!")
            st.success(f"Terima kasih **{sesi['nama']}**! Jawaban Anda telah tersimpan dengan aman.")
            st.write(f"**Mata Pelajaran**: {sesi.get('mapel', '')} | **Nilai Akhir**: {sesi.get('nilai_total', 0)}")
            if st.button("🚪 Keluar Ruang Ujian"):
                st.session_state.current_siswa = None
                st.rerun()
            st.stop()
            
        # JS ANTI-CHEAT + KEYBOARD LISTENER UNTUK HAK JAWABAN TEKAN A, B, C, D, E
        js_keyboard_and_security = """
        <script>
        // Disable Right Click & Inspect
        document.addEventListener('contextmenu', event => event.preventDefault());
        document.addEventListener('keydown', function(e) {
            if (e.ctrlKey && (e.key === 'c' || e.key === 'v' || e.key === 'u' || e.key === 'a')) {
                e.preventDefault();
            }
        });
        
        // Anti Cheat Tab Switching Warning
        document.addEventListener("visibilitychange", function() {
            if (document.hidden) {
                alert("⚠️ PERINGATAN KECURANGAN: Anda terdeteksi meninggalkan layar ujian! Kejadian ini dicatat oleh pengawas.");
            }
        });

        // KEYBOARD LISTENER: Tekan Huruf A, B, C, D, atau E untuk memilih Opsi
        document.addEventListener('keyup', function(event) {
            const key = event.key.toUpperCase();
            if (['A', 'B', 'C', 'D', 'E'].includes(key)) {
                // Cari radio input yang sesuai dengan opsi A, B, C, D, E
                const radioInputs = document.querySelectorAll('input[type="radio"]');
                if (key === 'A' && radioInputs[0]) { radioInputs[0].click(); }
                else if (key === 'B' && radioInputs[1]) { radioInputs[1].click(); }
                else if (key === 'C' && radioInputs[2]) { radioInputs[2].click(); }
                else if (key === 'D' && radioInputs[3]) { radioInputs[3].click(); }
                else if (key === 'E' && radioInputs[4]) { radioInputs[4].click(); }
            }
        });
        </script>
        """
        st.components.v1.html(js_keyboard_and_security, height=0)
        
        mapel_siswa = sesi.get('mapel', '')
        
        # HEADER BAR RUANG UJIAN ALA TKA (Font Arial 18pt Nama Instansi)
        st.markdown(f"""
        <div style="background-color: #1E3A8A; padding: 14px 22px; border-radius: 6px; color: white; margin-bottom: 15px;">
            <div style="display: flex; justify-content: space-between; align-items: center;">
                <div>
                    <div class="tka-instansi-title">{identitas['nama_sekolah']}</div>
                    <div class="tka-exam-sub">{identitas['judul_ujian']} ({mapel_siswa}) | Peserta: <b>{sesi['nama']} ({sesi['no_peserta']})</b></div>
                </div>
                <div style="text-align: right; background-color: #DC2626; padding: 6px 16px; border-radius: 6px;">
                    <span style="font-size: 10pt; font-family: Arial, sans-serif; display:block; color: #FFFFFF;">SISA WAKTU</span>
                    <strong style="font-size: 16pt; font-family: Arial, sans-serif; color: #FFFFFF;">⏱️ {identitas['durasi_menit']}:00 Min</strong>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        # PROPORSIONAL GRID: Left Side (78% Soal & Opsi) | Right Side (22% Navigation Grid)
        col_left, col_right = st.columns([3.5, 1.1])
        
        soal_list = current_db['bank_soal_mapel'].get(mapel_siswa, [])
        total_soal = len(soal_list)
        
        if total_soal == 0:
            st.error(f"Bank soal untuk mata pelajaran **{mapel_siswa}** belum diunggah oleh guru.")
            st.stop()
            
        curr_idx = st.session_state.no_soal_aktif
        if curr_idx >= total_soal:
            curr_idx = 0
            st.session_state.no_soal_aktif = 0
            
        soal = soal_list[curr_idx]
        tipe = soal.get('Tipe_Soal', 'PG_STANDAR')
        
        with col_left:
            # Container Soal TKA
            st.markdown(f"""
            <div style="font-family: Arial, sans-serif; font-size: 11pt; font-weight: bold; color: #1E3A8A; margin-bottom: 10px;">
                SOAL NOMOR {curr_idx + 1} / {total_soal} &nbsp;<span style="font-size: 9pt; background-color: #E2E8F0; padding: 2px 8px; border-radius: 4px; color: #334155;">[{tipe}]</span>
            </div>
            """, unsafe_allow_html=True)
            
            if pd.notna(soal.get('Media_Gambar')) and str(soal.get('Media_Gambar')).strip() != '':
                st.info(f"🖼️ Media Gambar / Grafik Attached: `{soal.get('Media_Gambar')}`")
                
            # Stimulus Teks (Arial 12pt)
            if pd.notna(soal.get('Stimulus_Teks')) and str(soal.get('Stimulus_Teks')).strip() != '':
                st.markdown(f"""
                <div class="tka-stimulus-box">
                    {soal.get('Stimulus_Teks')}
                </div>
                """, unsafe_allow_html=True)
                
            # Soal Utama (Arial 12pt Bold)
            st.markdown(f"""
            <div class="tka-soal-text">
                {soal.get('Soal_Utama', '')}
            </div>
            """, unsafe_allow_html=True)
            
            curr_ans = sesi['jawaban'].get(curr_idx, None)
            
            if tipe in ['PG_STANDAR', 'PG_KOMPLEKS_ASOSIASI', 'SEBAB_AKIBAT']:
                if tipe == 'PG_KOMPLEKS_ASOSIASI':
                    st.write("Daftar Pernyataan:")
                    for i in range(1, 5):
                        p_val = soal.get(f'Pernyataan_{i}') or soal.get(f'Pernyataan_{i}_atau_Sebab') or soal.get(f'Pernyataan_{i}_atau_Akibat')
                        if pd.notna(p_val) and str(p_val).strip() != '':
                            st.write(f"• {p_val}")
                elif tipe == 'SEBAB_AKIBAT':
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
                
                st.caption("💡 *Petunjuk: Anda bisa mengklik langsung opsi di bawah atau menekan tombol A, B, C, D, E pada keyboard.*")
                
                jawaban_user = st.radio(
                    "Pilih Jawaban Anda:",
                    opsi_choices,
                    format_func=lambda x: opsi_labels[x],
                    index=def_index,
                    key=f"radio_soal_{curr_idx}",
                    label_visibility="collapsed"
                )
                if jawaban_user:
                    sesi['jawaban'][curr_idx] = jawaban_user
                    save_current_state()

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
                save_current_state()

            elif tipe == 'URAIAN_SINGKAT':
                val_uraian = str(curr_ans) if curr_ans is not None else ""
                ans_uraian = st.text_area("Ketik Jawaban Singkat Anda di Sini:", value=val_uraian, key=f"uraian_{curr_idx}")
                sesi['jawaban'][curr_idx] = ans_uraian
                save_current_state()

            st.markdown("---")
            # FOOTER BAR TOMBOL KETERANGAN SOAL (Font Arial 10pt)
            nav_col1, nav_col2, nav_col3 = st.columns([1.2, 1, 1.2])
            with nav_col1:
                if st.button("⬅️ SOAL SEBELUMNYA", disabled=(curr_idx == 0), use_container_width=True):
                    st.session_state.no_soal_aktif -= 1
                    st.rerun()
            with nav_col2:
                is_ragu = sesi['ragu'].get(curr_idx, False)
                if st.checkbox("🟨 RAGU-RAGU", value=is_ragu, key=f"ragu_{curr_idx}"):
                    sesi['ragu'][curr_idx] = True
                else:
                    sesi['ragu'][curr_idx] = False
                save_current_state()
            with nav_col3:
                if curr_idx < total_soal - 1:
                    if st.button("SOAL BERIKUTNYA ➡️", type="primary", use_container_width=True):
                        st.session_state.no_soal_aktif += 1
                        st.rerun()
                else:
                    if st.button("🏁 SELESAI & SIMPAN UJIAN", type="primary", use_container_width=True):
                        sesi['status'] = 'Selesai'
                        hitung_skor_siswa(id_siswa, current_db)
                        save_current_state()
                        st.rerun()

        # KOTAK PENANDA SOAL SEBELAH KANAN (KISI-KISI NOMOR SOAL FONT ARIAL 10pt)
        with col_right:
            st.markdown("""
            <div style="background-color: #FFFFFF; padding: 12px; border: 1px solid #CBD5E1; border-radius: 6px; margin-bottom: 15px;">
                <div style="font-family: Arial, sans-serif; font-size: 10pt; font-weight: bold; color: #1E3A8A; margin-bottom: 10px; text-align: center;">
                    🔲 NAVIGASI NOMOR SOAL
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            cols_grid = st.columns(4)
            for i in range(total_soal):
                is_ans = is_jawaban_terjawab(sesi['jawaban'].get(i))
                is_ragu = sesi['ragu'].get(i, False)
                
                label_num = f"{i+1}"
                if is_ragu:
                    label_num += " 🟨"
                elif is_ans:
                    label_num += " 🟢"
                    
                with cols_grid[i % 4]:
                    if st.button(label_num, key=f"grid_btn_{i}", use_container_width=True):
                        st.session_state.no_soal_aktif = i
                        st.rerun()
                        
            st.markdown("""
            <div style="font-family: Arial, sans-serif; font-size: 10pt; color: #475569; margin-top: 15px; background-color: #FFFFFF; padding: 10px; border-radius: 6px; border: 1px solid #E2E8F0;">
                <b>Keterangan:</b><br>
                🟢 : Sudah Dijawab<br>
                🟨 : Ragu-Ragu<br>
                ⚪ : Belum Dijawab
            </div>
            """, unsafe_allow_html=True)
