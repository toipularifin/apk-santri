"""
================================================================================
SISTEM MANAJEMEN PONDOK PESANTREN (e-PESANTREN) - FULL-STACK PYTHON APPLICATION
================================================================================
Fitur Utama:
1. Multi-Role Terhubung: Admin, Guru, Pengurus, dan Orang Tua (Wali Santri).
2. Manajemen Santri & Pengurus/Guru: CRUD lengkap, upload foto, impor massal
   via Excel (.xlsx), dan unduh template Excel otomatis.
3. Kustomisasi Identitas Ponpes: Ubah nama pesantren, alamat, pimpinan, & logo.
4. Absensi Cerdas (QR Code, RFID Reader, & Manual): Terintegrasi langsung
   dengan jadwal pengajian yang diatur oleh Admin (otomatis hitung Hadir/Terlambat).
5. Cetak Kartu Tanda Santri Bolak-Balik: Desain kartu ID standar CR80 lengkap
   dengan QR Code, kode RFID, foto santri, logo ponpes, dan tata tertib.
6. Keuangan & Pembayaran Real-Time: Pembayaran yang diinput Pengurus langsung
   memperbarui tagihan di dashboard Orang Tua secara seketika.
7. Akademik, Setoran Hafalan, Pelanggaran, & Ujian (Soal PG/Esai & Nilai).
8. Siap Publish GitHub & Hosting: Mode konkurensi SQLite WAL + Gunicorn.
================================================================================
"""

import os
import io
import sys
import json
import uuid
import base64
import sqlite3
import zipfile
import secrets
import subprocess
import threading
import webbrowser
from datetime import datetime, date
from functools import wraps

from flask import (
    Flask, request, session, redirect, url_for,
    render_template_string, jsonify, send_file
)
from werkzeug.security import generate_password_hash, check_password_hash

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    import qrcode
    HAS_QRCODE = True
except ImportError:
    HAS_QRCODE = False

try:
    from PIL import Image
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "pesantren-secure-key-2026-change-in-prod")
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
DB_PATH = os.environ.get("SQLITE_DB_PATH", os.path.join(BASE_DIR, "pesantren.db"))


def get_db():
    conn = sqlite3.connect(DB_PATH, timeout=15.0, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute("PRAGMA synchronous=NORMAL;")
    conn.execute("PRAGMA foreign_keys=ON;")
    conn.execute("PRAGMA busy_timeout=15000;")
    return conn


def init_db():
    conn = get_db()
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS settings (
            id INTEGER PRIMARY KEY CHECK (id = 1),
            ponpes_name TEXT NOT NULL,
            ponpes_tagline TEXT,
            ponpes_address TEXT,
            ponpes_phone TEXT,
            pimpinan_name TEXT,
            logo_data TEXT,
            updated_at TEXT
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            full_name TEXT NOT NULL,
            role TEXT NOT NULL CHECK (role IN ('admin', 'guru', 'pengurus', 'ortu')),
            jabatan TEXT,
            phone TEXT,
            photo_data TEXT,
            created_at TEXT
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS santri (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nis TEXT UNIQUE NOT NULL,
            full_name TEXT NOT NULL,
            gender TEXT DEFAULT 'Laki-laki',
            asrama TEXT,
            kelas TEXT,
            address TEXT,
            wali_name TEXT,
            wali_phone TEXT,
            rfid_uid TEXT UNIQUE,
            qr_code_val TEXT UNIQUE,
            photo_data TEXT,
            ortu_user_id INTEGER,
            status TEXT DEFAULT 'Aktif',
            created_at TEXT,
            FOREIGN KEY (ortu_user_id) REFERENCES users(id) ON DELETE SET NULL
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS pengajian (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            mapel TEXT NOT NULL,
            ustadz_name TEXT,
            start_time TEXT NOT NULL,
            end_time TEXT NOT NULL,
            late_after_minutes INTEGER DEFAULT 15,
            days TEXT DEFAULT 'Setiap Hari',
            created_at TEXT
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS absensi (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            santri_id INTEGER NOT NULL,
            pengajian_id INTEGER,
            tanggal TEXT NOT NULL,
            waktu_scan TEXT NOT NULL,
            status TEXT NOT NULL CHECK (status IN ('Hadir', 'Terlambat', 'Izin', 'Sakit', 'Alpha')),
            metode TEXT DEFAULT 'Manual',
            recorded_by TEXT,
            notes TEXT,
            created_at TEXT,
            FOREIGN KEY (santri_id) REFERENCES santri(id) ON DELETE CASCADE,
            FOREIGN KEY (pengajian_id) REFERENCES pengajian(id) ON DELETE SET NULL
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS progres_pengajian (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            santri_id INTEGER NOT NULL,
            pengajian_id INTEGER,
            tanggal TEXT NOT NULL,
            jenis TEXT NOT NULL,
            materi TEXT NOT NULL,
            nilai TEXT,
            catatan TEXT,
            recorded_by TEXT,
            created_at TEXT,
            FOREIGN KEY (santri_id) REFERENCES santri(id) ON DELETE CASCADE,
            FOREIGN KEY (pengajian_id) REFERENCES pengajian(id) ON DELETE SET NULL
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS pelanggaran (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            santri_id INTEGER NOT NULL,
            nama_pelanggaran TEXT NOT NULL,
            kategori TEXT DEFAULT 'Ringan',
            poin INTEGER DEFAULT 5,
            tanggal TEXT NOT NULL,
            tindakan TEXT,
            recorded_by TEXT,
            created_at TEXT,
            FOREIGN KEY (santri_id) REFERENCES santri(id) ON DELETE CASCADE
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS jenis_pembayaran (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nama_tagihan TEXT NOT NULL,
            nominal INTEGER NOT NULL,
            periode TEXT DEFAULT 'Bulanan',
            keterangan TEXT,
            created_at TEXT
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS pembayaran (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            kode_transaksi TEXT UNIQUE NOT NULL,
            santri_id INTEGER NOT NULL,
            jenis_id INTEGER,
            nama_tagihan_snapshot TEXT NOT NULL,
            periode_tagihan TEXT,
            tagihan_amount INTEGER NOT NULL,
            dibayar_amount INTEGER NOT NULL DEFAULT 0,
            status TEXT NOT NULL CHECK (status IN ('Lunas', 'Cicilan', 'Belum Lunas')),
            metode TEXT DEFAULT 'Tunai',
            tanggal_bayar TEXT,
            recorded_by TEXT,
            notes TEXT,
            created_at TEXT,
            FOREIGN KEY (santri_id) REFERENCES santri(id) ON DELETE CASCADE,
            FOREIGN KEY (jenis_id) REFERENCES jenis_pembayaran(id) ON DELETE SET NULL
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS ujian (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            judul_ujian TEXT NOT NULL,
            mapel TEXT NOT NULL,
            kelas TEXT NOT NULL,
            tanggal TEXT NOT NULL,
            durasi_menit INTEGER DEFAULT 60,
            deskripsi TEXT,
            created_by TEXT,
            created_at TEXT
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS soal_ujian (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ujian_id INTEGER NOT NULL,
            nomor_soal INTEGER DEFAULT 1,
            pertanyaan TEXT NOT NULL,
            opsi_a TEXT,
            opsi_b TEXT,
            opsi_c TEXT,
            opsi_d TEXT,
            kunci_jawaban TEXT,
            bobot INTEGER DEFAULT 10,
            created_by TEXT,
            FOREIGN KEY (ujian_id) REFERENCES ujian(id) ON DELETE CASCADE
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS nilai_ujian (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ujian_id INTEGER NOT NULL,
            santri_id INTEGER NOT NULL,
            skor REAL NOT NULL,
            catatan TEXT,
            graded_by TEXT,
            created_at TEXT,
            FOREIGN KEY (ujian_id) REFERENCES ujian(id) ON DELETE CASCADE,
            FOREIGN KEY (santri_id) REFERENCES santri(id) ON DELETE CASCADE
        )
    """)

    cur.execute("SELECT COUNT(*) FROM settings")
    if cur.fetchone()[0] == 0:
        cur.execute("""
            INSERT INTO settings (id, ponpes_name, ponpes_tagline, ponpes_address, ponpes_phone, pimpinan_name, logo_data, updated_at)
            VALUES (1, ?, ?, ?, ?, ?, ?, ?)
        """, (
            "Pondok Pesantren Modern Al-Hikmah",
            "Mencetak Generasi Qur'ani, Berakhlak Mulia & Berwawasan Luas",
            "Jl. Pesantren Raya No. 99, Kab. Jombang, Jawa Timur",
            "(0321) 889900 / 0812-3456-7890",
            "KH. Ahmad Mansyur, Lc., M.Ag.",
            "",
            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ))

    cur.execute("SELECT COUNT(*) FROM users")
    if cur.fetchone()[0] == 0:
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cur.execute("""
            INSERT INTO users (username, password_hash, full_name, role, jabatan, phone, created_at)
            VALUES (?, ?, ?, 'admin', 'Administrator Utama', '08111111111', ?)
        """, ("admin", generate_password_hash("admin123"), "Admin Utama Ponpes", now_str))

        cur.execute("""
            INSERT INTO users (username, password_hash, full_name, role, jabatan, phone, created_at)
            VALUES (?, ?, ?, 'guru', 'Ustadz Pengampu Tahfidz & Kitab', '08222222222', ?)
        """, ("guru1", generate_password_hash("guru123"), "Ust. Muhammad Ridwan, S.Pd.I", now_str))

        cur.execute("""
            INSERT INTO users (username, password_hash, full_name, role, jabatan, phone, created_at)
            VALUES (?, ?, ?, 'pengurus', 'Ketua Bagian Keamanan & Kesantrian', '08333333333', ?)
        """, ("pengurus1", generate_password_hash("pengurus123"), "Ust. Hasan Basri", now_str))

        cur.execute("""
            INSERT INTO users (username, password_hash, full_name, role, jabatan, phone, created_at)
            VALUES (?, ?, ?, 'ortu', 'Wali Santri (Bpk. H. Sulaiman)', '08444444444', ?)
        """, ("wali1", generate_password_hash("wali123"), "Bpk. H. Sulaiman", now_str))
        ortu_id = cur.lastrowid

        cur.execute("""
            INSERT INTO santri (nis, full_name, gender, asrama, kelas, address, wali_name, wali_phone, rfid_uid, qr_code_val, ortu_user_id, status, created_at)
            VALUES (?, ?, 'Laki-laki', ?, ?, ?, ?, ?, ?, ?, ?, 'Aktif', ?)
        """, (
            "2026001", "Ahmad Fauzi Al-Fatih", "Asrama Al-Ghazali 01", "1 Ulya",
            "Jl. Merdeka No. 12, Surabaya", "Bpk. H. Sulaiman", "08444444444",
            "0001234567", "SANTRI-2026001-ALHIKMAH", ortu_id, now_str
        ))
        santri_id = cur.lastrowid

        cur.execute("""
            INSERT INTO santri (nis, full_name, gender, asrama, kelas, address, wali_name, wali_phone, rfid_uid, qr_code_val, ortu_user_id, status, created_at)
            VALUES (?, ?, 'Perempuan', ?, ?, ?, ?, ?, ?, ?, NULL, 'Aktif', ?)
        """, (
            "2026002", "Siti Aisyah Humaira", "Asrama Khadijah 03", "2 Wustha",
            "Jl. Melati No. 45, Malang", "Bpk. Abdul Karim", "08555555555",
            "0007654321", "SANTRI-2026002-ALHIKMAH", now_str
        ))

        cur.execute("""
            INSERT INTO pengajian (title, mapel, ustadz_name, start_time, end_time, late_after_minutes, days, created_at)
            VALUES (?, ?, ?, '04:45', '06:00', 15, 'Setiap Hari', ?)
        """, ("Pengajian Tahfidz Ba'da Subuh", "Tahfidzul Qur'an", "Ust. Muhammad Ridwan, S.Pd.I", now_str))
        pengajian_id = cur.lastrowid

        cur.execute("""
            INSERT INTO pengajian (title, mapel, ustadz_name, start_time, end_time, late_after_minutes, days, created_at)
            VALUES (?, ?, ?, '18:30', '20:00', 15, 'Senin - Sabtu', ?)
        """, ("Kajian Kitab Fathul Qorib", "Fiqih Syafi'i", "KH. Ahmad Mansyur, Lc.", now_str))

        cur.execute("""
            INSERT INTO jenis_pembayaran (nama_tagihan, nominal, periode, keterangan, created_at)
            VALUES (?, 650000, 'Bulanan', 'Syahriah SPP & Biaya Makan Santri', ?)
        """, ("SPP & Makan Bulanan", now_str))
        jenis_id = cur.lastrowid

        cur.execute("""
            INSERT INTO jenis_pembayaran (nama_tagihan, nominal, periode, keterangan, created_at)
            VALUES (?, 1200000, 'Tahunan', 'Seragam, Kitab, & Pemeliharaan Asrama', ?)
        """, ("Daftar Ulang & Kitab Tahunan", now_str))

        today_str = date.today().strftime("%Y-%m-%d")
        cur.execute("""
            INSERT INTO pembayaran (kode_transaksi, santri_id, jenis_id, nama_tagihan_snapshot, periode_tagihan, tagihan_amount, dibayar_amount, status, metode, tanggal_bayar, recorded_by, notes, created_at)
            VALUES (?, ?, ?, 'SPP & Makan Bulanan', 'November 2026', 650000, 650000, 'Lunas', 'Transfer', ?, 'Admin Utama Ponpes', 'Lunas tepat waktu', ?)
        """, ("TRX-202611-001", santri_id, jenis_id, today_str, now_str))

        cur.execute("""
            INSERT INTO progres_pengajian (santri_id, pengajian_id, tanggal, jenis, materi, nilai, catatan, recorded_by, created_at)
            VALUES (?, ?, ?, 'Setoran Al-Qur''an', 'Juz 30 - Surah An-Naba s/d An-Nazi''at', 'Mumtaz (96)', 'Tajwid dan makhraj sangat lancar', 'Ust. Muhammad Ridwan, S.Pd.I', ?)
        """, (santri_id, pengajian_id, today_str, now_str))

        cur.execute("""
            INSERT INTO ujian (judul_ujian, mapel, kelas, tanggal, durasi_menit, deskripsi, created_by, created_at)
            VALUES ('Ujian Tengah Semester Ganjil', 'Fiqih & Tahfidz', '1 Ulya', ?, 90, 'Evaluasi hafalan Juz 30 dan Bab Thaharah', 'Admin Utama Ponpes', ?)
        """, (today_str, now_str))
        ujian_id = cur.lastrowid

        cur.execute("""
            INSERT INTO soal_ujian (ujian_id, nomor_soal, pertanyaan, opsi_a, opsi_b, opsi_c, opsi_d, kunci_jawaban, bobot, created_by)
            VALUES (?, 1, 'Berapakah jumlah rukun wudhu menurut Mazhab Syafi''i?', '4 Rukun', '5 Rukun', '6 Rukun', '7 Rukun', 'C', 20, 'Ust. Muhammad Ridwan, S.Pd.I')
        """, (ujian_id,))

    conn.commit()
    conn.close()


def process_uploaded_image(file_storage, max_size=(400, 400)):
    if not file_storage or not file_storage.filename:
        return None
    raw = file_storage.read()
    if not raw:
        return None
    if HAS_PIL:
        try:
            img = Image.open(io.BytesIO(raw))
            if img.mode in ("RGBA", "P"):
                img = img.convert("RGB")
            img.thumbnail(max_size)
            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=82, optimize=True)
            encoded = base64.b64encode(buf.getvalue()).decode("utf-8")
            return f"data:image/jpeg;base64,{encoded}"
        except Exception:
            pass
    ext = file_storage.filename.rsplit(".", 1)[-1].lower()
    mime = "image/png" if ext == "png" else "image/jpeg"
    encoded = base64.b64encode(raw).decode("utf-8")
    return f"data:{mime};base64,{encoded}"


def generate_qr_base64(text_data):
    if not HAS_QRCODE:
        return ""
    try:
        qr = qrcode.QRCode(version=1, box_size=6, border=2)
        qr.add_data(text_data)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white")
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode("utf-8")
    except Exception:
        return ""


def get_ponpes_settings():
    conn = get_db()
    row = conn.execute("SELECT * FROM settings WHERE id = 1").fetchone()
    conn.close()
    return dict(row) if row else {
        "ponpes_name": "Pondok Pesantren",
        "ponpes_tagline": "",
        "ponpes_address": "",
        "ponpes_phone": "",
        "pimpinan_name": "",
        "logo_data": ""
    }


def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if "user_id" not in session:
            return redirect(url_for("login", msg="Silakan login terlebih dahulu.", type="warning"))
        return f(*args, **kwargs)
    return decorated


def role_required(*roles):
    def decorator(f):
        @wraps(f)
        def decorated(*args, **kwargs):
            if "user_id" not in session:
                return redirect(url_for("login"))
            if session.get("role") not in roles:
                return redirect(url_for("dashboard", msg="Anda tidak memiliki izin akses.", type="error"))
            return f(*args, **kwargs)
        return decorated
    return decorator


BASE_HTML = """
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{{ title }} - {{ settings.ponpes_name }}</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script defer src="https://cdn.jsdelivr.net/npm/alpinejs@3.x.x/dist/cdn.min.js"></script>
    <style>
        body { font-family: system-ui, -apple-system, sans-serif; background-color: #f8fafc; }
        [x-cloak] { display: none !important; }
        @media print { .no-print { display: none !important; } body { background: white !important; } }
    </style>
</head>
<body class="text-slate-800 antialiased" x-data="appShell()">
    <div x-show="toast.show" x-cloak class="fixed top-5 right-5 z-50 max-w-md w-full shadow-xl rounded-xl border p-4 flex items-start gap-3 no-print"
         :class="{'bg-emerald-50 border-emerald-200 text-emerald-900': toast.type === 'success', 'bg-rose-50 border-rose-200 text-rose-900': toast.type === 'error', 'bg-amber-50 border-amber-200 text-amber-900': toast.type === 'warning'}">
        <div class="flex-1 text-sm font-medium" x-text="toast.message"></div>
        <button @click="toast.show = false" class="text-slate-400 hover:text-slate-700 font-bold">&times;</button>
    </div>

    <div x-show="confirmModal.show" x-cloak class="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/60 p-4 no-print">
        <div class="bg-white rounded-2xl max-w-md w-full p-6 shadow-2xl" @click.away="confirmModal.show = false">
            <h3 class="text-lg font-bold text-slate-900 mb-2" x-text="confirmModal.title"></h3>
            <p class="text-sm text-slate-600 mb-6" x-text="confirmModal.message"></p>
            <div class="flex justify-end gap-3">
                <button type="button" @click="confirmModal.show = false" class="px-4 py-2 rounded-xl border text-sm font-semibold">Batal</button>
                <form :action="confirmModal.action" method="POST">
                    <button type="submit" class="px-4 py-2 rounded-xl bg-rose-600 text-white text-sm font-semibold">Ya, Hapus</button>
                </form>
            </div>
        </div>
    </div>

    {% if session.get('user_id') %}
    <div class="min-h-screen flex flex-col lg:flex-row">
        <aside class="w-full lg:w-72 bg-emerald-950 text-white flex flex-col shadow-xl no-print">
            <div class="p-5 border-b border-emerald-900 flex items-center gap-3">
                {% if settings.logo_data %}
                    <img src="{{ settings.logo_data }}" class="w-11 h-11 rounded-xl object-cover bg-white p-0.5">
                {% else %}
                    <div class="w-11 h-11 rounded-xl bg-amber-400 text-emerald-950 flex items-center justify-center font-extrabold text-xl">🕌</div>
                {% endif %}
                <div>
                    <h1 class="font-extrabold text-sm leading-tight">{{ settings.ponpes_name }}</h1>
                    <span class="inline-block mt-1 px-2 py-0.5 text-[10px] font-bold uppercase rounded-full bg-emerald-800 text-amber-300">Panel {{ session.get('role') }}</span>
                </div>
            </div>
            <nav class="flex-1 p-4 space-y-1 text-sm">
                <a href="{{ url_for('dashboard') }}" class="block px-3.5 py-2.5 rounded-xl font-medium {{ 'bg-emerald-800 text-amber-300 font-bold' if active_page == 'dashboard' else 'text-emerald-100 hover:bg-emerald-900' }}">📊 Dashboard Utama</a>
                {% if session.get('role') == 'admin' %}
                    <a href="{{ url_for('admin_santri') }}" class="block px-3.5 py-2.5 rounded-xl font-medium {{ 'bg-emerald-800 text-amber-300 font-bold' if active_page == 'santri' else 'text-emerald-100 hover:bg-emerald-900' }}">🎓 Manajemen Santri</a>
                    <a href="{{ url_for('admin_users') }}" class="block px-3.5 py-2.5 rounded-xl font-medium {{ 'bg-emerald-800 text-amber-300 font-bold' if active_page == 'users' else 'text-emerald-100 hover:bg-emerald-900' }}">👥 Pengurus, Guru & Ortu</a>
                    <a href="{{ url_for('kartu_santri') }}" class="block px-3.5 py-2.5 rounded-xl font-medium {{ 'bg-emerald-800 text-amber-300 font-bold' if active_page == 'kartu' else 'text-emerald-100 hover:bg-emerald-900' }}">🪪 Cetak Kartu Santri</a>
                    <a href="{{ url_for('admin_pengajian') }}" class="block px-3.5 py-2.5 rounded-xl font-medium {{ 'bg-emerald-800 text-amber-300 font-bold' if active_page == 'pengajian' else 'text-emerald-100 hover:bg-emerald-900' }}">📖 Pengajian & Waktu Absen</a>
                    <a href="{{ url_for('absensi_page') }}" class="block px-3.5 py-2.5 rounded-xl font-medium {{ 'bg-emerald-800 text-amber-300 font-bold' if active_page == 'absensi' else 'text-emerald-100 hover:bg-emerald-900' }}">📷 Scan Absensi QR & RFID</a>
                    <a href="{{ url_for('ujian_page') }}" class="block px-3.5 py-2.5 rounded-xl font-medium {{ 'bg-emerald-800 text-amber-300 font-bold' if active_page == 'ujian' else 'text-emerald-100 hover:bg-emerald-900' }}">📝 Manajemen Ujian & Soal</a>
                    <a href="{{ url_for('pembayaran_page') }}" class="block px-3.5 py-2.5 rounded-xl font-medium {{ 'bg-emerald-800 text-amber-300 font-bold' if active_page == 'pembayaran' else 'text-emerald-100 hover:bg-emerald-900' }}">💳 Pembayaran Pondok</a>
                    <a href="{{ url_for('admin_settings') }}" class="block px-3.5 py-2.5 rounded-xl font-medium {{ 'bg-emerald-800 text-amber-300 font-bold' if active_page == 'settings' else 'text-emerald-100 hover:bg-emerald-900' }}">⚙️ Identitas Ponpes & Akun</a>
                {% endif %}
                {% if session.get('role') in ['pengurus', 'guru'] %}
                    <a href="{{ url_for('absensi_page') }}" class="block px-3.5 py-2.5 rounded-xl font-medium {{ 'bg-emerald-800 text-amber-300 font-bold' if active_page == 'absensi' else 'text-emerald-100 hover:bg-emerald-900' }}">📷 Input & Scan Absensi</a>
                    <a href="{{ url_for('progres_page') }}" class="block px-3.5 py-2.5 rounded-xl font-medium {{ 'bg-emerald-800 text-amber-300 font-bold' if active_page == 'progres' else 'text-emerald-100 hover:bg-emerald-900' }}">🕌 Progres Pengajian & Setoran</a>
                    <a href="{{ url_for('pelanggaran_page') }}" class="block px-3.5 py-2.5 rounded-xl font-medium {{ 'bg-emerald-800 text-amber-300 font-bold' if active_page == 'pelanggaran' else 'text-emerald-100 hover:bg-emerald-900' }}">⚠️ Input Pelanggaran</a>
                    <a href="{{ url_for('pembayaran_page') }}" class="block px-3.5 py-2.5 rounded-xl font-medium {{ 'bg-emerald-800 text-amber-300 font-bold' if active_page == 'pembayaran' else 'text-emerald-100 hover:bg-emerald-900' }}">💳 Input Pembayaran</a>
                    <a href="{{ url_for('ujian_page') }}" class="block px-3.5 py-2.5 rounded-xl font-medium {{ 'bg-emerald-800 text-amber-300 font-bold' if active_page == 'ujian' else 'text-emerald-100 hover:bg-emerald-900' }}">📝 Input Soal & Nilai Ujian</a>
                {% endif %}
                {% if session.get('role') == 'ortu' %}
                    <a href="{{ url_for('ortu_monitoring') }}" class="block px-3.5 py-2.5 rounded-xl font-medium {{ 'bg-emerald-800 text-amber-300 font-bold' if active_page == 'ortu' else 'text-emerald-100 hover:bg-emerald-900' }}">👨‍👩‍👦 Laporan Lengkap Anak</a>
                    <a href="{{ url_for('kartu_santri') }}" class="block px-3.5 py-2.5 rounded-xl font-medium {{ 'bg-emerald-800 text-amber-300 font-bold' if active_page == 'kartu' else 'text-emerald-100 hover:bg-emerald-900' }}">🪪 Kartu Digital Santri</a>
                {% endif %}
            </nav>
            <div class="p-4 border-t border-emerald-900">
                <div class="text-xs font-bold mb-2">{{ session.get('full_name') }} (@{{ session.get('username') }})</div>
                <a href="{{ url_for('logout') }}" class="block w-full text-center py-2 rounded-xl bg-rose-600 text-white text-xs font-bold">Keluar / Logout</a>
            </div>
        </aside>
        <main class="flex-1 p-5 sm:p-8">{{ content | safe }}</main>
    </div>
    {% else %}
        {{ content | safe }}
    {% endif %}
    <script>
        function appShell() {
            return {
                toast: { show: {{ 'true' if request.args.get('msg') else 'false' }}, message: {{ request.args.get('msg', '') | tojson }}, type: {{ request.args.get('type', 'success') | tojson }} },
                confirmModal: { show: false, title: 'Konfirmasi', message: '', action: '' },
                init() { if (this.toast.show) setTimeout(() => { this.toast.show = false; }, 5000); },
                askDelete(title, message, actionUrl) { this.confirmModal = { show: true, title, message, action: actionUrl }; }
            }
        }
    </script>
</body>
</html>
"""


def render_page(title, active_page, inner_html, **context):
    settings = get_ponpes_settings()
    rendered_inner = render_template_string(inner_html, settings=settings, **context)
    return render_template_string(BASE_HTML, title=title, active_page=active_page, settings=settings, content=rendered_inner)


@app.route("/")
def index():
    return redirect(url_for("dashboard" if "user_id" in session else "login"))


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        conn = get_db()
        user = conn.execute("SELECT * FROM users WHERE username = ?", (username,)).fetchone()
        conn.close()
        if user and check_password_hash(user["password_hash"], password):
            session["user_id"] = user["id"]
            session["username"] = user["username"]
            session["full_name"] = user["full_name"]
            session["role"] = user["role"]
            return redirect(url_for("dashboard", msg=f"Selamat datang, {user['full_name']}!", type="success"))
        return redirect(url_for("login", msg="Username atau password salah!", type="error"))

    login_html = """
    <div class="min-h-screen bg-gradient-to-br from-emerald-950 via-emerald-900 to-teal-900 flex items-center justify-center p-4">
        <div class="max-w-3xl w-full bg-white rounded-3xl shadow-2xl overflow-hidden grid grid-cols-1 md:grid-cols-12">
            <div class="md:col-span-5 bg-emerald-950 text-white p-8 flex flex-col justify-between">
                <div>
                    <div class="w-12 h-12 rounded-2xl bg-amber-400 text-emerald-950 flex items-center justify-center text-2xl font-extrabold mb-4">🕌</div>
                    <h1 class="text-lg font-extrabold">{{ settings.ponpes_name }}</h1>
                    <p class="text-xs text-emerald-200 mt-2">{{ settings.ponpes_tagline }}</p>
                </div>
                <div class="text-[11px] text-emerald-300 mt-6">Sistem Terpadu Admin, Guru, Pengurus & Orang Tua</div>
            </div>
            <div class="md:col-span-7 p-8">
                <h2 class="text-xl font-extrabold text-slate-900 mb-4">Login Sistem Pesantren</h2>
                <form method="POST" class="space-y-4">
                    <div>
                        <label class="block text-xs font-bold text-slate-600 mb-1">Username</label>
                        <input type="text" name="username" id="lu" required class="w-full px-4 py-2.5 rounded-xl border text-sm">
                    </div>
                    <div>
                        <label class="block text-xs font-bold text-slate-600 mb-1">Password</label>
                        <input type="password" name="password" id="lp" required class="w-full px-4 py-2.5 rounded-xl border text-sm">
                    </div>
                    <button type="submit" class="w-full py-3 rounded-xl bg-emerald-700 text-white font-bold text-sm">Masuk</button>
                </form>
                <div class="mt-6 pt-4 border-t grid grid-cols-2 gap-2 text-xs">
                    <button type="button" onclick="lu.value='admin';lp.value='admin123'" class="p-2 rounded-lg bg-slate-100 font-semibold">👑 Admin</button>
                    <button type="button" onclick="lu.value='guru1';lp.value='guru123'" class="p-2 rounded-lg bg-slate-100 font-semibold">👳 Guru</button>
                    <button type="button" onclick="lu.value='pengurus1';lp.value='pengurus123'" class="p-2 rounded-lg bg-slate-100 font-semibold">📋 Pengurus</button>
                    <button type="button" onclick="lu.value='wali1';lp.value='wali123'" class="p-2 rounded-lg bg-slate-100 font-semibold">👨‍👩‍👦 Orang Tua</button>
                </div>
            </div>
        </div>
    </div>
    """
    return render_page("Login", "login", login_html)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login", msg="Berhasil keluar.", type="success"))


@app.route("/dashboard")
@login_required
def dashboard():
    if session.get("role") == "ortu":
        return redirect(url_for("ortu_monitoring"))
    conn = get_db()
    today_str = date.today().strftime("%Y-%m-%d")
    total_santri = conn.execute("SELECT COUNT(*) FROM santri WHERE status='Aktif'").fetchone()[0]
    total_pengurus = conn.execute("SELECT COUNT(*) FROM users WHERE role IN ('pengurus','guru')").fetchone()[0]
    hadir_hari_ini = conn.execute("SELECT COUNT(*) FROM absensi WHERE tanggal=? AND status IN ('Hadir','Terlambat')", (today_str,)).fetchone()[0]
    total_pemasukan = conn.execute("SELECT COALESCE(SUM(dibayar_amount), 0) FROM pembayaran").fetchone()[0]
    jadwal_list = conn.execute("SELECT * FROM pengajian ORDER BY start_time ASC").fetchall()
    recent_absensi = conn.execute("""
        SELECT a.*, s.full_name as santri_name, s.kelas, p.title as pengajian_title
        FROM absensi a JOIN santri s ON a.santri_id = s.id LEFT JOIN pengajian p ON a.pengajian_id = p.id
        ORDER BY a.id DESC LIMIT 6
    """).fetchall()
    conn.close()
    dash_html = """
    <div class="space-y-6">
        <div class="bg-gradient-to-r from-emerald-900 to-teal-800 rounded-2xl p-6 text-white">
            <h1 class="text-2xl font-extrabold">{{ session.get('full_name') }}</h1>
            <p class="text-xs text-emerald-100 mt-1">{{ settings.ponpes_name }} &bull; {{ today_str }}</p>
        </div>
        <div class="grid grid-cols-1 sm:grid-cols-4 gap-4">
            <div class="bg-white p-5 rounded-2xl border"><div class="text-xs font-bold text-slate-400">Santri Aktif</div><div class="text-2xl font-extrabold mt-1">{{ total_santri }}</div></div>
            <div class="bg-white p-5 rounded-2xl border"><div class="text-xs font-bold text-slate-400">Pengurus & Guru</div><div class="text-2xl font-extrabold mt-1">{{ total_pengurus }}</div></div>
            <div class="bg-white p-5 rounded-2xl border"><div class="text-xs font-bold text-slate-400">Absensi Hari Ini</div><div class="text-2xl font-extrabold mt-1">{{ hadir_hari_ini }}</div></div>
            <div class="bg-white p-5 rounded-2xl border"><div class="text-xs font-bold text-slate-400">Total Pembayaran</div><div class="text-xl font-extrabold text-emerald-700 mt-1">Rp {{ "{:,.0f}".format(total_pemasukan).replace(',', '.') }}</div></div>
        </div>
        <div class="bg-white rounded-2xl border p-6">
            <h2 class="font-bold mb-3">Jadwal Pengajian Aktif</h2>
            <div class="grid grid-cols-1 md:grid-cols-2 gap-3">
                {% for j in jadwal_list %}
                <div class="p-3.5 rounded-xl bg-slate-50 border flex justify-between items-center text-xs">
                    <div><div class="font-bold">{{ j.title }}</div><div class="text-slate-500">{{ j.mapel }}</div></div>
                    <span class="px-2.5 py-1 rounded-lg bg-emerald-100 text-emerald-800 font-bold">{{ j.start_time }} - {{ j.end_time }}</span>
                </div>
                {% endfor %}
            </div>
        </div>
    </div>
    """
    return render_page("Dashboard", "dashboard", dash_html, today_str=today_str, total_santri=total_santri, total_pengurus=total_pengurus, hadir_hari_ini=hadir_hari_ini, total_pemasukan=total_pemasukan, jadwal_list=jadwal_list, recent_absensi=recent_absensi)


@app.route("/admin/santri", methods=["GET", "POST"])
@role_required("admin")
def admin_santri():
    conn = get_db()
    if request.method == "POST":
        nis = request.form.get("nis", "").strip()
        full_name = request.form.get("full_name", "").strip()
        gender = request.form.get("gender", "Laki-laki")
        asrama = request.form.get("asrama", "").strip()
        kelas = request.form.get("kelas", "").strip()
        address = request.form.get("address", "").strip()
        wali_name = request.form.get("wali_name", "").strip()
        wali_phone = request.form.get("wali_phone", "").strip()
        rfid_uid = request.form.get("rfid_uid", "").strip() or f"RFID-{secrets.token_hex(4).upper()}"
        ortu_user_id = request.form.get("ortu_user_id") or None
        photo_data = process_uploaded_image(request.files.get("photo"))
        qr_val = f"SANTRI-{nis}-{secrets.token_hex(2).upper()}"
        try:
            conn.execute("""
                INSERT INTO santri (nis, full_name, gender, asrama, kelas, address, wali_name, wali_phone, rfid_uid, qr_code_val, photo_data, ortu_user_id, status, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Aktif', ?)
            """, (nis, full_name, gender, asrama, kelas, address, wali_name, wali_phone, rfid_uid, qr_val, photo_data, ortu_user_id, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
            conn.commit()
            conn.close()
            return redirect(url_for("admin_santri", msg=f"Santri {full_name} berhasil ditambahkan!", type="success"))
        except sqlite3.IntegrityError:
            conn.close()
            return redirect(url_for("admin_santri", msg="NIS atau RFID sudah terdaftar!", type="error"))

    santri_list = conn.execute("SELECT s.*, u.username as ortu_username FROM santri s LEFT JOIN users u ON s.ortu_user_id = u.id ORDER BY s.id DESC").fetchall()
    ortu_accounts = conn.execute("SELECT id, username, full_name FROM users WHERE role = 'ortu'").fetchall()
    conn.close()
    html = """
    <div class="space-y-6">
        <div class="flex flex-wrap items-center justify-between gap-3">
            <h1 class="text-2xl font-extrabold">🎓 Manajemen Data Santri</h1>
            <div class="flex gap-2">
                <a href="{{ url_for('download_template_santri') }}" class="px-3.5 py-2 rounded-xl bg-slate-100 font-bold text-xs border">📥 Template Excel</a>
                <form method="POST" action="{{ url_for('import_santri_excel') }}" enctype="multipart/form-data" class="flex items-center gap-1">
                    <input type="file" name="excel_file" accept=".xlsx" required class="text-xs border rounded-lg p-1 bg-white">
                    <button type="submit" class="px-3 py-2 rounded-xl bg-amber-500 text-white font-bold text-xs">📤 Impor Excel</button>
                </form>
            </div>
        </div>
        <form method="POST" enctype="multipart/form-data" class="bg-white p-5 rounded-2xl border grid grid-cols-1 sm:grid-cols-4 gap-3 text-xs">
            <input type="text" name="nis" required placeholder="NIS *" class="px-3 py-2 rounded-xl border">
            <input type="text" name="full_name" required placeholder="Nama Santri *" class="px-3 py-2 rounded-xl border">
            <select name="gender" class="px-3 py-2 rounded-xl border"><option value="Laki-laki">Laki-laki</option><option value="Perempuan">Perempuan</option></select>
            <input type="text" name="kelas" placeholder="Kelas" class="px-3 py-2 rounded-xl border">
            <input type="text" name="asrama" placeholder="Asrama" class="px-3 py-2 rounded-xl border">
            <input type="text" name="rfid_uid" placeholder="Kode RFID (Opsional)" class="px-3 py-2 rounded-xl border font-mono">
            <input type="text" name="wali_name" placeholder="Nama Wali" class="px-3 py-2 rounded-xl border">
            <select name="ortu_user_id" class="px-3 py-2 rounded-xl border">
                <option value="">-- Tautkan Akun Ortu --</option>
                {% for o in ortu_accounts %}<option value="{{ o.id }}">{{ o.full_name }} (@{{ o.username }})</option>{% endfor %}
            </select>
            <input type="text" name="address" placeholder="Alamat" class="px-3 py-2 rounded-xl border sm:col-span-2">
            <input type="file" name="photo" accept="image/*" class="text-xs border rounded-xl p-1.5">
            <button type="submit" class="px-4 py-2 rounded-xl bg-emerald-700 text-white font-bold">+ Tambah Santri</button>
        </form>
        <div class="bg-white rounded-2xl border overflow-x-auto">
            <table class="w-full text-left text-xs">
                <thead class="bg-slate-50 border-b text-slate-500 uppercase"><tr><th class="p-3">NIS & Nama</th><th class="p-3">Kelas & Asrama</th><th class="p-3">RFID / QR</th><th class="p-3">Akun Ortu</th><th class="p-3 text-right">Aksi</th></tr></thead>
                <tbody class="divide-y">
                    {% for s in santri_list %}
                    <tr>
                        <td class="p-3 font-bold">{{ s.full_name }} <span class="text-slate-400 font-mono">({{ s.nis }})</span></td>
                        <td class="p-3">{{ s.kelas }} - {{ s.asrama }}</td>
                        <td class="p-3 font-mono text-[11px]">{{ s.rfid_uid }} / {{ s.qr_code_val }}</td>
                        <td class="p-3">{{ '@' + s.ortu_username if s.ortu_username else '-' }}</td>
                        <td class="p-3 text-right space-x-1">
                            <a href="{{ url_for('kartu_santri', santri_id=s.id) }}" class="px-2 py-1 rounded bg-emerald-50 text-emerald-700 font-bold">🪪 Kartu</a>
                            <button type="button" @click="askDelete('Hapus Santri?', 'Hapus {{ s.full_name }}?', '{{ url_for('delete_santri', sid=s.id) }}')" class="px-2 py-1 rounded bg-rose-50 text-rose-700 font-bold">Hapus</button>
                        </td>
                    </tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
    </div>
    """
    return render_page("Manajemen Santri", "santri", html, santri_list=santri_list, ortu_accounts=ortu_accounts)


@app.route("/admin/santri/delete/<int:sid>", methods=["POST"])
@role_required("admin")
def delete_santri(sid):
    conn = get_db()
    conn.execute("DELETE FROM santri WHERE id = ?", (sid,))
    conn.commit()
    conn.close()
    return redirect(url_for("admin_santri", msg="Santri dihapus.", type="success"))


@app.route("/template/santri.xlsx")
@role_required("admin")
def download_template_santri():
    if not HAS_OPENPYXL:
        return redirect(url_for("admin_santri", msg="Install openpyxl terlebih dahulu.", type="error"))
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["NIS", "Nama_Lengkap", "Jenis_Kelamin", "Kelas", "Asrama", "Alamat", "Nama_Wali", "No_HP_Wali", "RFID_UID"])
    ws.append(["2026101", "Muhammad Zayd", "Laki-laki", "1 Ulya", "Asrama Al-Farabi 01", "Jl. Kenanga No. 10", "Bpk. Hendra", "081234567890", ""])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="Template_Import_Santri.xlsx")


@app.route("/import/santri", methods=["POST"])
@role_required("admin")
def import_santri_excel():
    if not HAS_OPENPYXL:
        return redirect(url_for("admin_santri", msg="openpyxl tidak tersedia.", type="error"))
    f = request.files.get("excel_file")
    wb = openpyxl.load_workbook(f, data_only=True)
    ws = wb.active
    conn = get_db()
    inserted = 0
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not row[0] or not row[1]:
            continue
        nis = str(row[0]).strip()
        rfid = str(row[8]).strip() if (len(row) > 8 and row[8]) else f"RFID-{secrets.token_hex(4).upper()}"
        qr = f"SANTRI-{nis}-{secrets.token_hex(2).upper()}"
        try:
            conn.execute("""
                INSERT INTO santri (nis, full_name, gender, kelas, asrama, address, wali_name, wali_phone, rfid_uid, qr_code_val, status, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Aktif', ?)
            """, (nis, str(row[1]).strip(), str(row[2] or "Laki-laki"), str(row[3] or ""), str(row[4] or ""), str(row[5] or ""), str(row[6] or ""), str(row[7] or ""), rfid, qr, now_str))
            inserted += 1
        except sqlite3.IntegrityError:
            continue
    conn.commit()
    conn.close()
    return redirect(url_for("admin_santri", msg=f"Berhasil impor {inserted} santri!", type="success"))


@app.route("/admin/users", methods=["GET", "POST"])
@role_required("admin")
def admin_users():
    conn = get_db()
    if request.method == "POST":
        full_name = request.form.get("full_name", "").strip()
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "").strip() or "ponpes123"
        role = request.form.get("role", "pengurus")
        jabatan = request.form.get("jabatan", "").strip()
        phone = request.form.get("phone", "").strip()
        link_santri_id = request.form.get("link_santri_id") or None
        photo_data = process_uploaded_image(request.files.get("photo"))
        try:
            cur = conn.execute("""
                INSERT INTO users (username, password_hash, full_name, role, jabatan, phone, photo_data, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (username, generate_password_hash(password), full_name, role, jabatan, phone, photo_data, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
            if role == "ortu" and link_santri_id:
                conn.execute("UPDATE santri SET ortu_user_id = ? WHERE id = ?", (cur.lastrowid, link_santri_id))
            conn.commit()
            conn.close()
            return redirect(url_for("admin_users", msg=f"Akun {full_name} berhasil dibuat!", type="success"))
        except sqlite3.IntegrityError:
            conn.close()
            return redirect(url_for("admin_users", msg="Username sudah dipakai!", type="error"))

    users_list = conn.execute("SELECT * FROM users ORDER BY role ASC, id DESC").fetchall()
    santri_all = conn.execute("SELECT id, nis, full_name FROM santri ORDER BY full_name ASC").fetchall()
    conn.close()
    html = """
    <div class="space-y-6">
        <div class="flex flex-wrap items-center justify-between gap-3">
            <h1 class="text-2xl font-extrabold">👥 Manajemen Pengurus, Guru & Orang Tua</h1>
            <div class="flex gap-2">
                <a href="{{ url_for('download_template_pengurus') }}" class="px-3.5 py-2 rounded-xl bg-slate-100 font-bold text-xs border">📥 Template Excel</a>
                <form method="POST" action="{{ url_for('import_pengurus_excel') }}" enctype="multipart/form-data" class="flex items-center gap-1">
                    <input type="file" name="excel_file" accept=".xlsx" required class="text-xs border rounded-lg p-1 bg-white">
                    <button type="submit" class="px-3 py-2 rounded-xl bg-amber-500 text-white font-bold text-xs">📤 Impor Excel</button>
                </form>
            </div>
        </div>
        <form method="POST" enctype="multipart/form-data" class="bg-white p-5 rounded-2xl border grid grid-cols-1 sm:grid-cols-4 gap-3 text-xs">
            <select name="role" class="px-3 py-2 rounded-xl border font-bold">
                <option value="pengurus">Pengurus</option><option value="guru">Guru</option><option value="ortu">Orang Tua</option><option value="admin">Admin</option>
            </select>
            <input type="text" name="full_name" required placeholder="Nama Lengkap *" class="px-3 py-2 rounded-xl border">
            <input type="text" name="username" required placeholder="Username *" class="px-3 py-2 rounded-xl border">
            <input type="text" name="password" required placeholder="Password *" class="px-3 py-2 rounded-xl border">
            <input type="text" name="jabatan" placeholder="Jabatan" class="px-3 py-2 rounded-xl border">
            <input type="text" name="phone" placeholder="No HP" class="px-3 py-2 rounded-xl border">
            <select name="link_santri_id" class="px-3 py-2 rounded-xl border">
                <option value="">-- Pilih Anak (Jika Ortu) --</option>
                {% for s in santri_all %}<option value="{{ s.id }}">{{ s.full_name }} ({{ s.nis }})</option>{% endfor %}
            </select>
            <button type="submit" class="px-4 py-2 rounded-xl bg-emerald-700 text-white font-bold">+ Buat Akun</button>
        </form>
        <div class="bg-white rounded-2xl border overflow-x-auto">
            <table class="w-full text-left text-xs">
                <thead class="bg-slate-50 border-b text-slate-500 uppercase"><tr><th class="p-3">Nama</th><th class="p-3">Username</th><th class="p-3">Role</th><th class="p-3">Jabatan</th><th class="p-3 text-right">Aksi</th></tr></thead>
                <tbody class="divide-y">
                    {% for u in users_list %}
                    <tr>
                        <td class="p-3 font-bold">{{ u.full_name }}</td>
                        <td class="p-3 font-mono">@{{ u.username }}</td>
                        <td class="p-3 uppercase font-bold">{{ u.role }}</td>
                        <td class="p-3">{{ u.jabatan or '-' }}</td>
                        <td class="p-3 text-right">
                            {% if u.id != session.get('user_id') %}
                            <button type="button" @click="askDelete('Hapus Akun?', 'Hapus {{ u.full_name }}?', '{{ url_for('delete_user', uid=u.id) }}')" class="px-2 py-1 rounded bg-rose-50 text-rose-700 font-bold">Hapus</button>
                            {% endif %}
                        </td>
                    </tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
    </div>
    """
    return render_page("Manajemen Akun", "users", html, users_list=users_list, santri_all=santri_all)


@app.route("/admin/users/delete/<int:uid>", methods=["POST"])
@role_required("admin")
def delete_user(uid):
    if uid != session.get("user_id"):
        conn = get_db()
        conn.execute("DELETE FROM users WHERE id = ?", (uid,))
        conn.commit()
        conn.close()
    return redirect(url_for("admin_users", msg="Akun dihapus.", type="success"))


@app.route("/template/pengurus.xlsx")
@role_required("admin")
def download_template_pengurus():
    if not HAS_OPENPYXL:
        return redirect(url_for("admin_users", msg="Install openpyxl.", type="error"))
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Nama_Lengkap", "Username", "Password", "Role_pengurus_guru_ortu", "Jabatan", "No_HP"])
    ws.append(["Ust. Ahmad Syauqi", "ust_syauqi", "syauqi123", "guru", "Pengampu Kitab", "08122334455"])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="Template_Import_Pengurus_Guru.xlsx")


@app.route("/import/pengurus", methods=["POST"])
@role_required("admin")
def import_pengurus_excel():
    if not HAS_OPENPYXL:
        return redirect(url_for("admin_users", msg="openpyxl tidak tersedia.", type="error"))
    f = request.files.get("excel_file")
    wb = openpyxl.load_workbook(f, data_only=True)
    ws = wb.active
    conn = get_db()
    inserted = 0
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not row[0] or not row[1]:
            continue
        role = str(row[3] or "pengurus").strip().lower()
        if role not in ("admin", "guru", "pengurus", "ortu"):
            role = "pengurus"
        try:
            conn.execute("""
                INSERT INTO users (username, password_hash, full_name, role, jabatan, phone, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (str(row[1]).strip(), generate_password_hash(str(row[2] or "ponpes123")), str(row[0]).strip(), role, str(row[4] or ""), str(row[5] or ""), now_str))
            inserted += 1
        except sqlite3.IntegrityError:
            continue
    conn.commit()
    conn.close()
    return redirect(url_for("admin_users", msg=f"Berhasil impor {inserted} akun!", type="success"))


@app.route("/admin/pengajian", methods=["GET", "POST"])
@role_required("admin")
def admin_pengajian():
    conn = get_db()
    if request.method == "POST":
        conn.execute("""
            INSERT INTO pengajian (title, mapel, ustadz_name, start_time, end_time, late_after_minutes, days, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (request.form.get("title"), request.form.get("mapel"), request.form.get("ustadz_name"), request.form.get("start_time"), request.form.get("end_time"), int(request.form.get("late_after_minutes", 15)), request.form.get("days", "Setiap Hari"), datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        conn.commit()
        conn.close()
        return redirect(url_for("admin_pengajian", msg="Jadwal pengajian ditambahkan!", type="success"))
    pengajian_list = conn.execute("SELECT * FROM pengajian ORDER BY start_time ASC").fetchall()
    conn.close()
    html = """
    <div class="space-y-6">
        <h1 class="text-2xl font-extrabold">📖 Pengaturan Pengajian & Waktu Absensi</h1>
        <form method="POST" class="bg-white p-5 rounded-2xl border grid grid-cols-1 sm:grid-cols-3 gap-3 text-xs">
            <input type="text" name="title" required placeholder="Nama Sesi Pengajian *" class="px-3 py-2 rounded-xl border">
            <input type="text" name="mapel" required placeholder="Mata Pelajaran / Kitab *" class="px-3 py-2 rounded-xl border">
            <input type="text" name="ustadz_name" placeholder="Nama Ustadz" class="px-3 py-2 rounded-xl border">
            <input type="time" name="start_time" value="04:45" required class="px-3 py-2 rounded-xl border">
            <input type="time" name="end_time" value="06:00" required class="px-3 py-2 rounded-xl border">
            <input type="number" name="late_after_minutes" value="15" placeholder="Toleransi Telat (Menit)" class="px-3 py-2 rounded-xl border">
            <button type="submit" class="sm:col-span-3 py-2.5 rounded-xl bg-emerald-700 text-white font-bold">+ Tambah Jadwal Pengajian</button>
        </form>
        <div class="bg-white rounded-2xl border divide-y">
            {% for p in pengajian_list %}
            <div class="p-4 flex justify-between items-center text-xs">
                <div><div class="font-bold text-sm">{{ p.title }}</div><div class="text-slate-500">{{ p.mapel }} &bull; {{ p.start_time }} - {{ p.end_time }} (Telat >{{ p.late_after_minutes }}m)</div></div>
                <button type="button" @click="askDelete('Hapus Jadwal?', 'Hapus {{ p.title }}?', '{{ url_for('delete_pengajian', pid=p.id) }}')" class="px-3 py-1 rounded bg-rose-50 text-rose-700 font-bold">Hapus</button>
            </div>
            {% endfor %}
        </div>
    </div>
    """
    return render_page("Pengajian & Waktu Absensi", "pengajian", html, pengajian_list=pengajian_list)


@app.route("/admin/pengajian/delete/<int:pid>", methods=["POST"])
@role_required("admin")
def delete_pengajian(pid):
    conn = get_db()
    conn.execute("DELETE FROM pengajian WHERE id = ?", (pid,))
    conn.commit()
    conn.close()
    return redirect(url_for("admin_pengajian", msg="Jadwal dihapus.", type="success"))


@app.route("/absensi", methods=["GET", "POST"])
@role_required("admin", "pengurus", "guru")
def absensi_page():
    conn = get_db()
    if request.method == "POST":
        conn.execute("""
            INSERT INTO absensi (santri_id, pengajian_id, tanggal, waktu_scan, status, metode, recorded_by, notes, created_at)
            VALUES (?, ?, ?, ?, ?, 'Manual', ?, ?, ?)
        """, (request.form.get("santri_id"), request.form.get("pengajian_id") or None, request.form.get("tanggal") or date.today().strftime("%Y-%m-%d"), datetime.now().strftime("%H:%M:%S"), request.form.get("status", "Hadir"), session.get("full_name"), request.form.get("notes", ""), datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        conn.commit()
        conn.close()
        return redirect(url_for("absensi_page", msg="Absensi manual dicatat!", type="success"))
    pengajian_list = conn.execute("SELECT * FROM pengajian ORDER BY start_time ASC").fetchall()
    santri_list = conn.execute("SELECT id, nis, full_name, kelas FROM santri WHERE status='Aktif' ORDER BY full_name ASC").fetchall()
    absensi_logs = conn.execute("""
        SELECT a.*, s.full_name as santri_name, s.kelas, p.title as pengajian_title
        FROM absensi a JOIN santri s ON a.santri_id = s.id LEFT JOIN pengajian p ON a.pengajian_id = p.id
        ORDER BY a.id DESC LIMIT 40
    """).fetchall()
    conn.close()
    html = """
    <div class="space-y-6" x-data="{ code: '', pg: '{{ pengajian_list[0].id if pengajian_list else '' }}', res: null, async scan(){ if(!this.code)return; const r=await fetch('{{ url_for('api_scan_absensi') }}',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({code:this.code,pengajian_id:this.pg})}); this.res=await r.json(); this.code=''; if(this.res.ok) setTimeout(()=>location.reload(),900); } }">
        <h1 class="text-2xl font-extrabold">📷 Absensi Santri (QR Code, RFID & Manual)</h1>
        <div class="bg-emerald-950 text-white p-6 rounded-2xl space-y-3">
            <select x-model="pg" class="w-full px-3 py-2 rounded-xl text-slate-900 text-xs font-bold">
                <option value="">-- Absensi Umum --</option>
                {% for p in pengajian_list %}<option value="{{ p.id }}">{{ p.title }} ({{ p.start_time }} - {{ p.end_time }})</option>{% endfor %}
            </select>
            <form @submit.prevent="scan()" class="flex gap-2">
                <input type="text" x-model="code" autofocus placeholder="Tap kartu RFID atau Scan QR Code di sini..." class="flex-1 px-4 py-3 rounded-xl text-slate-900 font-mono text-sm">
                <button type="submit" class="px-5 py-3 rounded-xl bg-amber-400 text-emerald-950 font-bold text-sm">Scan</button>
            </form>
            <div x-show="res" class="p-3 rounded-xl bg-white/10 text-xs" x-text="res ? (res.title + ' - ' + res.detail) : ''"></div>
        </div>
        <form method="POST" class="bg-white p-5 rounded-2xl border grid grid-cols-1 sm:grid-cols-4 gap-3 text-xs">
            <select name="santri_id" required class="px-3 py-2 rounded-xl border">{% for s in santri_list %}<option value="{{ s.id }}">{{ s.full_name }} ({{ s.nis }})</option>{% endfor %}</select>
            <select name="pengajian_id" class="px-3 py-2 rounded-xl border"><option value="">-- Umum --</option>{% for p in pengajian_list %}<option value="{{ p.id }}">{{ p.title }}</option>{% endfor %}</select>
            <select name="status" class="px-3 py-2 rounded-xl border font-bold"><option value="Hadir">Hadir</option><option value="Terlambat">Terlambat</option><option value="Izin">Izin</option><option value="Sakit">Sakit</option><option value="Alpha">Alpha</option></select>
            <button type="submit" class="px-4 py-2 rounded-xl bg-emerald-700 text-white font-bold">Simpan Manual</button>
        </form>
        <div class="bg-white rounded-2xl border overflow-x-auto">
            <table class="w-full text-left text-xs">
                <thead class="bg-slate-50 border-b"><tr><th class="p-3">Waktu</th><th class="p-3">Santri</th><th class="p-3">Kegiatan</th><th class="p-3">Metode</th><th class="p-3">Status</th></tr></thead>
                <tbody class="divide-y">
                    {% for a in absensi_logs %}
                    <tr><td class="p-3">{{ a.tanggal }} {{ a.waktu_scan }}</td><td class="p-3 font-bold">{{ a.santri_name }}</td><td class="p-3">{{ a.pengajian_title or 'Umum' }}</td><td class="p-3">{{ a.metode }}</td><td class="p-3 font-bold">{{ a.status }}</td></tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
    </div>
    """
    return render_page("Absensi QR & RFID", "absensi", html, pengajian_list=pengajian_list, santri_list=santri_list, absensi_logs=absensi_logs)


@app.route("/api/scan-absensi", methods=["POST"])
@role_required("admin", "pengurus", "guru")
def api_scan_absensi():
    data = request.get_json(silent=True) or {}
    code = str(data.get("code", "")).strip()
    pengajian_id = data.get("pengajian_id") or None
    conn = get_db()
    santri = conn.execute("SELECT * FROM santri WHERE rfid_uid = ? OR qr_code_val = ? OR nis = ?", (code, code, code)).fetchone()
    if not santri:
        conn.close()
        return jsonify({"ok": False, "title": "Tidak Dikenal", "detail": f"Kode {code} tidak ditemukan."})
    metode = "RFID" if santri["rfid_uid"] == code else "QR Code"
    now = datetime.now()
    status = "Hadir"
    if pengajian_id:
        pg = conn.execute("SELECT * FROM pengajian WHERE id = ?", (pengajian_id,)).fetchone()
        if pg:
            try:
                sh, sm = map(int, pg["start_time"].split(":"))
                start_dt = now.replace(hour=sh, minute=sm, second=0, microsecond=0)
                if (now - start_dt).total_seconds() / 60.0 > int(pg["late_after_minutes"] or 15):
                    status = "Terlambat"
            except Exception:
                pass
    conn.execute("""
        INSERT INTO absensi (santri_id, pengajian_id, tanggal, waktu_scan, status, metode, recorded_by, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (santri["id"], pengajian_id, now.strftime("%Y-%m-%d"), now.strftime("%H:%M:%S"), status, metode, session.get("full_name"), now.strftime("%Y-%m-%d %H:%M:%S")))
    conn.commit()
    conn.close()
    return jsonify({"ok": True, "title": f"{santri['full_name']} ({status})", "detail": f"Via {metode} pukul {now.strftime('%H:%M:%S')}"})


@app.route("/progres-pengajian", methods=["GET", "POST"])
@role_required("admin", "pengurus", "guru")
def progres_page():
    conn = get_db()
    if request.method == "POST":
        conn.execute("""
            INSERT INTO progres_pengajian (santri_id, tanggal, jenis, materi, nilai, catatan, recorded_by, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (request.form.get("santri_id"), date.today().strftime("%Y-%m-%d"), request.form.get("jenis"), request.form.get("materi"), request.form.get("nilai"), request.form.get("catatan"), session.get("full_name"), datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        conn.commit()
        conn.close()
        return redirect(url_for("progres_page", msg="Progres setoran disimpan!", type="success"))
    santri_list = conn.execute("SELECT id, nis, full_name FROM santri WHERE status='Aktif'").fetchall()
    progres_list = conn.execute("SELECT pr.*, s.full_name as santri_name FROM progres_pengajian pr JOIN santri s ON pr.santri_id = s.id ORDER BY pr.id DESC LIMIT 40").fetchall()
    conn.close()
    html = """
    <div class="space-y-6">
        <h1 class="text-2xl font-extrabold">🕌 Progres Pengajian & Setoran Santri</h1>
        <form method="POST" class="bg-white p-5 rounded-2xl border grid grid-cols-1 sm:grid-cols-4 gap-3 text-xs">
            <select name="santri_id" required class="px-3 py-2 rounded-xl border">{% for s in santri_list %}<option value="{{ s.id }}">{{ s.full_name }}</option>{% endfor %}</select>
            <select name="jenis" class="px-3 py-2 rounded-xl border"><option value="Setoran Al-Qur'an">Setoran Al-Qur'an</option><option value="Kitab Kuning">Kitab Kuning</option></select>
            <input type="text" name="materi" required placeholder="Surah / Juz / Kitab *" class="px-3 py-2 rounded-xl border">
            <input type="text" name="nilai" placeholder="Nilai / Predikat" class="px-3 py-2 rounded-xl border">
            <input type="text" name="catatan" placeholder="Catatan" class="px-3 py-2 rounded-xl border sm:col-span-3">
            <button type="submit" class="px-4 py-2 rounded-xl bg-emerald-700 text-white font-bold">Simpan Progres</button>
        </form>
        <div class="bg-white rounded-2xl border divide-y">
            {% for pr in progres_list %}
            <div class="p-4 text-xs"><div class="font-bold">{{ pr.santri_name }} - {{ pr.jenis }}: {{ pr.materi }} ({{ pr.nilai }})</div><div class="text-slate-500">{{ pr.tanggal }} &bull; {{ pr.catatan }}</div></div>
            {% endfor %}
        </div>
    </div>
    """
    return render_page("Progres Pengajian", "progres", html, santri_list=santri_list, progres_list=progres_list)


@app.route("/pelanggaran", methods=["GET", "POST"])
@role_required("admin", "pengurus", "guru")
def pelanggaran_page():
    conn = get_db()
    if request.method == "POST":
        conn.execute("""
            INSERT INTO pelanggaran (santri_id, nama_pelanggaran, kategori, poin, tanggal, tindakan, recorded_by, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (request.form.get("santri_id"), request.form.get("nama_pelanggaran"), request.form.get("kategori", "Ringan"), int(request.form.get("poin", 5)), request.form.get("tanggal") or date.today().strftime("%Y-%m-%d"), request.form.get("tindakan"), session.get("full_name"), datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        conn.commit()
        conn.close()
        return redirect(url_for("pelanggaran_page", msg="Pelanggaran dicatat!", type="warning"))
    santri_list = conn.execute("SELECT id, full_name FROM santri WHERE status='Aktif'").fetchall()
    pelanggaran_list = conn.execute("SELECT pl.*, s.full_name as santri_name FROM pelanggaran pl JOIN santri s ON pl.santri_id = s.id ORDER BY pl.id DESC LIMIT 40").fetchall()
    conn.close()
    html = """
    <div class="space-y-6">
        <h1 class="text-2xl font-extrabold">⚠️ Pelanggaran Santri</h1>
        <form method="POST" class="bg-white p-5 rounded-2xl border grid grid-cols-1 sm:grid-cols-4 gap-3 text-xs">
            <select name="santri_id" required class="px-3 py-2 rounded-xl border">{% for s in santri_list %}<option value="{{ s.id }}">{{ s.full_name }}</option>{% endfor %}</select>
            <input type="text" name="nama_pelanggaran" required placeholder="Nama Pelanggaran *" class="px-3 py-2 rounded-xl border">
            <input type="date" name="tanggal" value="{{ today_str }}" class="px-3 py-2 rounded-xl border">
            <input type="text" name="tindakan" placeholder="Tindakan / Pembinaan" class="px-3 py-2 rounded-xl border">
            <button type="submit" class="sm:col-span-4 py-2.5 rounded-xl bg-rose-600 text-white font-bold">Simpan Pelanggaran</button>
        </form>
        <div class="bg-white rounded-2xl border divide-y">
            {% for pl in pelanggaran_list %}
            <div class="p-4 text-xs"><div class="font-bold text-rose-700">{{ pl.santri_name }} - {{ pl.nama_pelanggaran }}</div><div class="text-slate-500">{{ pl.tanggal }} &bull; Pembinaan: {{ pl.tindakan }}</div></div>
            {% endfor %}
        </div>
    </div>
    """
    return render_page("Pelanggaran", "pelanggaran", html, santri_list=santri_list, pelanggaran_list=pelanggaran_list, today_str=date.today().strftime("%Y-%m-%d"))


@app.route("/pembayaran", methods=["GET", "POST"])
@role_required("admin", "pengurus", "guru")
def pembayaran_page():
    conn = get_db()
    if request.method == "POST":
        form_type = request.form.get("form_type", "input_bayar")
        if form_type == "add_jenis" and session.get("role") == "admin":
            conn.execute("INSERT INTO jenis_pembayaran (nama_tagihan, nominal, periode, created_at) VALUES (?, ?, 'Bulanan', ?)", (request.form.get("nama_tagihan"), int(request.form.get("nominal", 0)), datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
            conn.commit()
            conn.close()
            return redirect(url_for("pembayaran_page", msg="Jenis pembayaran ditambahkan!", type="success"))
        if form_type == "input_bayar":
            jenis = conn.execute("SELECT * FROM jenis_pembayaran WHERE id = ?", (request.form.get("jenis_id"),)).fetchone()
            dibayar = int(request.form.get("dibayar_amount", 0))
            tagihan = jenis["nominal"] if jenis else dibayar
            status = "Lunas" if dibayar >= tagihan else ("Cicilan" if dibayar > 0 else "Belum Lunas")
            conn.execute("""
                INSERT INTO pembayaran (kode_transaksi, santri_id, jenis_id, nama_tagihan_snapshot, periode_tagihan, tagihan_amount, dibayar_amount, status, metode, tanggal_bayar, recorded_by, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (f"TRX-{secrets.token_hex(3).upper()}", request.form.get("santri_id"), request.form.get("jenis_id"), jenis["nama_tagihan"] if jenis else "Pembayaran", request.form.get("periode_tagihan"), tagihan, dibayar, status, request.form.get("metode", "Tunai"), date.today().strftime("%Y-%m-%d"), session.get("full_name"), datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
            conn.commit()
            conn.close()
            return redirect(url_for("pembayaran_page", msg="Pembayaran tersimpan & langsung terupdate di akun Orang Tua!", type="success"))
    jenis_list = conn.execute("SELECT * FROM jenis_pembayaran ORDER BY id DESC").fetchall()
    santri_list = conn.execute("SELECT id, nis, full_name FROM santri WHERE status='Aktif'").fetchall()
    pembayaran_list = conn.execute("SELECT pb.*, s.full_name as santri_name FROM pembayaran pb JOIN santri s ON pb.santri_id = s.id ORDER BY pb.id DESC LIMIT 40").fetchall()
    conn.close()
    html = """
    <div class="space-y-6">
        <h1 class="text-2xl font-extrabold">💳 Pembayaran Pondok Pesantren</h1>
        <form method="POST" class="bg-white p-5 rounded-2xl border grid grid-cols-1 sm:grid-cols-4 gap-3 text-xs">
            <input type="hidden" name="form_type" value="input_bayar">
            <select name="santri_id" required class="px-3 py-2 rounded-xl border">{% for s in santri_list %}<option value="{{ s.id }}">{{ s.full_name }} ({{ s.nis }})</option>{% endfor %}</select>
            <select name="jenis_id" required class="px-3 py-2 rounded-xl border">{% for j in jenis_list %}<option value="{{ j.id }}">{{ j.nama_tagihan }} (Rp {{ j.nominal }})</option>{% endfor %}</select>
            <input type="text" name="periode_tagihan" required placeholder="Bulan / Periode *" class="px-3 py-2 rounded-xl border">
            <input type="number" name="dibayar_amount" required placeholder="Nominal Dibayar (Rp) *" class="px-3 py-2 rounded-xl border font-bold">
            <button type="submit" class="sm:col-span-4 py-2.5 rounded-xl bg-emerald-700 text-white font-bold">Simpan Pembayaran (Sinkron Real-Time ke Ortu)</button>
        </form>
        {% if session.get('role') == 'admin' %}
        <form method="POST" class="bg-white p-4 rounded-2xl border flex gap-2 text-xs">
            <input type="hidden" name="form_type" value="add_jenis">
            <input type="text" name="nama_tagihan" required placeholder="Nama Tagihan Baru" class="flex-1 px-3 py-2 rounded-xl border">
            <input type="number" name="nominal" required placeholder="Nominal (Rp)" class="px-3 py-2 rounded-xl border">
            <button type="submit" class="px-4 py-2 rounded-xl bg-amber-500 text-white font-bold">+ Tambah Jenis Tagihan</button>
        </form>
        {% endif %}
        <div class="bg-white rounded-2xl border overflow-x-auto">
            <table class="w-full text-left text-xs">
                <thead class="bg-slate-50 border-b"><tr><th class="p-3">Kode</th><th class="p-3">Santri</th><th class="p-3">Tagihan</th><th class="p-3">Dibayar</th><th class="p-3">Status</th></tr></thead>
                <tbody class="divide-y">
                    {% for pb in pembayaran_list %}
                    <tr><td class="p-3 font-mono">{{ pb.kode_transaksi }}</td><td class="p-3 font-bold">{{ pb.santri_name }}</td><td class="p-3">{{ pb.nama_tagihan_snapshot }} ({{ pb.periode_tagihan }})</td><td class="p-3 font-bold text-emerald-700">Rp {{ "{:,.0f}".format(pb.dibayar_amount).replace(',', '.') }}</td><td class="p-3 font-bold">{{ pb.status }}</td></tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
    </div>
    """
    return render_page("Pembayaran", "pembayaran", html, jenis_list=jenis_list, santri_list=santri_list, pembayaran_list=pembayaran_list)


@app.route("/ujian", methods=["GET", "POST"])
@role_required("admin", "pengurus", "guru")
def ujian_page():
    conn = get_db()
    if request.method == "POST":
        ft = request.form.get("form_type")
        if ft == "create_ujian" and session.get("role") == "admin":
            conn.execute("INSERT INTO ujian (judul_ujian, mapel, kelas, tanggal, created_by, created_at) VALUES (?, ?, ?, ?, ?, ?)", (request.form.get("judul_ujian"), request.form.get("mapel"), request.form.get("kelas"), date.today().strftime("%Y-%m-%d"), session.get("full_name"), datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        elif ft == "add_soal":
            conn.execute("INSERT INTO soal_ujian (ujian_id, nomor_soal, pertanyaan, opsi_a, opsi_b, opsi_c, opsi_d, kunci_jawaban, created_by) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)", (request.form.get("ujian_id"), int(request.form.get("nomor_soal", 1)), request.form.get("pertanyaan"), request.form.get("opsi_a"), request.form.get("opsi_b"), request.form.get("opsi_c"), request.form.get("opsi_d"), request.form.get("kunci_jawaban", "A"), session.get("full_name")))
        elif ft == "input_nilai":
            conn.execute("INSERT INTO nilai_ujian (ujian_id, santri_id, skor, catatan, graded_by, created_at) VALUES (?, ?, ?, ?, ?, ?)", (request.form.get("ujian_id"), request.form.get("santri_id"), float(request.form.get("skor", 0)), request.form.get("catatan"), session.get("full_name"), datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        conn.commit()
        conn.close()
        return redirect(url_for("ujian_page", msg="Data ujian diperbarui!", type="success"))
    ujian_list = conn.execute("SELECT * FROM ujian ORDER BY id DESC").fetchall()
    soal_list = conn.execute("SELECT su.*, u.judul_ujian FROM soal_ujian su JOIN ujian u ON su.ujian_id = u.id ORDER BY su.id DESC").fetchall()
    santri_list = conn.execute("SELECT id, full_name FROM santri WHERE status='Aktif'").fetchall()
    conn.close()
    html = """
    <div class="space-y-6">
        <h1 class="text-2xl font-extrabold">📝 Manajemen Ujian, Bank Soal & Nilai</h1>
        {% if session.get('role') == 'admin' %}
        <form method="POST" class="bg-white p-5 rounded-2xl border grid grid-cols-1 sm:grid-cols-4 gap-3 text-xs">
            <input type="hidden" name="form_type" value="create_ujian">
            <input type="text" name="judul_ujian" required placeholder="Judul Ujian *" class="px-3 py-2 rounded-xl border">
            <input type="text" name="mapel" required placeholder="Mata Pelajaran *" class="px-3 py-2 rounded-xl border">
            <input type="text" name="kelas" required placeholder="Kelas *" class="px-3 py-2 rounded-xl border">
            <button type="submit" class="px-4 py-2 rounded-xl bg-emerald-700 text-white font-bold">+ Buat Ujian (Admin)</button>
        </form>
        {% endif %}
        <form method="POST" class="bg-white p-5 rounded-2xl border grid grid-cols-1 sm:grid-cols-3 gap-3 text-xs">
            <input type="hidden" name="form_type" value="add_soal">
            <select name="ujian_id" required class="px-3 py-2 rounded-xl border">{% for u in ujian_list %}<option value="{{ u.id }}">{{ u.judul_ujian }} ({{ u.mapel }})</option>{% endfor %}</select>
            <input type="number" name="nomor_soal" value="1" class="px-3 py-2 rounded-xl border">
            <select name="kunci_jawaban" class="px-3 py-2 rounded-xl border"><option value="A">Kunci A</option><option value="B">Kunci B</option><option value="C">Kunci C</option><option value="D">Kunci D</option></select>
            <input type="text" name="pertanyaan" required placeholder="Pertanyaan Soal *" class="px-3 py-2 rounded-xl border sm:col-span-3">
            <input type="text" name="opsi_a" placeholder="Opsi A" class="px-3 py-2 rounded-xl border">
            <input type="text" name="opsi_b" placeholder="Opsi B" class="px-3 py-2 rounded-xl border">
            <input type="text" name="opsi_c" placeholder="Opsi C" class="px-3 py-2 rounded-xl border">
            <input type="text" name="opsi_d" placeholder="Opsi D" class="px-3 py-2 rounded-xl border">
            <button type="submit" class="sm:col-span-2 py-2 rounded-xl bg-amber-500 text-white font-bold">+ Tambah Soal Ujian</button>
        </form>
        <form method="POST" class="bg-white p-5 rounded-2xl border grid grid-cols-1 sm:grid-cols-4 gap-3 text-xs">
            <input type="hidden" name="form_type" value="input_nilai">
            <select name="ujian_id" required class="px-3 py-2 rounded-xl border">{% for u in ujian_list %}<option value="{{ u.id }}">{{ u.judul_ujian }}</option>{% endfor %}</select>
            <select name="santri_id" required class="px-3 py-2 rounded-xl border">{% for s in santri_list %}<option value="{{ s.id }}">{{ s.full_name }}</option>{% endfor %}</select>
            <input type="number" step="0.1" name="skor" required placeholder="Nilai (0-100) *" class="px-3 py-2 rounded-xl border font-bold">
            <button type="submit" class="px-4 py-2 rounded-xl bg-blue-600 text-white font-bold">Simpan Nilai</button>
        </form>
        <div class="bg-white rounded-2xl border divide-y">
            {% for sl in soal_list %}
            <div class="p-4 text-xs"><div class="font-bold">{{ sl.judul_ujian }} (#{{ sl.nomor_soal }}) - Kunci: {{ sl.kunci_jawaban }}</div><div>{{ sl.pertanyaan }}</div></div>
            {% endfor %}
        </div>
    </div>
    """
    return render_page("Ujian & Soal", "ujian", html, ujian_list=ujian_list, soal_list=soal_list, santri_list=santri_list)


@app.route("/kartu-santri")
@login_required
def kartu_santri():
    conn = get_db()
    if session.get("role") == "ortu":
        rows = conn.execute("SELECT * FROM santri WHERE ortu_user_id = ?", (session.get("user_id"),)).fetchall()
    elif request.args.get("santri_id"):
        rows = conn.execute("SELECT * FROM santri WHERE id = ?", (request.args.get("santri_id"),)).fetchall()
    else:
        rows = conn.execute("SELECT * FROM santri WHERE status='Aktif' LIMIT 20").fetchall()
    conn.close()
    cards = []
    for r in rows:
        d = dict(r)
        d["qr_img"] = generate_qr_base64(d["qr_code_val"] or d["nis"])
        cards.append(d)
    html = """
    <div class="space-y-6">
        <div class="flex justify-between items-center no-print">
            <h1 class="text-2xl font-extrabold">🪪 Kartu Tanda Santri Bolak-Balik</h1>
            <button onclick="window.print()" class="px-5 py-2.5 rounded-xl bg-emerald-700 text-white font-bold text-xs">🖨️ Cetak Kartu</button>
        </div>
        {% for s in cards %}
        <div class="grid grid-cols-1 md:grid-cols-2 gap-6 bg-white p-5 rounded-2xl border">
            <div class="rounded-2xl bg-gradient-to-br from-emerald-950 to-teal-900 text-white p-5 flex flex-col justify-between">
                <div class="font-extrabold text-xs text-amber-300 border-b border-emerald-800 pb-2">{{ settings.ponpes_name }}</div>
                <div class="flex items-center justify-between gap-3 my-4">
                    <div class="text-xs space-y-1">
                        <div class="font-extrabold text-base">{{ s.full_name }}</div>
                        <div>NIS: <span class="font-mono text-amber-300">{{ s.nis }}</span></div>
                        <div>Kelas: {{ s.kelas }} &bull; Asrama: {{ s.asrama }}</div>
                        <div class="font-mono text-[10px] text-amber-300">RFID: {{ s.rfid_uid }}</div>
                    </div>
                    {% if s.qr_img %}<img src="{{ s.qr_img }}" class="w-20 h-20 bg-white p-1 rounded-xl">{% endif %}
                </div>
                <div class="text-[10px] text-emerald-300 font-mono">{{ s.qr_code_val }}</div>
            </div>
            <div class="rounded-2xl bg-slate-50 border p-5 flex flex-col justify-between text-xs">
                <div>
                    <div class="font-extrabold text-center border-b pb-2 mb-2">TATA TERTIB KARTU SANTRI</div>
                    <ol class="list-decimal list-inside space-y-1 text-slate-600 text-[11px]">
                        <li>Kartu resmi santri {{ settings.ponpes_name }}.</li>
                        <li>Wajib dibawa untuk absensi QR Code & RFID setiap kegiatan.</li>
                        <li>Jika menemukan kartu ini harap kembalikan ke sekretariat pondok.</li>
                    </ol>
                </div>
                <div class="text-right text-[11px] pt-3 border-t">Pimpinan Pondok,<br><strong>{{ settings.pimpinan_name }}</strong></div>
            </div>
        </div>
        {% endfor %}
    </div>
    """
    return render_page("Kartu Santri", "kartu", html, cards=cards)


@app.route("/orang-tua")
@role_required("ortu")
def ortu_monitoring():
    conn = get_db()
    children = conn.execute("SELECT * FROM santri WHERE ortu_user_id = ?", (session.get("user_id"),)).fetchall()
    if not children:
        conn.close()
        return render_page("Monitoring Anak", "ortu", "<div class='bg-white p-8 rounded-2xl text-center'>Akun belum ditautkan ke data santri. Hubungi Admin.</div>")
    anak = children[0]
    absensi_list = conn.execute("SELECT a.*, p.title as pengajian_title FROM absensi a LEFT JOIN pengajian p ON a.pengajian_id = p.id WHERE a.santri_id = ? ORDER BY a.id DESC", (anak["id"],)).fetchall()
    pembayaran_list = conn.execute("SELECT * FROM pembayaran WHERE santri_id = ? ORDER BY id DESC", (anak["id"],)).fetchall()
    progres_list = conn.execute("SELECT * FROM progres_pengajian WHERE santri_id = ? ORDER BY id DESC", (anak["id"],)).fetchall()
    pelanggaran_list = conn.execute("SELECT * FROM pelanggaran WHERE santri_id = ? ORDER BY id DESC", (anak["id"],)).fetchall()
    nilai_list = conn.execute("SELECT nu.*, u.judul_ujian, u.mapel FROM nilai_ujian nu JOIN ujian u ON nu.ujian_id = u.id WHERE nu.santri_id = ? ORDER BY nu.id DESC", (anak["id"],)).fetchall()
    conn.close()
    html = """
    <div class="space-y-6">
        <div class="bg-emerald-950 text-white p-6 rounded-2xl flex justify-between items-center">
            <div><h1 class="text-xl font-extrabold">{{ anak.full_name }}</h1><p class="text-xs text-emerald-200">NIS: {{ anak.nis }} &bull; Kelas: {{ anak.kelas }} &bull; Asrama: {{ anak.asrama }}</p></div>
            <a href="{{ url_for('kartu_santri', santri_id=anak.id) }}" class="px-4 py-2 rounded-xl bg-amber-400 text-emerald-950 font-bold text-xs">🪪 Kartu Santri</a>
        </div>
        <div class="bg-white rounded-2xl border p-5">
            <h2 class="font-bold mb-3">💳 Status Pembayaran Pondok (Real-Time)</h2>
            {% for pb in pembayaran_list %}<div class="p-3 border-b text-xs flex justify-between"><span>{{ pb.nama_tagihan_snapshot }} ({{ pb.periode_tagihan }})</span><span class="font-bold text-emerald-700">Rp {{ "{:,.0f}".format(pb.dibayar_amount).replace(',', '.') }} - {{ pb.status }}</span></div>{% else %}<p class="text-xs text-slate-400">Belum ada pembayaran.</p>{% endfor %}
        </div>
        <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div class="bg-white rounded-2xl border p-5"><h2 class="font-bold mb-3">🕌 Progres Pengajian & Hafalan</h2>{% for pr in progres_list %}<div class="p-2.5 border-b text-xs"><b>{{ pr.jenis }}</b>: {{ pr.materi }} ({{ pr.nilai }})</div>{% endfor %}</div>
            <div class="bg-white rounded-2xl border p-5"><h2 class="font-bold mb-3">📅 Kehadiran</h2>{% for a in absensi_list %}<div class="p-2.5 border-b text-xs flex justify-between"><span>{{ a.pengajian_title or 'Kegiatan' }} ({{ a.tanggal }})</span><b>{{ a.status }}</b></div>{% endfor %}</div>
            <div class="bg-white rounded-2xl border p-5"><h2 class="font-bold mb-3">⚠️ Pelanggaran</h2>{% for pl in pelanggaran_list %}<div class="p-2.5 border-b text-xs text-rose-700"><b>{{ pl.nama_pelanggaran }}</b> ({{ pl.tanggal }})</div>{% else %}<p class="text-xs text-emerald-700">Tidak ada pelanggaran.</p>{% endfor %}</div>
            <div class="bg-white rounded-2xl border p-5"><h2 class="font-bold mb-3">🏅 Nilai Ujian</h2>{% for nl in nilai_list %}<div class="p-2.5 border-b text-xs flex justify-between"><span>{{ nl.judul_ujian }} ({{ nl.mapel }})</span><b>{{ nl.skor }}</b></div>{% endfor %}</div>
        </div>
    </div>
    """
    return render_page("Monitoring Anak", "ortu", html, anak=anak, absensi_list=absensi_list, pembayaran_list=pembayaran_list, progres_list=progres_list, pelanggaran_list=pelanggaran_list, nilai_list=nilai_list)


@app.route("/admin/settings", methods=["GET", "POST"])
@role_required("admin")
def admin_settings():
    conn = get_db()
    if request.method == "POST":
        ft = request.form.get("form_type")
        if ft == "ponpes":
            logo_data = process_uploaded_image(request.files.get("logo"), max_size=(300, 300))
            if logo_data:
                conn.execute("UPDATE settings SET ponpes_name=?, ponpes_tagline=?, ponpes_address=?, ponpes_phone=?, pimpinan_name=?, logo_data=? WHERE id=1", (request.form.get("ponpes_name"), request.form.get("ponpes_tagline"), request.form.get("ponpes_address"), request.form.get("ponpes_phone"), request.form.get("pimpinan_name"), logo_data))
            else:
                conn.execute("UPDATE settings SET ponpes_name=?, ponpes_tagline=?, ponpes_address=?, ponpes_phone=?, pimpinan_name=? WHERE id=1", (request.form.get("ponpes_name"), request.form.get("ponpes_tagline"), request.form.get("ponpes_address"), request.form.get("ponpes_phone"), request.form.get("pimpinan_name")))
            conn.commit()
            conn.close()
            return redirect(url_for("admin_settings", msg="Identitas Ponpes diperbarui!", type="success"))
        elif ft == "admin_account":
            uid = session.get("user_id")
            conn.execute("UPDATE users SET username=?, full_name=? WHERE id=?", (request.form.get("username"), request.form.get("full_name"), uid))
            if request.form.get("password"):
                conn.execute("UPDATE users SET password_hash=? WHERE id=?", (generate_password_hash(request.form.get("password")), uid))
            conn.commit()
            session["username"] = request.form.get("username")
            session["full_name"] = request.form.get("full_name")
            conn.close()
            return redirect(url_for("admin_settings", msg="Akun Admin diperbarui!", type="success"))
    conn.close()
    html = """
    <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
        <form method="POST" enctype="multipart/form-data" class="bg-white p-6 rounded-2xl border space-y-3 text-xs">
            <h2 class="font-bold text-base">🕌 Identitas & Logo Ponpes</h2>
            <input type="hidden" name="form_type" value="ponpes">
            <input type="text" name="ponpes_name" value="{{ settings.ponpes_name }}" required class="w-full px-3 py-2 rounded-xl border">
            <input type="text" name="ponpes_tagline" value="{{ settings.ponpes_tagline or '' }}" class="w-full px-3 py-2 rounded-xl border">
            <input type="text" name="pimpinan_name" value="{{ settings.pimpinan_name or '' }}" class="w-full px-3 py-2 rounded-xl border">
            <input type="text" name="ponpes_phone" value="{{ settings.ponpes_phone or '' }}" class="w-full px-3 py-2 rounded-xl border">
            <input type="text" name="ponpes_address" value="{{ settings.ponpes_address or '' }}" class="w-full px-3 py-2 rounded-xl border">
            <input type="file" name="logo" accept="image/*" class="w-full text-xs">
            <button type="submit" class="w-full py-2.5 rounded-xl bg-emerald-700 text-white font-bold">Simpan Identitas</button>
        </form>
        <form method="POST" class="bg-white p-6 rounded-2xl border space-y-3 text-xs h-fit">
            <h2 class="font-bold text-base">🔐 Akun Admin Sendiri</h2>
            <input type="hidden" name="form_type" value="admin_account">
            <input type="text" name="full_name" value="{{ session.get('full_name') }}" required class="w-full px-3 py-2 rounded-xl border">
            <input type="text" name="username" value="{{ session.get('username') }}" required class="w-full px-3 py-2 rounded-xl border">
            <input type="password" name="password" placeholder="Password Baru (Kosongkan jika tetap)" class="w-full px-3 py-2 rounded-xl border">
            <button type="submit" class="w-full py-2.5 rounded-xl bg-slate-900 text-white font-bold">Simpan Akun Admin</button>
        </form>
    </div>
    """
    return render_page("Pengaturan", "settings", html)


init_db()

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
