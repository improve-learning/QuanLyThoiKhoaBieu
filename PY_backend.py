
import functools
import hashlib
import hmac
import os
import re
import secrets
import shutil
import sqlite3
import sys
import zipfile
import xml.etree.ElementTree as ET
from contextlib import contextmanager
from datetime import datetime

# =====================================================================
# ĐƯỜNG DẪN (tính từ vị trí file này, không phụ thuộc thư mục chạy)
# =====================================================================
# Khi đóng gói bằng PyInstaller (--onefile), __file__ nằm trong thư mục tạm bị xóa khi thoát app
# -> dữ liệu người dùng phải nằm cạnh file .exe.
THU_MUC_DU_AN = (os.path.dirname(sys.executable) if getattr(sys, "frozen", False)
                 else os.path.dirname(os.path.abspath(__file__)))
DB_PATH = os.path.join(THU_MUC_DU_AN, "database.db")
THU_MUC_KHO = os.path.join(THU_MUC_DU_AN, "kho_luu_tru")
THU_MUC_NHAC = os.path.join(THU_MUC_DU_AN, "nhac_chuong")

MAT_KHAU_TOI_THIEU = 8
MAT_KHAU_TOI_DA = 128
BUOI_SANG, BUOI_CHIEU = "Sáng", "Chiều"
TIET_TOI_DA = 5                       # số tiết tối đa trong một buổi (sáng hoặc chiều)
GIOI_HAN_XLSX = 20 * 1024 * 1024      # tối đa 20 MB / tệp .xlsx tải lên kho
_KHOI_DOC = 1024 * 1024               # đọc/ghi tệp theo khối 1 MB
_EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

# Phiên đăng nhập hiện tại (một người dùng tại một thời điểm)
_phien = {"MaND": None}


# =====================================================================
# TIỆN ÍCH CHUNG
# =====================================================================
def _ok(data=None, message=""):
    return {"ok": True, "data": data, "message": message}


def _loi(message, data=None):
    return {"ok": False, "data": data, "message": message}


def _api(can_dang_nhap=True):
    """Bọc hàm: kiểm tra đăng nhập, bắt lỗi SQLite / file để không làm chương trình crash."""
    def deco(fn):
        @functools.wraps(fn)
        def wrapper(*a, **kw):
            try:
                if can_dang_nhap:
                    if _phien["MaND"] is None:
                        return _loi("Bạn chưa đăng nhập.")
                    return fn(_phien["MaND"], *a, **kw)
                return fn(*a, **kw)
            except sqlite3.Error as e:
                return _loi(f"Lỗi cơ sở dữ liệu: {e}")
            except OSError as e:
                return _loi(f"Lỗi thao tác tệp: {e}")
        return wrapper
    return deco


def _ket_noi():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


@contextmanager
def _giao_dich():
    """Mở kết nối, tự commit nếu thành công, rollback nếu lỗi, luôn đóng kết nối."""
    conn = _ket_noi()
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def _chuan_hoa_ngay(s):
    """'20/10/2026' hoặc '2026-10-20' -> '2026-10-20'. Sai định dạng -> None."""
    s = str(s or "").strip()
    for dang in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, dang).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return None


def _chuan_hoa_gio(s):
    """'7:00' hoặc '07:00' -> '07:00'. Sai định dạng -> None."""
    s = str(s or "").strip()
    for dang in ("%H:%M", "%H:%M:%S"):
        try:
            return datetime.strptime(s, dang).strftime("%H:%M")
        except ValueError:
            pass
    return None


def _ngay_hien_thi(iso):
    try:
        return datetime.strptime(iso, "%Y-%m-%d").strftime("%d/%m/%Y")
    except (ValueError, TypeError):
        return iso


def _thu(iso):
    """0 = Thứ Hai ... 6 = Chủ Nhật (khớp danh sách THU trong giao diện)."""
    try:
        return datetime.strptime(iso, "%Y-%m-%d").weekday()
    except (ValueError, TypeError):
        return None


# =====================================================================
# KHỞI TẠO / NÂNG CẤP CƠ SỞ DỮ LIỆU (không làm mất dữ liệu cũ)
# =====================================================================
def _chuan_ten(s):
    """Gộp khoảng trắng + không phân biệt hoa/thường (đúng cả với tiếng Việt có dấu)."""
    return " ".join(str(s or "").split()).casefold()


def _ma_ke_tiep(c, bang, cot, tien_to):
    """Sinh mã kế tiếp dạng MH0001, GV0001, ... (không bao giờ trùng mã cũ)."""
    r = c.execute(
        f"SELECT MAX(CAST(SUBSTR({cot}, {len(tien_to) + 1}) AS INTEGER)) FROM {bang} WHERE {cot} LIKE ?",
        (tien_to + "%",)).fetchone()
    return f"{tien_to}{(r[0] or 0) + 1:04d}"


def _tim_mon(c, ten):
    k = _chuan_ten(ten)
    for r in c.execute("SELECT MaMH, TenMH FROM MONHOC"):
        if _chuan_ten(r["TenMH"]) == k:
            return r["MaMH"]
    return None


def _tim_gv(c, ten, ma_mh):
    """Giảng viên được xác định bởi (tên, môn dạy) vì GIANGVIEN.MaMH là môn giảng viên đó dạy."""
    k = _chuan_ten(ten)
    for r in c.execute("SELECT MaGV, TenGV FROM GIANGVIEN WHERE MaMH=?", (ma_mh,)):
        if _chuan_ten(r["TenGV"]) == k:
            return r["MaGV"]
    return None


def _lay_hoac_tao_mon(c, ten):
    ma = _tim_mon(c, ten)
    if ma is None:
        ma = _ma_ke_tiep(c, "MONHOC", "MaMH", "MH")
        c.execute("INSERT INTO MONHOC (MaMH, TenMH, SoTC) VALUES (?,?,NULL)", (ma, " ".join(ten.split())))
    return ma


def _lay_hoac_tao_gv(c, ten, ma_mh):
    ma = _tim_gv(c, ten, ma_mh)
    if ma is None:
        ma = _ma_ke_tiep(c, "GIANGVIEN", "MaGV", "GV")
        c.execute("INSERT INTO GIANGVIEN (MaGV, TenGV, MaMH) VALUES (?,?,?)",
                  (ma, " ".join(ten.split()), ma_mh))
    return ma


def _chuan_phong(phong):
    return " ".join(str(phong or "").split())


def _lay_hoac_tao_phong(c, phong):
    """Phòng đã có trong PHONGHOC (không phân biệt hoa/thường) thì dùng lại, chưa có thì thêm mới."""
    k = _chuan_ten(phong)
    for r in c.execute("SELECT MaPhong FROM PHONGHOC"):
        if _chuan_ten(r["MaPhong"]) == k:
            return r["MaPhong"]
    c.execute("INSERT INTO PHONGHOC (MaPhong) VALUES (?)", (phong,))
    return phong


def _tao_bang_tkb(c):
    c.execute("""CREATE TABLE IF NOT EXISTS MONHOC (
        MaMH  TEXT PRIMARY KEY,
        TenMH TEXT NOT NULL,
        SoTC  INTEGER)""")
    c.execute("""CREATE TABLE IF NOT EXISTS GIANGVIEN (
        MaGV  TEXT PRIMARY KEY,
        TenGV TEXT NOT NULL,
        MaMH  TEXT NOT NULL REFERENCES MONHOC(MaMH))""")
    c.execute("""CREATE TABLE IF NOT EXISTS PHONGHOC (
        MaPhong TEXT PRIMARY KEY)""")
    c.execute(f"""CREATE TABLE IF NOT EXISTS BUOIHOC (
        MaBH    INTEGER PRIMARY KEY AUTOINCREMENT,
        MaND    INTEGER REFERENCES NguoiDung(MaND) ON DELETE CASCADE,
        MaMH    TEXT NOT NULL REFERENCES MONHOC(MaMH),
        MaGV    TEXT REFERENCES GIANGVIEN(MaGV),
        NgayHoc TEXT NOT NULL,
        Buoi    TEXT NOT NULL CHECK (Buoi IN ('{BUOI_SANG}', '{BUOI_CHIEU}')),
        SoTiet  INTEGER NOT NULL CHECK (SoTiet > 0),
        MaPhong TEXT REFERENCES PHONGHOC(MaPhong))""")
    c.execute("CREATE INDEX IF NOT EXISTS idx_buoihoc_nd ON BUOIHOC(MaND, NgayHoc)")
    c.execute("CREATE INDEX IF NOT EXISTS idx_gv_mon ON GIANGVIEN(MaMH)")
    if c.execute("SELECT COUNT(*) FROM PHONGHOC").fetchone()[0] == 0:
        c.executemany("INSERT INTO PHONGHOC (MaPhong) VALUES (?)",
                      [(f"H{tang}{so:02d}",) for tang in (1, 2) for so in range(1, 6)])


def _so_tiet_tu_gio(g1, g2):
    """Ước lượng số tiết của buổi học cũ từ giờ bắt đầu/kết thúc (≈ 50 phút / tiết)."""
    try:
        a = datetime.strptime(g1, "%H:%M")
        b = datetime.strptime(g2, "%H:%M")
        return max(1, min(TIET_TOI_DA, round((b - a).total_seconds() / 60 / 50)))
    except (ValueError, TypeError):
        return 1


def _chuyen_buoi_hoc_cu(c):
    """BuoiHoc(MaHP, MaGV, MaPhong, NgayHoc, GioBatDau, GioKetThuc) -> MONHOC/GIANGVIEN/PHONGHOC/BUOIHOC."""
    for r in c.execute("SELECT * FROM BuoiHoc_cu ORDER BY MaBuoi").fetchall():
        ngay = _chuan_hoa_ngay(r["NgayHoc"])
        mon = " ".join(str(r["MaHP"] or "").split())
        if not ngay or not mon:
            continue                                   # dòng hỏng: không đủ thông tin để chuyển
        g1, g2 = _chuan_hoa_gio(r["GioBatDau"]), _chuan_hoa_gio(r["GioKetThuc"])
        buoi = BUOI_SANG if (g1 or "00:00") < "12:00" else BUOI_CHIEU
        ma_mh = _lay_hoac_tao_mon(c, mon)
        gv = " ".join(str(r["MaGV"] or "").split())
        ma_gv = _lay_hoac_tao_gv(c, gv, ma_mh) if gv else None
        phong = " ".join(str(r["MaPhong"] or "").split()) or None
        if phong:
            c.execute("INSERT OR IGNORE INTO PHONGHOC (MaPhong) VALUES (?)", (phong,))
        c.execute("INSERT INTO BUOIHOC (MaBH, MaND, MaMH, MaGV, NgayHoc, Buoi, SoTiet, MaPhong) "
                  "VALUES (?,?,?,?,?,?,?,?)",
                  (r["MaBuoi"], r["MaND"], ma_mh, ma_gv, ngay, buoi, _so_tiet_tu_gio(g1, g2), phong))


def _dong_bo_tai_khoan(c, uid, email_cu=None):
    """Ghi (hoặc cập nhật) dòng TaiKhoan(Email, TenNguoiDung, MatKhau) theo NguoiDung của uid.
    email_cu: email trước khi đổi (để xóa dòng cũ, vì Email là khóa chính của TaiKhoan)."""
    r = c.execute("SELECT HoTen, Email, MatKhau FROM NguoiDung WHERE MaND=?", (uid,)).fetchone()
    if not r:
        return
    if email_cu:
        c.execute("DELETE FROM TaiKhoan WHERE Email=?", (email_cu,))
    c.execute("INSERT OR REPLACE INTO TaiKhoan (Email, TenNguoiDung, MatKhau) VALUES (?,?,?)",
              (r["Email"], r["HoTen"], r["MatKhau"]))


def khoi_tao():
    """Tạo/nâng cấp CSDL và thư mục lưu trữ. Gọi một lần khi khởi động ứng dụng (an toàn khi chạy lặp lại)."""
    os.makedirs(THU_MUC_KHO, exist_ok=True)
    os.makedirs(THU_MUC_NHAC, exist_ok=True)

    with _giao_dich() as c:
        ver = c.execute("PRAGMA user_version").fetchone()[0]

        # Sao lưu DB trước khi đổi cấu trúc bảng lịch học (mỗi lần nâng cấp một bản riêng)
        if ver < 3 and os.path.exists(DB_PATH) and os.path.getsize(DB_PATH) > 0:
            bak = DB_PATH + ".truoc_v3.bak"
            if not os.path.exists(bak):
                shutil.copy2(DB_PATH, bak)

        c.execute("BEGIN")   # để cả quá trình nâng cấp (kể cả ALTER/DROP) rollback được nếu lỗi

        c.execute("""CREATE TABLE IF NOT EXISTS NguoiDung (
            MaND     INTEGER PRIMARY KEY AUTOINCREMENT,
            HoTen    TEXT NOT NULL,
            Email    TEXT NOT NULL UNIQUE COLLATE NOCASE,
            MatKhau  TEXT NOT NULL)""")
        c.execute("""CREATE TABLE IF NOT EXISTS TaiKhoan (
            Email        TEXT PRIMARY KEY COLLATE NOCASE,
            TenNguoiDung TEXT NOT NULL,
            MatKhau      TEXT NOT NULL)""")
        c.execute("""CREATE TABLE IF NOT EXISTS NhacHen (
            MaNH    INTEGER PRIMARY KEY AUTOINCREMENT,
            MaND    INTEGER NOT NULL REFERENCES NguoiDung(MaND) ON DELETE CASCADE,
            Ngay    TEXT NOT NULL,
            Gio     TEXT NOT NULL,
            NoiDung TEXT NOT NULL)""")
        c.execute("""CREATE TABLE IF NOT EXISTS KhoLuuTru (
            MaFile      INTEGER PRIMARY KEY AUTOINCREMENT,
            MaND        INTEGER NOT NULL REFERENCES NguoiDung(MaND) ON DELETE CASCADE,
            TenGoc      TEXT NOT NULL,
            DuongDan    TEXT NOT NULL,
            ThoiGianTai TEXT NOT NULL,
            Hash        TEXT)""")
        c.execute("""CREATE TABLE IF NOT EXISTS CaiDat (
            MaND       INTEGER PRIMARY KEY REFERENCES NguoiDung(MaND) ON DELETE CASCADE,
            NhacChuong TEXT)""")
        c.execute("CREATE INDEX IF NOT EXISTS idx_nhachen_nd ON NhacHen(MaND, Ngay, Gio)")

        # Nâng cấp kho lưu trữ: thêm cột Hash (SHA-256) để nhận biết file trùng nội dung
        cols_kho = {r["name"] for r in c.execute("PRAGMA table_info(KhoLuuTru)")}
        if "Hash" not in cols_kho:
            c.execute("ALTER TABLE KhoLuuTru ADD COLUMN Hash TEXT")
        c.execute("CREATE INDEX IF NOT EXISTS idx_kho_hash ON KhoLuuTru(MaND, Hash)")

        # Bảng BuoiHoc kiểu cũ (có cột MaHP) -> đổi tên để tạo BUOIHOC mới rồi chuyển dữ liệu sang
        # (SQLite không phân biệt hoa/thường nên BuoiHoc và BUOIHOC là cùng một tên bảng)
        cu = {r["name"] for r in c.execute("PRAGMA table_info(BuoiHoc)")}
        co_bang_cu = "MaHP" in cu
        if co_bang_cu:
            c.execute("DROP INDEX IF EXISTS idx_buoihoc_nd")
            c.execute("ALTER TABLE BuoiHoc RENAME TO BuoiHoc_cu")
        _tao_bang_tkb(c)
        if co_bang_cu:
            _chuyen_buoi_hoc_cu(c)
            c.execute("DROP TABLE BuoiHoc_cu")

        if ver < 2:
            # Tính Hash cho các file đã có trong kho (mỗi file đọc đúng một lần)
            for r in c.execute("SELECT MaFile, DuongDan FROM KhoLuuTru WHERE Hash IS NULL").fetchall():
                dd = _tuyet_doi(r["DuongDan"])
                if os.path.isfile(dd):
                    try:
                        c.execute("UPDATE KhoLuuTru SET Hash=? WHERE MaFile=?", (_bam_tep(dd), r["MaFile"]))
                    except OSError:
                        pass
        # Bảng TaiKhoan luôn được đồng bộ lại theo NguoiDung: xóa dòng không còn tài khoản tương ứng
        # (kể cả dòng rỗng), thêm/cập nhật các tài khoản còn thiếu hoặc lệch (mật khẩu vẫn ở dạng băm)
        c.execute("DELETE FROM TaiKhoan WHERE Email NOT IN (SELECT Email FROM NguoiDung)")
        for r in c.execute("SELECT MaND FROM NguoiDung").fetchall():
            _dong_bo_tai_khoan(c, r["MaND"])
        if ver < 4:
            c.execute("PRAGMA user_version = 4")


# =====================================================================
# MẬT KHẨU (scrypt trong thư viện chuẩn, có salt riêng cho mỗi mật khẩu)
# =====================================================================
def _bam_mat_khau(mat_khau):
    n, r, p = 2 ** 14, 8, 1
    salt = secrets.token_bytes(16)
    h = hashlib.scrypt(mat_khau.encode("utf-8"), salt=salt, n=n, r=r, p=p, dklen=32)
    return f"scrypt${n}${r}${p}${salt.hex()}${h.hex()}"


def _kiem_tra_mat_khau(mat_khau, chuoi_bam):
    try:
        _, n, r, p, salt, h = chuoi_bam.split("$")
        tinh = hashlib.scrypt(mat_khau.encode("utf-8"), salt=bytes.fromhex(salt),
                              n=int(n), r=int(r), p=int(p), dklen=32)
        return hmac.compare_digest(tinh, bytes.fromhex(h))
    except (ValueError, TypeError):
        return False


def _kiem_tra_mat_khau_moi(mk):
    if len(mk or "") < MAT_KHAU_TOI_THIEU:
        return f"Mật khẩu phải có ít nhất {MAT_KHAU_TOI_THIEU} ký tự."
    if len(mk) > MAT_KHAU_TOI_DA:
        return f"Mật khẩu tối đa {MAT_KHAU_TOI_DA} ký tự."
    if not re.search(r"[A-Za-z]", mk) or not re.search(r"\d", mk):
        return "Mật khẩu phải gồm cả chữ và số."
    return None


def _kiem_tra_ho_ten_email(ho_ten, email):
    if not (ho_ten or "").strip():
        return "Họ tên không được để trống."
    if not (email or "").strip():
        return "Email không được để trống."
    if not _EMAIL_RE.match(email.strip()):
        return "Email không hợp lệ."
    return None


# =====================================================================
# TÀI KHOẢN: ĐĂNG KÝ / ĐĂNG NHẬP / ĐĂNG XUẤT
# =====================================================================
@_api(can_dang_nhap=False)
def dang_ky(ho_ten, email, mat_khau):
    """Tạo tài khoản. Buổi học cũ chưa có chủ sẽ được gán cho tài khoản ĐẦU TIÊN."""
    loi = _kiem_tra_ho_ten_email(ho_ten, email) or _kiem_tra_mat_khau_moi(mat_khau)
    if loi:
        return _loi(loi)
    ho_ten, email = ho_ten.strip(), email.strip()
    try:
        with _giao_dich() as c:
            if c.execute("SELECT 1 FROM NguoiDung WHERE Email=?", (email,)).fetchone():
                return _loi("Email đã được sử dụng.")
            dau_tien = c.execute("SELECT COUNT(*) FROM NguoiDung").fetchone()[0] == 0
            cur = c.execute("INSERT INTO NguoiDung(HoTen, Email, MatKhau) VALUES (?,?,?)",
                            (ho_ten, email, _bam_mat_khau(mat_khau)))
            ma = cur.lastrowid
            c.execute("INSERT INTO CaiDat(MaND) VALUES (?)", (ma,))
            _dong_bo_tai_khoan(c, ma)
            if dau_tien:
                c.execute("UPDATE BUOIHOC SET MaND=? WHERE MaND IS NULL", (ma,))
    except sqlite3.IntegrityError:
        return _loi("Email đã được sử dụng.")
    return _ok({"id": ma, "ho_ten": ho_ten, "email": email}, "Đăng ký thành công.")


@_api(can_dang_nhap=False)
def dang_nhap(email, mat_khau):
    with _giao_dich() as c:
        r = c.execute("SELECT * FROM NguoiDung WHERE Email=?", ((email or "").strip(),)).fetchone()
    if not r or not _kiem_tra_mat_khau(mat_khau or "", r["MatKhau"]):
        return _loi("Email hoặc mật khẩu không đúng.")
    _phien["MaND"] = r["MaND"]
    return _ok({"id": r["MaND"], "ho_ten": r["HoTen"], "email": r["Email"]}, "Đăng nhập thành công.")


def dang_xuat():
    """Chỉ xóa phiên; không xóa bất kỳ dữ liệu nào đã lưu."""
    _phien["MaND"] = None
    return _ok(None, "Đã đăng xuất.")


def da_dang_nhap():
    return _phien["MaND"] is not None


# =====================================================================
# THÔNG TIN CÁ NHÂN
# =====================================================================
@_api()
def cap_nhat_thong_tin(uid, ho_ten, email):
    loi = _kiem_tra_ho_ten_email(ho_ten, email)
    if loi:
        return _loi(loi)
    ho_ten, email = ho_ten.strip(), email.strip()
    try:
        with _giao_dich() as c:
            if c.execute("SELECT 1 FROM NguoiDung WHERE Email=? AND MaND<>?",
                         (email, uid)).fetchone():
                return _loi("Email này đã thuộc về tài khoản khác.")
            cu = c.execute("SELECT Email FROM NguoiDung WHERE MaND=?", (uid,)).fetchone()
            c.execute("UPDATE NguoiDung SET HoTen=?, Email=? WHERE MaND=?", (ho_ten, email, uid))
            _dong_bo_tai_khoan(c, uid, cu["Email"] if cu else None)
    except sqlite3.IntegrityError:
        return _loi("Email này đã thuộc về tài khoản khác.")
    return _ok({"id": uid, "ho_ten": ho_ten, "email": email}, "Đã cập nhật thông tin.")


@_api()
def doi_mat_khau(uid, mat_khau_cu, mat_khau_moi):
    loi = _kiem_tra_mat_khau_moi(mat_khau_moi)
    if loi:
        return _loi(loi)
    with _giao_dich() as c:
        r = c.execute("SELECT MatKhau FROM NguoiDung WHERE MaND=?", (uid,)).fetchone()
        if not r or not _kiem_tra_mat_khau(mat_khau_cu or "", r["MatKhau"]):
            return _loi("Mật khẩu hiện tại không đúng.")
        c.execute("UPDATE NguoiDung SET MatKhau=? WHERE MaND=?", (_bam_mat_khau(mat_khau_moi), uid))
        _dong_bo_tai_khoan(c, uid)
    return _ok(None, "Đã đổi mật khẩu.")


# =====================================================================
# THỜI KHÓA BIỂU  (MONHOC, GIANGVIEN, PHONGHOC, BUOIHOC)
# =====================================================================
def _chuan_hoa_buoi(s):
    k = str(s or "").strip().casefold()
    if k in ("sáng", "sang"):
        return BUOI_SANG
    if k in ("chiều", "chieu"):
        return BUOI_CHIEU
    return None


def _buoi_hoc_dict(r):
    return {
        "id": r["id"], "ma_mh": r["MaMH"], "ma_gv": r["MaGV"],
        "mon_hoc": r["TenMH"], "giang_vien": r["TenGV"] or "",
        "ngay": r["NgayHoc"], "ngay_hien_thi": _ngay_hien_thi(r["NgayHoc"]),
        "thu": _thu(r["NgayHoc"]), "buoi": r["Buoi"],
        "so_tiet": r["SoTiet"], "phong": r["MaPhong"] or "",
    }


def _kiem_tra_buoi_hoc(mon_hoc, giang_vien, ngay, buoi, so_tiet):
    """Trả về (loi, mon, gv, ngay_iso, buoi, so_tiet) — dùng chung cho thêm và sửa."""
    mon, gv = " ".join(str(mon_hoc or "").split()), " ".join(str(giang_vien or "").split())
    if not mon:
        return "Tên môn học không được để trống.", None, None, None, None, None
    if not gv:
        return "Tên giảng viên không được để trống.", None, None, None, None, None
    n = _chuan_hoa_ngay(ngay)
    if not n:
        return "Ngày học không hợp lệ (dùng DD/MM/YYYY).", None, None, None, None, None
    b = _chuan_hoa_buoi(buoi)
    if not b:
        return "Buổi học phải là Sáng hoặc Chiều.", None, None, None, None, None
    try:
        st = int(str(so_tiet).strip())
    except ValueError:
        return "Số tiết phải là số nguyên.", None, None, None, None, None
    if not 1 <= st <= TIET_TOI_DA:
        return f"Số tiết phải từ 1 đến {TIET_TOI_DA}.", None, None, None, None, None
    return None, mon, gv, n, b, st


def _kiem_tra_cho_trong(c, uid, ma_mh, n, buoi, bo_qua_id=None):
    """Kiểm tra buổi (ngày + sáng/chiều) của tài khoản còn chỗ cho môn này.
    Trả về chuỗi lỗi, hoặc None nếu hợp lệ. Quy tắc: mỗi buổi học của một tài khoản chỉ có một môn học duy nhất
    (chỉ xét lịch của chính tài khoản đó). Giới hạn số tiết/buổi được kiểm tra ở _kiem_tra_buoi_hoc."""
    sql = ("SELECT b.MaBH, b.MaMH, m.TenMH FROM BUOIHOC b "
           "JOIN MONHOC m ON m.MaMH = b.MaMH "
           "WHERE b.MaND=? AND b.NgayHoc=? AND b.Buoi=?")
    tham_so = [uid, n, buoi]
    if bo_qua_id is not None:
        sql += " AND b.MaBH<>?"
        tham_so.append(bo_qua_id)
    rows = c.execute(sql, tham_so).fetchall()
    if ma_mh is not None and any(r["MaMH"] == ma_mh for r in rows):
        return f"Môn này đã có trong buổi {buoi.lower()} ngày {_ngay_hien_thi(n)}."
    if rows:
        return (f"Buổi {buoi.lower()} ngày {_ngay_hien_thi(n)} đã có môn "
                f"{rows[0]['TenMH']}.")
    return None


_SQL_TKB = ("SELECT b.MaBH AS id, b.MaMH, b.MaGV, m.TenMH, g.TenGV, b.NgayHoc, b.Buoi, b.SoTiet, b.MaPhong "
            "FROM BUOIHOC b JOIN MONHOC m ON m.MaMH = b.MaMH LEFT JOIN GIANGVIEN g ON g.MaGV = b.MaGV "
            "WHERE b.MaND=?")


@_api()
def lay_tkb(uid, ngay=None):
    """Danh sách buổi học của tài khoản (sắp theo ngày, sáng trước chiều). ngay=None -> tất cả."""
    sql, tham_so = _SQL_TKB, [uid]
    if ngay is not None:
        n = _chuan_hoa_ngay(ngay)
        if not n:
            return _loi("Ngày không hợp lệ.")
        sql += " AND b.NgayHoc=?"
        tham_so.append(n)
    sql += " ORDER BY b.NgayHoc, (b.Buoi=?), b.MaBH"
    tham_so.append(BUOI_CHIEU)
    with _giao_dich() as c:
        rows = c.execute(sql, tham_so).fetchall()
    return _ok([_buoi_hoc_dict(r) for r in rows])


@_api()
def them_buoi_hoc(uid, mon_hoc, giang_vien, ngay, buoi, so_tiet, phong):
    """Thêm buổi học. Môn học / giảng viên / phòng chưa có sẽ tự được thêm vào MONHOC / GIANGVIEN / PHONGHOC."""
    loi, mon, gv, n, b, st = _kiem_tra_buoi_hoc(mon_hoc, giang_vien, ngay, buoi, so_tiet)
    if loi:
        return _loi(loi)
    ph = _chuan_phong(phong)
    if not ph:
        return _loi("Mã phòng học không được để trống.")
    with _giao_dich() as c:
        loi = _kiem_tra_cho_trong(c, uid, _tim_mon(c, mon), n, b)
        if loi:
            return _loi(loi)
        phong = _lay_hoac_tao_phong(c, ph)
        ma_mh = _lay_hoac_tao_mon(c, mon)
        ma_gv = _lay_hoac_tao_gv(c, gv, ma_mh)
        cur = c.execute(
            "INSERT INTO BUOIHOC (MaND, MaMH, MaGV, NgayHoc, Buoi, SoTiet, MaPhong) VALUES (?,?,?,?,?,?,?)",
            (uid, ma_mh, ma_gv, n, b, st, phong))
        ma = cur.lastrowid
    return _ok({"id": ma, "phong": phong}, "Đã thêm buổi học.")


@_api()
def sua_buoi_hoc(uid, ma_buoi, mon_hoc, giang_vien, ngay, buoi, so_tiet, phong):
    """Sửa buổi học (chỉ của chính tài khoản), gồm cả mã phòng."""
    loi, mon, gv, n, b, st = _kiem_tra_buoi_hoc(mon_hoc, giang_vien, ngay, buoi, so_tiet)
    if loi:
        return _loi(loi)
    ph = _chuan_phong(phong)
    if not ph:
        return _loi("Mã phòng học không được để trống.")
    with _giao_dich() as c:
        cu = c.execute("SELECT MaPhong FROM BUOIHOC WHERE MaBH=? AND MaND=?", (ma_buoi, uid)).fetchone()
        if not cu:
            return _loi("Không tìm thấy buổi học.")
        loi = _kiem_tra_cho_trong(c, uid, _tim_mon(c, mon), n, b, bo_qua_id=ma_buoi)
        if loi:
            return _loi(loi)
        phong = _lay_hoac_tao_phong(c, ph)
        ma_mh = _lay_hoac_tao_mon(c, mon)
        ma_gv = _lay_hoac_tao_gv(c, gv, ma_mh)
        c.execute("UPDATE BUOIHOC SET MaMH=?, MaGV=?, NgayHoc=?, Buoi=?, SoTiet=?, MaPhong=? "
                  "WHERE MaBH=? AND MaND=?", (ma_mh, ma_gv, n, b, st, phong, ma_buoi, uid))
    return _ok({"id": ma_buoi, "phong": phong}, "Đã cập nhật buổi học.")


@_api()
def xoa_buoi_hoc(uid, ma_buoi):
    """Chỉ xóa bản ghi buổi học; không đụng tới MONHOC/GIANGVIEN/PHONGHOC và file trong kho lưu trữ."""
    with _giao_dich() as c:
        cur = c.execute("DELETE FROM BUOIHOC WHERE MaBH=? AND MaND=?", (ma_buoi, uid))
        if cur.rowcount == 0:
            return _loi("Không tìm thấy buổi học để xóa.")
    return _ok(None, "Đã xóa buổi học.")


# =====================================================================
# NHẮC HẸN
# =====================================================================
def _nhac_hen_dict(r):
    return {"id": r["MaNH"], "ngay": r["Ngay"], "ngay_hien_thi": _ngay_hien_thi(r["Ngay"]),
            "thu": _thu(r["Ngay"]), "gio": r["Gio"], "noi_dung": r["NoiDung"]}


def _kiem_tra_nhac_hen(ngay, gio, noi_dung):
    """Trả về (loi, ngay_iso, gio, noi_dung_da_cat_khoang_trang). Dùng chung cho thêm và sửa."""
    noi_dung = (noi_dung or "").strip()
    if not noi_dung:
        return "Nội dung nhắc hẹn không được để trống.", None, None, None
    if len(noi_dung) > 500:
        return "Nội dung nhắc hẹn tối đa 500 ký tự.", None, None, None
    n, g = _chuan_hoa_ngay(ngay), _chuan_hoa_gio(gio)
    if not n:
        return "Ngày không hợp lệ (dùng DD/MM/YYYY).", None, None, None
    if not g:
        return "Giờ không hợp lệ (dùng HH:MM).", None, None, None
    return None, n, g, noi_dung


_LOI_NHAC_QUA_KHU = "Thời điểm nhắc hẹn phải sau thời điểm hiện tại (sớm nhất là phút kế tiếp)."


def _da_qua(ngay_iso, gio):
    """Nhắc hẹn chỉ kêu khi sang phút mới -> thời điểm <= phút hiện tại sẽ không bao giờ kêu."""
    return f"{ngay_iso} {gio}" <= datetime.now().strftime("%Y-%m-%d %H:%M")


@_api()
def them_nhac_hen(uid, ngay, gio, noi_dung):
    loi, n, g, noi_dung = _kiem_tra_nhac_hen(ngay, gio, noi_dung)
    if loi:
        return _loi(loi)
    if _da_qua(n, g):
        return _loi(_LOI_NHAC_QUA_KHU)
    with _giao_dich() as c:
        cur = c.execute("INSERT INTO NhacHen (MaND, Ngay, Gio, NoiDung) VALUES (?,?,?,?)",
                        (uid, n, g, noi_dung))
        ma = cur.lastrowid
    return _ok({"id": ma}, "Đã thêm nhắc hẹn.")


@_api()
def sua_nhac_hen(uid, ma_nhac_hen, ngay, gio, noi_dung):
    """Sửa ngày, giờ và nội dung của một nhắc hẹn đã có (chỉ sửa được nhắc hẹn của chính tài khoản)."""
    loi, n, g, noi_dung = _kiem_tra_nhac_hen(ngay, gio, noi_dung)
    if loi:
        return _loi(loi)
    with _giao_dich() as c:
        cu = c.execute("SELECT Ngay, Gio FROM NhacHen WHERE MaNH=? AND MaND=?", (ma_nhac_hen, uid)).fetchone()
        if not cu:
            return _loi("Không tìm thấy nhắc hẹn để sửa.")
        # chỉ chặn khi người dùng ĐỔI sang thời điểm đã qua (vẫn cho sửa nội dung của nhắc hẹn cũ)
        if (cu["Ngay"], cu["Gio"]) != (n, g) and _da_qua(n, g):
            return _loi(_LOI_NHAC_QUA_KHU)
        c.execute("UPDATE NhacHen SET Ngay=?, Gio=?, NoiDung=? WHERE MaNH=? AND MaND=?",
                  (n, g, noi_dung, ma_nhac_hen, uid))
    return _ok({"id": ma_nhac_hen}, "Đã cập nhật nhắc hẹn.")


@_api()
def lay_nhac_hen(uid, ngay=None):
    """Nhắc hẹn của tài khoản (ngày, giờ tăng dần). ngay=None -> tất cả."""
    sql, tham_so = "SELECT * FROM NhacHen WHERE MaND=?", [uid]
    if ngay is not None:
        n = _chuan_hoa_ngay(ngay)
        if not n:
            return _loi("Ngày không hợp lệ.")
        sql += " AND Ngay=?"
        tham_so.append(n)
    sql += " ORDER BY Ngay, Gio, MaNH"
    with _giao_dich() as c:
        rows = c.execute(sql, tham_so).fetchall()
    return _ok([_nhac_hen_dict(r) for r in rows])


@_api()
def xoa_nhac_hen(uid, ma_nhac_hen):
    with _giao_dich() as c:
        cur = c.execute("DELETE FROM NhacHen WHERE MaNH=? AND MaND=?", (ma_nhac_hen, uid))
        if cur.rowcount == 0:
            return _loi("Không tìm thấy nhắc hẹn để xóa.")
    return _ok(None, "Đã xóa nhắc hẹn.")


@_api()
def lay_nhac_hen_den_han(uid, tu_luc, den_luc):
    """Nhắc hẹn có thời điểm trong (tu_luc, den_luc]. Giao diện gọi định kỳ (vd mỗi giây)
    với tu_luc = lần kiểm tra trước, den_luc = bây giờ (datetime) để biết khi nào cần phát nhạc."""
    a, b = tu_luc.strftime("%Y-%m-%d %H:%M"), den_luc.strftime("%Y-%m-%d %H:%M")
    with _giao_dich() as c:
        rows = c.execute(
            "SELECT * FROM NhacHen WHERE MaND=? AND (Ngay || ' ' || Gio) > ? AND (Ngay || ' ' || Gio) <= ? "
            "ORDER BY Ngay, Gio", (uid, a, b)).fetchall()
    return _ok([_nhac_hen_dict(r) for r in rows])


# =====================================================================
# KHO LƯU TRỮ (.xlsx)
# =====================================================================
def _ten_an_toan(ten):
    ten = re.sub(r"[^\w.\- ]", "_", os.path.basename(ten), flags=re.UNICODE)
    return ten[:120] or "tep"


def _ten_luu_duy_nhat(thu_muc, ten_goc):
    """Thêm mốc thời gian + chuỗi ngẫu nhiên để hai file trùng tên không bao giờ ghi đè nhau."""
    while True:
        ten = f"{datetime.now():%Y%m%d_%H%M%S}_{secrets.token_hex(3)}_{_ten_an_toan(ten_goc)}"
        if not os.path.exists(os.path.join(thu_muc, ten)):
            return ten


def _kiem_tra_xlsx(duong_dan):
    """Kiểm tra NHANH cấu trúc .xlsx (không nạp cả workbook, không giải nén toàn bộ file).
    Trả về chuỗi lỗi nếu file không hợp lệ, None nếu hợp lệ."""
    try:
        if not zipfile.is_zipfile(duong_dan):
            return "Tệp không phải Excel .xlsx hợp lệ."
        with zipfile.ZipFile(duong_dan) as z:
            ten = set(z.namelist())
            if "xl/workbook.xml" not in ten or "[Content_Types].xml" not in ten:
                return "Tệp không phải Excel .xlsx hợp lệ."
            if z.getinfo("xl/workbook.xml").file_size > 5 * 1024 * 1024:
                return "Tệp Excel có cấu trúc bất thường."
            goc = ET.fromstring(z.read("xl/workbook.xml"))   # đọc hết mục này -> tự kiểm tra CRC
    except Exception:
        return "Không đọc được tệp Excel (tệp hỏng hoặc không đúng định dạng)."
    if not any(e.tag == "sheet" or e.tag.endswith("}sheet") for e in goc.iter()):
        return "Tệp Excel không có trang tính nào."
    return None


def _bam_tep(duong_dan):
    h = hashlib.sha256()
    with open(duong_dan, "rb") as f:
        for khoi in iter(lambda: f.read(_KHOI_DOC), b""):
            h.update(khoi)
    return h.hexdigest()


def _sao_chep_va_bam(nguon, dich):
    """Sao chép theo khối và tính SHA-256 trong CÙNG MỘT lượt đọc."""
    h = hashlib.sha256()
    with open(nguon, "rb") as fi, open(dich, "wb") as fo:
        for khoi in iter(lambda: fi.read(_KHOI_DOC), b""):
            h.update(khoi)
            fo.write(khoi)
    return h.hexdigest()


def _go_tep(duong_dan):
    try:
        os.remove(duong_dan)
    except OSError:
        pass


def _tuyet_doi(duong_dan_luu):
    if os.path.isabs(duong_dan_luu):
        return duong_dan_luu
    # DB lưu dấu "/" (bản ghi cũ có thể là "\") -> chuyển về dạng đường dẫn của hệ điều hành đang chạy
    return os.path.join(THU_MUC_DU_AN, *duong_dan_luu.replace("\\", "/").split("/"))


@_api()
def tai_file_len(uid, duong_dan):
    """Nhận đường dẫn file .xlsx do giao diện chọn, kiểm tra, sao chép vào kho_luu_tru/ và ghi metadata.
    - Kiểm tra cấu trúc nhanh, giới hạn dung lượng (GIOI_HAN_XLSX).
    - Sao chép + tính SHA-256 trong một lượt; nếu tài khoản đã có file cùng nội dung thì KHÔNG lưu thêm bản sao.
    - File gốc của người dùng không bị đụng tới."""
    nguon = os.path.abspath(str(duong_dan or ""))
    if not os.path.isfile(nguon):
        return _loi("Tệp không tồn tại.")
    if os.path.splitext(nguon)[1].lower() != ".xlsx":
        return _loi("Chỉ chấp nhận tệp Excel có đuôi .xlsx.")
    kich_thuoc = os.path.getsize(nguon)
    if kich_thuoc == 0:
        return _loi("Tệp rỗng.")
    if kich_thuoc > GIOI_HAN_XLSX:
        return _loi(f"Tệp quá lớn (tối đa {GIOI_HAN_XLSX // (1024 * 1024)} MB).")
    loi = _kiem_tra_xlsx(nguon)
    if loi:
        return _loi(loi)

    os.makedirs(THU_MUC_KHO, exist_ok=True)
    dich = os.path.join(THU_MUC_KHO, _ten_luu_duy_nhat(THU_MUC_KHO, os.path.basename(nguon)))
    tam = dich + ".tmp"
    trung_ten = None
    ma = None
    try:
        ma_bam = _sao_chep_va_bam(nguon, tam)
        with _giao_dich() as c:
            for r in c.execute("SELECT TenGoc, DuongDan FROM KhoLuuTru WHERE MaND=? AND Hash=?",
                               (uid, ma_bam)).fetchall():
                if os.path.isfile(_tuyet_doi(r["DuongDan"])):      # bản cũ còn tồn tại -> coi là trùng
                    trung_ten = r["TenGoc"]
                    break
            if trung_ten is None:
                os.replace(tam, dich)                               # đổi tên nguyên tử, không để lại file dở
                cur = c.execute(
                    "INSERT INTO KhoLuuTru (MaND, TenGoc, DuongDan, ThoiGianTai, Hash) VALUES (?,?,?,?,?)",
                    (uid, os.path.basename(nguon), os.path.relpath(dich, THU_MUC_DU_AN).replace(os.sep, "/"),
                     datetime.now().strftime("%Y-%m-%d %H:%M:%S"), ma_bam))
                ma = cur.lastrowid
    except Exception:
        _go_tep(tam)
        _go_tep(dich)             # lưu metadata thất bại -> dọn file vừa sao chép
        raise
    if trung_ten is not None:
        _go_tep(tam)
        return _loi(f"Tệp này đã có trong kho lưu trữ (trùng nội dung với \"{trung_ten}\").")
    return _ok({"id": ma, "ten_file": os.path.basename(nguon), "duong_dan": dich},
               "Đã tải tệp lên kho lưu trữ.")


@_api()
def lay_danh_sach_file(uid):
    """File của tài khoản, mới nhất trước. 'ton_tai' = False nếu file đã mất khỏi thư mục lưu."""
    with _giao_dich() as c:
        rows = c.execute("SELECT * FROM KhoLuuTru WHERE MaND=? ORDER BY ThoiGianTai DESC, MaFile DESC",
                         (uid,)).fetchall()
    ds = []
    for r in rows:
        dd = _tuyet_doi(r["DuongDan"])
        ds.append({"id": r["MaFile"], "ten_file": r["TenGoc"], "thoi_gian_tai": r["ThoiGianTai"],
                   "duong_dan": dd, "ton_tai": os.path.isfile(dd)})
    thieu = sum(1 for f in ds if not f["ton_tai"])
    return _ok(ds, f"{thieu} tệp đã bị mất khỏi thư mục lưu trữ." if thieu else "")


# =====================================================================
# NHẠC CHUÔNG (.mp3)
# =====================================================================
def _kiem_tra_mp3(duong_dan):
    """Trả về chuỗi lỗi nếu không phải MP3 hợp lệ, None nếu hợp lệ."""
    try:
        with open(duong_dan, "rb") as f:
            dau = f.read(10)
            vi_tri = 0
            if dau[:3] == b"ID3" and len(dau) == 10:       # bỏ qua thẻ ID3v2
                kich_thuoc = ((dau[6] & 0x7F) << 21) | ((dau[7] & 0x7F) << 14) | \
                             ((dau[8] & 0x7F) << 7) | (dau[9] & 0x7F)
                vi_tri = 10 + kich_thuoc
            f.seek(vi_tri)
            khoi = f.read(65536)
    except OSError as e:
        return f"Không đọc được tệp: {e}"
    for i in range(len(khoi) - 3):
        # Khung MP3: 11 bit đồng bộ, version != 01, layer != 00
        if khoi[i] == 0xFF and (khoi[i + 1] & 0xE0) == 0xE0 \
                and (khoi[i + 1] & 0x18) != 0x08 and (khoi[i + 1] & 0x06) != 0x00:
            return None
    return "Tệp không phải MP3 hợp lệ."


@_api()
def chon_nhac_chuong(uid, duong_dan):
    """Kiểm tra .mp3, sao chép vào nhac_chuong/, lưu đường dẫn vào cài đặt của tài khoản."""
    nguon = os.path.abspath(str(duong_dan or ""))
    if not os.path.isfile(nguon):
        return _loi("Tệp nhạc không tồn tại.")
    if os.path.splitext(nguon)[1].lower() != ".mp3":
        return _loi("Chỉ chấp nhận tệp có đuôi .mp3.")
    if os.path.getsize(nguon) == 0:
        return _loi("Tệp nhạc rỗng.")
    loi = _kiem_tra_mp3(nguon)
    if loi:
        return _loi(loi)

    os.makedirs(THU_MUC_NHAC, exist_ok=True)
    dich = os.path.join(THU_MUC_NHAC, _ten_luu_duy_nhat(THU_MUC_NHAC, os.path.basename(nguon)))
    shutil.copy2(nguon, dich)
    rel = os.path.relpath(dich, THU_MUC_DU_AN).replace(os.sep, "/")
    try:
        with _giao_dich() as c:
            cu = c.execute("SELECT NhacChuong FROM CaiDat WHERE MaND=?", (uid,)).fetchone()
            c.execute("INSERT INTO CaiDat (MaND, NhacChuong) VALUES (?,?) "
                      "ON CONFLICT(MaND) DO UPDATE SET NhacChuong=excluded.NhacChuong", (uid, rel))
            # Bản sao cũ chỉ còn tài khoản này dùng -> dọn để không tích file rác
            cu_rel = cu["NhacChuong"] if cu else None
            dung_chung = cu_rel and c.execute(
                "SELECT 1 FROM CaiDat WHERE NhacChuong=? AND MaND<>?", (cu_rel, uid)).fetchone()
    except Exception:
        try:
            os.remove(dich)
        except OSError:
            pass
        raise
    if cu_rel and not dung_chung:
        cu_abs = _tuyet_doi(cu_rel)
        if os.path.dirname(os.path.abspath(cu_abs)) == os.path.abspath(THU_MUC_NHAC):
            try:
                os.remove(cu_abs)
            except OSError:
                pass
    return _ok({"duong_dan": dich, "ten_file": os.path.basename(nguon)}, "Đã đặt nhạc chuông.")


@_api()
def lay_nhac_chuong(uid):
    """Nhạc chuông hiện tại; data = None nếu chưa chọn. Giao diện/bộ phát âm dùng 'duong_dan' để phát."""
    with _giao_dich() as c:
        r = c.execute("SELECT NhacChuong FROM CaiDat WHERE MaND=?", (uid,)).fetchone()
    if not r or not r["NhacChuong"]:
        return _ok(None, "Chưa chọn nhạc chuông.")
    dd = _tuyet_doi(r["NhacChuong"])
    ton_tai = os.path.isfile(dd)
    return _ok({"duong_dan": dd, "ten_file": os.path.basename(dd), "ton_tai": ton_tai},
               "" if ton_tai else "Tệp nhạc chuông đã bị mất, vui lòng chọn lại.")

