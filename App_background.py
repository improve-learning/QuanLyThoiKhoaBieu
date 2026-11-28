#Giao dien da ket noi voi database
import ctypes
import os
import shutil
import subprocess
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from datetime import datetime, timedelta
import PY_backend as db
from PIL import Image, ImageTk

THU = ["Thứ Hai", "Thứ Ba", "Thứ Tư", "Thứ Năm", "Thứ Sáu", "Thứ Bảy", "Chủ Nhật"]
MAU_CHINH = "#273c75"
NEN_TONG = "#FFF3C4"
NEN_KHUNG = "#FFFBE8"
NEN_TIEU_DE = "#F6E7A8"
CHU = "#2f3640"
FONT = "Segoe UI"


def duong_dan_tai_nguyen(ten_file):
    """Lấy đường dẫn tài nguyên khi chạy bằng Python hoặc bản đóng gói PyInstaller."""
    thu_muc = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(thu_muc, ten_file)


ANH_NEN = None
for _ten_anh in ("schooltimebackground.png", "schooltimebackroun.png"):
    try:
        ANH_NEN = Image.open(duong_dan_tai_nguyen(_ten_anh))
        break
    except (FileNotFoundError, OSError):
        pass

def bao_ket_qua(kq, thanh_cong=None):
    if not kq.get("ok"):
        messagebox.showerror("Có lỗi", kq.get("message") or "Thao tác thất bại.")
        return False
    if thanh_cong:
        messagebox.showinfo("Thông báo", thanh_cong)
    return True


# PHÁT NHẠC CHUÔNG MP3 (không cần thư viện ngoài)
_MCI_ALIAS = "nhac_nhac_hen"
_tien_trinh_nhac = None


def _mci(lenh):
    try:
        return ctypes.windll.winmm.mciSendStringW(lenh, None, 0, 0)
    except Exception:
        return -1


def dung_nhac_chuong():
    global _tien_trinh_nhac
    if os.name == "nt":
        _mci(f"stop {_MCI_ALIAS}")
        _mci(f"close {_MCI_ALIAS}")
    if _tien_trinh_nhac is not None:
        try:
            _tien_trinh_nhac.terminate()
        except Exception:
            pass
        _tien_trinh_nhac = None


def phat_nhac_chuong(duong_dan):
    """Bắt đầu phát file mp3. Trả về True nếu phát được, False nếu máy không hỗ trợ."""
    global _tien_trinh_nhac
    dung_nhac_chuong()
    try:
        if os.name == "nt":
            return _mci(f'open "{duong_dan}" type mpegvideo alias {_MCI_ALIAS}') == 0 \
                and _mci(f"play {_MCI_ALIAS}") == 0
        if sys.platform == "darwin":
            lenh = ["afplay", duong_dan]
        else:
            lenh = None
            for ten, tham_so in (("mpg123", ["-q"]), ("ffplay", ["-nodisp", "-autoexit", "-loglevel", "quiet"]),
                                 ("mpv", ["--no-video", "--really-quiet"]), ("cvlc", ["--play-and-exit", "--quiet"])):
                if shutil.which(ten):
                    lenh = [ten, *tham_so, duong_dan]
                    break
            if lenh is None:
                return False
        _tien_trinh_nhac = subprocess.Popen(lenh, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return True
    except Exception:
        return False

class UngDung(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Quản Lý Thời Khóa Biểu")
        self.geometry("1100x700")
        self.minsize(900, 560)
        self.configure(bg=NEN_TONG)
        self.style = ttk.Style(self)
        self.style.theme_use("clam")
        self.style.configure("TNotebook", background=NEN_TONG, borderwidth=0)
        self.style.configure("TNotebook.Tab", font=(FONT, 11, "bold"), padding=(18, 8), background=NEN_TIEU_DE)
        self.style.map("TNotebook.Tab", background=[("selected", MAU_CHINH)], foreground=[("selected", "white")])
        self._after_nhac = None
        self._moc_nhac = datetime.now()
        self._chay_nen = False       # True khi người dùng bấm X và app đang thu nhỏ để chờ nhắc hẹn
        self._da_bao_nen = False
        self._popup_nhac = 0         # số cửa sổ nhắc hẹn đang mở
        self.protocol("WM_DELETE_WINDOW", self._dong_cua_so)
        self._khung_dang_nhap()

    def _can_giua(self, win, theo_man_hinh=False):
        """Đặt cửa sổ con vào chính giữa cửa sổ chính (hoặc giữa màn hình) -> vị trí luôn cố định."""
        win.update_idletasks()
        w, h = win.winfo_reqwidth(), win.winfo_reqheight()
        if theo_man_hinh:
            x, y = (win.winfo_screenwidth() - w) // 2, (win.winfo_screenheight() - h) // 2
        else:
            x = self.winfo_rootx() + (self.winfo_width() - w) // 2
            y = self.winfo_rooty() + (self.winfo_height() - h) // 2
        win.geometry(f"+{max(x, 0)}+{max(y, 0)}")

    def _xoa_man_hinh(self):
        fid = getattr(self, "_bind_nen", None)
        if fid:      # gỡ <Configure> của ảnh nền cũ, tránh dồn callback trỏ vào widget đã hủy
            self.unbind("<Configure>", fid)
            self._bind_nen = None
        for w in self.winfo_children():
            w.destroy()
        self._lbl_anh_nen = None

    def _dat_anh_nen(self, parent):
        """Đặt ảnh nền co giãn theo kích thước cửa sổ/khung."""
        if ANH_NEN is None:
            return

        lbl = tk.Label(parent, bd=0, highlightthickness=0)
        lbl.place(x=0, y=0, relwidth=1, relheight=1)
        lbl.lower()
        cho = [None]    # lần vẽ lại đang chờ (gộp nhiều sự kiện <Configure> liên tiếp thành một lần vẽ)

        def ve():
            cho[0] = None
            if not lbl.winfo_exists():
                return
            rong, cao = parent.winfo_width(), parent.winfo_height()
            if rong <= 1 or cao <= 1:
                return
            anh = ANH_NEN.resize((rong, cao), Image.Resampling.LANCZOS)
            anh_tk = ImageTk.PhotoImage(anh)
            lbl.configure(image=anh_tk)
            lbl.image = anh_tk  # Giữ tham chiếu để Tkinter không làm mất ảnh

        def cap_nhat(event=None):
            if event is not None and event.widget is not parent:
                return
            if cho[0] is not None:
                parent.after_cancel(cho[0])
            cho[0] = parent.after(60, ve)

        fid = parent.bind("<Configure>", cap_nhat, add="+")
        ve()
        if parent is self:
            self._lbl_anh_nen = lbl
            self._bind_nen = fid

    def _khung_dang_nhap(self):
        self._xoa_man_hinh()
        self._dat_anh_nen(self)
        khung = tk.Frame(self, bg=NEN_KHUNG, padx=35, pady=28, bd=1, relief="solid")
        khung.place(relx=.5, rely=.5, anchor="center")
        tk.Label(khung, text="ĐĂNG NHẬP", font=(FONT, 18, "bold"), bg=NEN_KHUNG, fg=MAU_CHINH).pack(pady=(0, 18))
        tk.Label(khung, text="Email", bg=NEN_KHUNG, font=(FONT, 10)).pack(anchor="w")
        self.in_email = ttk.Entry(khung, width=38)
        self.in_email.pack(pady=(3, 12), ipady=4)
        tk.Label(khung, text="Mật khẩu", bg=NEN_KHUNG, font=(FONT, 10)).pack(anchor="w")
        self.in_mk = ttk.Entry(khung, width=38, show="*")
        self.in_mk.pack(pady=(3, 15), ipady=4)
        ttk.Button(khung, text="Đăng nhập", command=self.dang_nhap).pack(fill="x", pady=3)
        ttk.Button(khung, text="Tạo tài khoản mới", command=self.dang_ky).pack(fill="x", pady=3)
        self.in_mk.bind("<Return>", lambda e: self.dang_nhap())

    def dang_nhap(self):
        kq = db.dang_nhap(self.in_email.get().strip(), self.in_mk.get())
        if bao_ket_qua(kq):
            self.nguoi_dung = kq["data"]
            self._man_hinh_chinh()

    def dang_ky(self):
        win = tk.Toplevel(self)
        win.title("Tạo tài khoản")
        win.configure(bg=NEN_KHUNG)
        win.resizable(False, False)
        fields = []
        for label, show in [("Họ tên", ""), ("Email", ""), ("Mật khẩu (ít nhất 8 ký tự, gồm chữ và số)", "*" )]:
            tk.Label(win, text=label, bg=NEN_KHUNG, font=(FONT, 10)).pack(anchor="w", padx=16, pady=(9, 2))
            ent = ttk.Entry(win, width=40, show=show)
            ent.pack(padx=16, pady=(0, 4), ipady=3)
            fields.append(ent)
        def tao():
            kq = db.dang_ky(fields[0].get(), fields[1].get(), fields[2].get())
            if bao_ket_qua(kq):
                self.in_email.delete(0, "end")
                self.in_email.insert(0, fields[1].get().strip())
                self.in_mk.delete(0, "end")
                self.in_mk.insert(0, fields[2].get())
                win.destroy()
        ttk.Button(win, text="Đăng ký", command=tao).pack(fill="x", padx=16, pady=14)
        win.transient(self)
        self._can_giua(win)
        fields[0].focus_set()

    def _man_hinh_chinh(self):
        self._xoa_man_hinh()
        self._dat_anh_nen(self)
        top = tk.Frame(self, bg=MAU_CHINH, padx=14, pady=10)
        top.pack(fill="x")
        tk.Label(top, text="QUẢN LÝ THỜI KHÓA BIỂU", bg=MAU_CHINH, fg="white", font=(FONT, 15, "bold")).pack(side="left")
        tk.Label(top, text=f"Xin chào, {self.nguoi_dung['ho_ten']}", bg=MAU_CHINH, fg="white", font=(FONT, 10)).pack(side="right")
        self.tabs = ttk.Notebook(self)
        self.tabs.pack(fill="both", expand=True, padx=8, pady=8)
        self.tab_tkb = tk.Frame(self.tabs, bg=NEN_TONG)
        self.tab_nhac = tk.Frame(self.tabs, bg=NEN_TONG)
        self.tab_kho = tk.Frame(self.tabs, bg=NEN_TONG)
        self.tab_ca_nhan = tk.Frame(self.tabs, bg=NEN_TONG)
        for frame in (self.tab_tkb, self.tab_nhac, self.tab_kho, self.tab_ca_nhan):
            self._dat_anh_nen(frame)
        for frame, title in [(self.tab_tkb, "Thời khóa biểu"), (self.tab_nhac, "Nhắc hẹn"), (self.tab_kho, "Kho lưu trữ"), (self.tab_ca_nhan, "Cá nhân")]:
            self.tabs.add(frame, text=title)
        self._tao_tab_tkb()
        self._tao_tab_nhac()
        self._tao_tab_kho()
        self._tao_tab_ca_nhan()
        self._tai_du_lieu_ban_dau()
        self._bat_dau_nhac()

    def _tao_bang(self, parent, columns, headings, widths=None):
        outer = tk.Frame(parent, bg=NEN_KHUNG, bd=1, relief="solid")
        outer.pack(fill="both", expand=True, padx=12, pady=12)
        tree = ttk.Treeview(outer, columns=columns, show="headings", selectmode="browse")
        for i, (col, head) in enumerate(zip(columns, headings)):
            tree.heading(col, text=head)
            tree.column(col, width=(widths[i] if widths else 130), anchor="w")
        scroll = ttk.Scrollbar(outer, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        return tree

    def _tao_tab_tkb(self):
        bar = tk.Frame(self.tab_tkb, bg=NEN_TONG)
        bar.pack(fill="x", padx=12, pady=(12, 0))
        ttk.Button(bar, text="+ Thêm thời khóa biểu", command=self.them_buoi_hoc).pack(side="left", padx=(0, 6))
        ttk.Button(bar, text="Sửa thời khóa biểu", command=self.sua_buoi_hoc).pack(side="left", padx=6)
        ttk.Button(bar, text="Xóa thời khóa biểu", command=self.xoa_buoi_hoc).pack(side="left", padx=6)
        self._ds_tkb = {}       # iid (mã buổi học dạng chuỗi) -> bản ghi đầy đủ; dùng để điền sẵn khi sửa
        # Thứ tự cột: Ngày học + sáng/chiều | Tên môn học | Tên giảng viên | Số tiết | Mã phòng
        self.bang_tkb = self._tao_bang(self.tab_tkb, ("ngay", "mon", "gv", "tiet", "phong"),
            ("Ngày học", "Tên môn học", "Tên giảng viên", "Số tiết", "Mã phòng"), (240, 260, 180, 80, 110))
        self.bang_tkb.bind("<Double-1>", lambda e: self.sua_buoi_hoc() if self.bang_tkb.identify_region(e.x, e.y) == "cell" else None)
        self.lbl_hom_nay = tk.Label(self.tab_tkb, bg=NEN_TONG, fg=MAU_CHINH, font=(FONT, 10, "bold"), anchor="w")
        self.lbl_hom_nay.pack(fill="x", padx=14, pady=(0, 12))

    @staticmethod
    def _ngay_buoi(r):
        thu = THU[r["thu"]] if r.get("thu") is not None else ""
        return f'{thu}, {r["ngay_hien_thi"]} - {r["buoi"]}' if thu else f'{r["ngay_hien_thi"]} - {r["buoi"]}'

    def tai_tkb(self, chon=None):
        self.bang_tkb.delete(*self.bang_tkb.get_children())
        self._ds_tkb = {}
        kq = db.lay_tkb()
        if not kq.get("ok"):
            bao_ket_qua(kq); return
        hom_nay = datetime.now().strftime("%Y-%m-%d")
        ds_hom_nay = []
        for r in kq["data"]:
            iid = str(r["id"])
            self._ds_tkb[iid] = r
            self.bang_tkb.insert("", "end", iid=iid, values=(self._ngay_buoi(r), r["mon_hoc"], r["giang_vien"], r["so_tiet"], r["phong"]))
            if r["ngay"] == hom_nay:
                ds_hom_nay.append(f'{r["buoi"]}: {r["mon_hoc"]} ({r["giang_vien"]}, {r["so_tiet"]} tiết, phòng {r["phong"] or "chưa có"})')
        self.lbl_hom_nay.config(text="Lịch hôm nay: " + (";  ".join(ds_hom_nay) if ds_hom_nay else "Chưa có buổi học."))
        if chon is not None and str(chon) in self._ds_tkb:       # chọn lại dòng vừa thêm/sửa
            self.bang_tkb.selection_set(str(chon)); self.bang_tkb.see(str(chon))

    def _hop_thoai_buoi_hoc(self, tieu_de, ten_nut, mac_dinh, luu):
        """Cửa sổ dùng chung cho Thêm và Sửa. mac_dinh = (môn, giảng viên, ngày, buổi, số tiết, phòng).
        luu(mon, gv, ngay, buoi, so_tiet, phong) -> kết quả backend."""
        win = tk.Toplevel(self); win.title(tieu_de); win.configure(bg=NEN_KHUNG); win.resizable(False, False)
        mon, gv, ngay, buoi, tiet, phong = mac_dinh
        o = {}
        for khoa, label, gia_tri in (("mon", "Tên môn học", mon), ("gv", "Tên giảng viên", gv), ("ngay", "Ngày học (DD/MM/YYYY)", ngay), ("phong", "Phòng học", phong)):
            tk.Label(win, text=label, bg=NEN_KHUNG, font=(FONT, 10)).pack(anchor="w", padx=14, pady=(8, 1))
            e = ttk.Entry(win, width=40); e.pack(padx=14, ipady=3); e.insert(0, gia_tri); o[khoa] = e
        hang = tk.Frame(win, bg=NEN_KHUNG); hang.pack(fill="x", padx=14, pady=(8, 0))
        cot_b = tk.Frame(hang, bg=NEN_KHUNG); cot_b.pack(side="left", fill="x", expand=True, padx=(0, 8))
        tk.Label(cot_b, text="Buổi", bg=NEN_KHUNG, font=(FONT, 10)).pack(anchor="w", pady=(0, 1))
        o["buoi"] = ttk.Combobox(cot_b, values=("Sáng", "Chiều"), state="readonly", width=16); o["buoi"].pack(fill="x", ipady=2)
        o["buoi"].set(buoi)
        cot_t = tk.Frame(hang, bg=NEN_KHUNG); cot_t.pack(side="left", fill="x", expand=True)
        tk.Label(cot_t, text="Số tiết", bg=NEN_KHUNG, font=(FONT, 10)).pack(anchor="w", pady=(0, 1))
        o["tiet"] = ttk.Combobox(cot_t, values=tuple(str(i) for i in range(1, db.TIET_TOI_DA + 1)), state="readonly", width=16); o["tiet"].pack(fill="x", ipady=2)
        o["tiet"].set(str(tiet))
        def save(_e=None):
            kq = luu(o["mon"].get().strip(), o["gv"].get().strip(), o["ngay"].get().strip(), o["buoi"].get(), o["tiet"].get(), o["phong"].get().strip())
            if bao_ket_qua(kq):
                win.destroy(); self.tai_tkb(chon=(kq.get("data") or {}).get("id"))
        for k in ("mon", "gv", "ngay", "phong"):
            o[k].bind("<Return>", save)
        ttk.Button(win, text=ten_nut, command=save).pack(fill="x", padx=14, pady=14)
        win.transient(self)
        self._can_giua(win)
        o["mon"].focus_set()

    def them_buoi_hoc(self):
        self._hop_thoai_buoi_hoc("Thêm thời khóa biểu", "Lưu buổi học",
                                 ("", "", datetime.now().strftime("%d/%m/%Y"), "Sáng", 3, ""), db.them_buoi_hoc)

    def sua_buoi_hoc(self):
        sel = self.bang_tkb.selection()
        if not sel:
            messagebox.showwarning("Chú ý", "Chọn một buổi học để sửa"); return
        r = self._ds_tkb.get(sel[0])
        if r is None: return
        ma = int(sel[0])
        self._hop_thoai_buoi_hoc("Sửa thời khóa biểu", "Lưu thay đổi",
                                 (r["mon_hoc"], r["giang_vien"], r["ngay_hien_thi"], r["buoi"], r["so_tiet"], r["phong"]),
                                 lambda mon, gv, ngay, buoi, tiet, phong: db.sua_buoi_hoc(ma, mon, gv, ngay, buoi, tiet, phong))

    def xoa_buoi_hoc(self):
        sel = self.bang_tkb.selection()
        if not sel:
            messagebox.showwarning("Chú ý", "Chọn một buổi học để xóa"); return
        if messagebox.askyesno("Xác nhận", "Xóa buổi học đã chọn?"):
            if bao_ket_qua(db.xoa_buoi_hoc(int(sel[0]))): self.tai_tkb()

    def _tao_tab_nhac(self):
        bar = tk.Frame(self.tab_nhac, bg=NEN_TONG); bar.pack(fill="x", padx=12, pady=(12, 0))
        ttk.Button(bar, text="+ Thêm nhắc hẹn", command=self.them_nhac_hen).pack(side="left", padx=(0, 6))
        ttk.Button(bar, text="Sửa nhắc hẹn đã chọn", command=self.sua_nhac_hen).pack(side="left", padx=6)
        ttk.Button(bar, text="Xóa nhắc hẹn đã chọn", command=self.xoa_nhac_hen).pack(side="left", padx=6)
        self._ds_nhac = {}      # iid (mã nhắc hẹn dạng chuỗi) -> bản ghi đầy đủ; dùng để điền sẵn khi sửa
        self.bang_nhac = self._tao_bang(self.tab_nhac, ("id", "ngay", "gio", "noidung"), ("Mã", "Ngày", "Giờ", "Nội dung"), (60, 140, 100, 550))
        self.bang_nhac.bind("<Double-1>", lambda e: self.sua_nhac_hen() if self.bang_nhac.identify_region(e.x, e.y) == "cell" else None)

    def _bat_dau_nhac(self):
        self._dung_nhac()
        self._chay_nen = False
        self._moc_nhac = datetime.now()
        self._after_nhac = self.after(1000, self._kiem_tra_nhac)

    def _dung_nhac(self):
        if self._after_nhac is not None:
            try:
                self.after_cancel(self._after_nhac)
            except Exception:
                pass
            self._after_nhac = None

    def _co_nhac_sap_toi(self):
        """Còn nhắc hẹn nào chưa tới giờ không?"""
        kq = db.lay_nhac_hen()
        if not kq.get("ok"):
            return False
        bay_gio = datetime.now().strftime("%Y-%m-%d %H:%M")
        return any(f'{r["ngay"]} {r["gio"]}' > bay_gio for r in kq["data"])

    def _thoat(self):
        self._dung_nhac()
        dung_nhac_chuong()
        db.dang_xuat()
        self.destroy()

    def _dong_cua_so(self):
        """Bấm X: còn nhắc hẹn sắp tới -> chạy nền (thu nhỏ); không còn / chưa đăng nhập -> thoát hẳn."""
        if db.da_dang_nhap() and self._co_nhac_sap_toi():
            if not self._da_bao_nen:
                messagebox.showinfo("Chạy nền", "Ứng dụng sẽ thu nhỏ xuống thanh tác vụ và tiếp tục nhắc hẹn.\n"
                                    "Để tắt hẳn: mở lại ứng dụng rồi bấm Đăng xuất.\n"
                                    "Ứng dụng cũng tự tắt sau khi nhắc hẹn cuối cùng được báo.", parent=self)
                self._da_bao_nen = True
            self._chay_nen = True
            self.iconify()
        else:
            self._thoat()

    def _tu_thoat_neu_het_nhac(self):
        """Đang chạy nền mà hết nhắc hẹn (và không còn cửa sổ nhắc nào mở) -> tắt app."""
        if self._chay_nen and self._popup_nhac == 0 and not self._co_nhac_sap_toi():
            self._thoat()

    def _kiem_tra_nhac(self):
        self._after_nhac = self.after(1000, self._kiem_tra_nhac)
        tu, bay_gio = self._moc_nhac, datetime.now()
        self._moc_nhac = bay_gio
        if not db.da_dang_nhap():
            return
        if self._chay_nen and self.state() != "iconic":   # người dùng đã mở lại app -> hết chế độ chạy nền
            self._chay_nen = False
        dang = "%Y-%m-%d %H:%M"
        if tu.strftime(dang) == bay_gio.strftime(dang):   # chưa sang phút mới -> chưa có gì đến hạn
            return
        kq = db.lay_nhac_hen_den_han(tu, bay_gio)
        if kq.get("ok"):
            for r in kq["data"]:
                self._hien_thong_bao_nhac(r)
        if self._chay_nen:
            self._tu_thoat_neu_het_nhac()

    def _hien_thong_bao_nhac(self, r):
        win = tk.Toplevel(self); win.title("Nhắc hẹn"); win.configure(bg=NEN_KHUNG); win.resizable(False, False)
        tk.Label(win, text="NHẮC HẸN", font=(FONT, 15, "bold"), bg=NEN_KHUNG, fg=MAU_CHINH).pack(padx=34, pady=(18, 4))
        tk.Label(win, text=f'{r["gio"]} - {r["ngay_hien_thi"]}', font=(FONT, 10, "bold"), bg=NEN_KHUNG, fg=CHU).pack()
        tk.Label(win, text=r["noi_dung"], font=(FONT, 11), bg=NEN_KHUNG, fg=CHU, wraplength=360, justify="left").pack(padx=34, pady=12)
        def dong():
            self._popup_nhac = max(0, self._popup_nhac - 1)
            dung_nhac_chuong()
            win.destroy()
            self._tu_thoat_neu_het_nhac()
        ttk.Button(win, text="Đã biết", command=dong).pack(fill="x", padx=34, pady=(2, 18))
        win.protocol("WM_DELETE_WINDOW", dong)
        self._popup_nhac += 1
        self._can_giua(win, theo_man_hinh=True)
        win.deiconify()
        win.attributes("-topmost", True)
        win.lift(); win.focus_force()
        if not self._phat_nhac():
            self.bell()

    def _phat_nhac(self):
        """Phát nhạc chuông đã chọn. Trả về False nếu không phát được (khi đó dùng tiếng bíp)."""
        try:
            kq = db.lay_nhac_chuong()
            if not (kq.get("ok") and kq.get("data") and kq["data"].get("ton_tai")):
                return False
            return phat_nhac_chuong(kq["data"]["duong_dan"])
        except Exception:
            return False

    def tai_nhac_hen(self, chon=None):
        self.bang_nhac.delete(*self.bang_nhac.get_children())
        self._ds_nhac = {}
        kq = db.lay_nhac_hen()
        if not kq.get("ok"):
            bao_ket_qua(kq); return
        for r in kq["data"]:
            iid = str(r["id"])
            self._ds_nhac[iid] = r
            self.bang_nhac.insert("", "end", iid=iid, values=(r["id"], r["ngay_hien_thi"], r["gio"], r["noi_dung"]))
        if chon is not None and str(chon) in self._ds_nhac:      # chọn lại dòng vừa thêm/sửa
            self.bang_nhac.selection_set(str(chon)); self.bang_nhac.see(str(chon))

    def _hop_thoai_nhac_hen(self, tieu_de, ten_nut, mac_dinh, luu):
        """Hộp thoại dùng chung cho Thêm và Sửa nhắc hẹn. luu(ngay, gio, noi_dung) -> kết quả backend."""
        win = tk.Toplevel(self); win.title(tieu_de); win.configure(bg=NEN_KHUNG); win.resizable(False, False)
        entries = []
        for label, gia_tri in zip(("Ngày (DD/MM/YYYY)", "Giờ (HH:MM)", "Nội dung"), mac_dinh):
            tk.Label(win, text=label, bg=NEN_KHUNG, font=(FONT, 10)).pack(anchor="w", padx=14, pady=(8, 1))
            e = ttk.Entry(win, width=42); e.pack(padx=14, ipady=3); e.insert(0, gia_tri); entries.append(e)
        def save(_e=None):
            kq = luu(*[e.get().strip() for e in entries])
            if bao_ket_qua(kq):
                win.destroy(); self.tai_nhac_hen(chon=(kq.get("data") or {}).get("id"))
        for e in entries:
            e.bind("<Return>", save)
        ttk.Button(win, text=ten_nut, command=save).pack(fill="x", padx=14, pady=14)
        win.transient(self)
        self._can_giua(win)
        entries[-1].focus_set()

    def them_nhac_hen(self):
        bay_gio = datetime.now() + timedelta(minutes=1)
        self._hop_thoai_nhac_hen("Thêm nhắc hẹn", "Lưu nhắc hẹn",
                                 (bay_gio.strftime("%d/%m/%Y"), bay_gio.strftime("%H:%M"), ""), db.them_nhac_hen)

    def sua_nhac_hen(self):
        sel = self.bang_nhac.selection()
        if not sel: messagebox.showwarning("Chú ý", "Chọn một nhắc hẹn."); return
        ma = int(sel[0])
        r = self._ds_nhac.get(sel[0])
        if r is None: return
        self._hop_thoai_nhac_hen("Sửa nhắc hẹn", "Lưu thay đổi",
                                 (r["ngay_hien_thi"], r["gio"], r["noi_dung"]),
                                 lambda ngay, gio, nd: db.sua_nhac_hen(ma, ngay, gio, nd))

    def xoa_nhac_hen(self):
        sel = self.bang_nhac.selection()
        if not sel: messagebox.showwarning("Chú ý", "Chọn một nhắc hẹn"); return
        if messagebox.askyesno("Xác nhận", "Xóa nhắc hẹn đã chọn?"):
            if bao_ket_qua(db.xoa_nhac_hen(int(sel[0]))): self.tai_nhac_hen()

    def _tao_tab_kho(self):
        bar = tk.Frame(self.tab_kho, bg=NEN_TONG); bar.pack(fill="x", padx=12, pady=(12, 0))
        ttk.Button(bar, text="+ Tải file Excel (.xlsx)", command=self.tai_file).pack(side="left", padx=(0, 6))
        ttk.Button(bar, text="Mở file đã chọn", command=self.mo_file).pack(side="left", padx=6)
        self._ds_kho = {}       # iid (mã file dạng chuỗi) -> bản ghi đầy đủ
        self.bang_kho = self._tao_bang(self.tab_kho, ("id", "ten", "ngay", "ton_tai", "duong_dan"), ("Mã", "Tên file gốc", "Thời gian tải", "Tồn tại", "Đường dẫn lưu"), (55, 220, 160, 80, 420))

    def tai_kho(self, canh_bao=False, chon=None):
        self.bang_kho.delete(*self.bang_kho.get_children())
        self._ds_kho = {}
        kq = db.lay_danh_sach_file()
        if not kq.get("ok"):
            bao_ket_qua(kq); return
        for r in kq["data"]:
            iid = str(r["id"])
            self._ds_kho[iid] = r
            self.bang_kho.insert("", "end", iid=iid, values=(r["id"], r["ten_file"], r["thoi_gian_tai"], "Có" if r["ton_tai"] else "Mất", r["duong_dan"]))
        if chon is not None and str(chon) in self._ds_kho:
            self.bang_kho.selection_set(str(chon)); self.bang_kho.see(str(chon))
        if canh_bao and kq.get("message"):          # chỉ cảnh báo file bị mất khi mở app, không hiện lại ở mỗi lần tải danh sách
            messagebox.showwarning("Kho lưu trữ", kq["message"])

    def tai_file(self):
        """Chọn một hoặc nhiều file .xlsx cùng lúc; báo kết quả gộp trong một hộp thoại."""
        chon = filedialog.askopenfilenames(title="Chọn file Excel", filetypes=[("Excel .xlsx", "*.xlsx")])
        paths = self.tk.splitlist(chon) if chon else ()
        if not paths: return
        so_ok, loi, id_cuoi = 0, [], None
        self.config(cursor="watch"); self.update_idletasks()
        try:
            for p in paths:
                kq = db.tai_file_len(p)
                if kq.get("ok"):
                    so_ok += 1; id_cuoi = kq["data"]["id"]
                else:
                    loi.append(f"{os.path.basename(p)}: {kq.get('message') or 'Thao tác thất bại.'}")
        finally:
            self.config(cursor="")
        self.tai_kho(chon=id_cuoi)
        if len(paths) == 1:
            if so_ok: messagebox.showinfo("Thông báo", "Đã tải tệp lên kho lưu trữ.")
            else: messagebox.showerror("Có lỗi", loi[0])
            return
        msg = f"Đã tải lên {so_ok}/{len(paths)} tệp."
        if loi:
            msg += "\n\nKhông tải được:\n" + "\n".join(loi[:8]) + ("\n..." if len(loi) > 8 else "")
        (messagebox.showinfo if not loi else messagebox.showwarning)("Kết quả tải lên", msg)

    def mo_file(self):
        sel = self.bang_kho.selection()
        if not sel: messagebox.showwarning("Chú ý", "Chọn một file."); return
        path = self._ds_kho[sel[0]]["duong_dan"]
        if not os.path.isfile(path): messagebox.showerror("Lỗi", "File không còn tồn tại."); return
        try:
            if os.name == "nt": os.startfile(path)
            elif sys.platform == "darwin": subprocess.Popen(["open", path])
            else: subprocess.Popen(["xdg-open", path])
        except Exception as e: messagebox.showerror("Lỗi", str(e))

    def _tao_tab_ca_nhan(self):
        khung = tk.Frame(self.tab_ca_nhan, bg=NEN_KHUNG, padx=22, pady=22, bd=1, relief="solid")
        khung.pack(fill="x", padx=12, pady=12)
        self.lbl_ten = tk.Label(khung, text=self.nguoi_dung["ho_ten"], font=(FONT, 16, "bold"), bg=NEN_KHUNG, fg=MAU_CHINH)
        self.lbl_ten.pack(anchor="w", pady=(0, 12))
        self.lbl_email = tk.Label(khung, text=self.nguoi_dung["email"], font=(FONT, 11), bg=NEN_KHUNG, fg=CHU)
        self.lbl_email.pack(anchor="w", pady=(0, 14))
        ttk.Button(khung, text="Cập nhật thông tin cá nhân", command=self.cap_nhat_thong_tin).pack(fill="x", pady=4)
        ttk.Button(khung, text="Đổi mật khẩu", command=self.doi_mat_khau).pack(fill="x", pady=4)
        ttk.Button(khung, text="Chọn nhạc chuông MP3", command=self.chon_nhac_chuong).pack(fill="x", pady=4)
        ttk.Button(khung, text="Đăng xuất", command=self.dang_xuat).pack(fill="x", pady=4)

    def cap_nhat_thong_tin(self):
        ten = simpledialog.askstring("Thông tin cá nhân", "Họ tên mới:", initialvalue=self.nguoi_dung["ho_ten"], parent=self)
        if ten is None: return
        email = simpledialog.askstring("Thông tin cá nhân", "Email mới:", initialvalue=self.nguoi_dung["email"], parent=self)
        if email is None: return
        kq = db.cap_nhat_thong_tin(ten, email)
        if bao_ket_qua(kq):
            self.nguoi_dung.update(kq["data"]); self.lbl_ten.config(text=ten.strip()); self.lbl_email.config(text=email.strip())

    def doi_mat_khau(self):
        # Bước 1: bắt buộc nhập ĐÚNG mật khẩu hiện tại, sai thì báo lỗi và cho nhập lại
        while True:
            cu = simpledialog.askstring("Đổi mật khẩu", "Mật khẩu hiện tại:", show="*", parent=self)
            if cu is None: return
            kq = db.dang_nhap(self.nguoi_dung["email"], cu)
            if kq.get("ok") and kq["data"]["id"] == self.nguoi_dung["id"]:
                break
            messagebox.showerror("Có lỗi", "Mật khẩu hiện tại không đúng. Vui lòng nhập lại.", parent=self)
        # Bước 2: nhập mật khẩu mới (không đạt yêu cầu thì báo lỗi và cho nhập lại)
        while True:
            moi = simpledialog.askstring("Đổi mật khẩu", "Mật khẩu mới (ít nhất 8 ký tự, gồm chữ và số):", show="*", parent=self)
            if moi is None: return
            kq = db.doi_mat_khau(cu, moi)
            if kq.get("ok"):
                messagebox.showinfo("Thông báo", "Đổi mật khẩu thành công.", parent=self)
                return
            messagebox.showerror("Có lỗi", kq.get("message") or "Thao tác thất bại.", parent=self)

    def chon_nhac_chuong(self):
        path = filedialog.askopenfilename(title="Chọn nhạc chuông", filetypes=[("MP3", "*.mp3")])
        if not path: return
        kq = db.chon_nhac_chuong(path)
        bao_ket_qua(kq, "Đã lưu nhạc chuông.")

    def dang_xuat(self):
        if messagebox.askyesno("Đăng xuất", "Đăng xuất?"):
            self._dung_nhac()
            self._chay_nen = False
            dung_nhac_chuong()
            db.dang_xuat()
            self._khung_dang_nhap()

    def _tai_du_lieu_ban_dau(self):
        self.tai_tkb(); self.tai_nhac_hen(); self.tai_kho(canh_bao=True)
        self.tabs.bind("<<NotebookTabChanged>>", self._khi_doi_tab)   # thay cho nút "Làm mới": tự tải lại tab vừa mở

    def _khi_doi_tab(self, _e=None):
        if not db.da_dang_nhap():
            return
        try:
            tab = self.tabs.select()
            if tab == str(self.tab_tkb): self.tai_tkb()
            elif tab == str(self.tab_nhac): self.tai_nhac_hen()
            elif tab == str(self.tab_kho): self.tai_kho()
        except tk.TclError:
            pass


if __name__ == "__main__":
    db.khoi_tao()
    UngDung().mainloop()
