

from __future__ import annotations

import os
import sys
import json
import math
import logging
import traceback
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

import requests
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# ------------------------
# تنظیمات لاگ
# ------------------------
logging.basicConfig(level=logging.INFO, format='[%(asctime)s] %(message)s', datefmt='%Y-%m-%d %H:%M:%S')

# ------------------------
# فایل تنظیمات محلی
# ------------------------
SETTINGS_FILE = "client_type_export_settings.json"

def load_settings() -> Dict[str, Any]:
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except json.JSONDecodeError:
            try:
                bak = SETTINGS_FILE + ".corrupt"
                os.replace(SETTINGS_FILE, bak)
                logging.warning("فایل تنظیمات خراب بود؛ به %s منتقل شد.", bak)
            except Exception:
                logging.exception("خطا هنگام جابجایی فایل تنظیمات خراب.")
            return {}
        except Exception:
            logging.exception("خطا هنگام بارگذاری فایل تنظیمات.")
            return {}
    return {}

def save_settings(d: Dict[str, Any]) -> None:
    try:
        with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
            json.dump(d, f, ensure_ascii=False, indent=2)
    except Exception:
        logging.exception("خطا هنگام ذخیره تنظیمات.")

settings_store: Dict[str, Any] = load_settings()

# مقادیر پیش‌فرض
DEFAULTS = {
    "client_url_template": "https://cdn.tsetmc.com/api/ClientType/GetClientTypeHistory/{inscode}",
    "price_url_template": "https://cdn.tsetmc.com/api/ClosingPrice/GetChartData/{inscode}/D",
    "last_out_dir": ".",
    "dEven_offset_mode": "auto"  # "auto", "ms", "s", "none"
}
for k, v in DEFAULTS.items():
    settings_store.setdefault(k, v)
save_settings(settings_store)

# ------------------------
# تبدیل تاریخ شمسی (با jdatetime اگر موجود باشد، در غیر این صورت تبدیل داخلی)
# ------------------------
try:
    import jdatetime  # type: ignore
    def gregorian_to_jalali_str(dt: Optional[datetime]) -> str:
        if dt is None:
            return ""
        try:
            d = dt.date() if hasattr(dt, "date") else dt
            j = jdatetime.date.fromgregorian(date=d)
            return f"{j.year:04d}{j.month:02d}{j.day:02d}"
        except Exception:
            return ""
except Exception:
    # تبدیل داخلی بدون وابستگی
    def _gregorian_to_jalali(y: int, m: int, d: int) -> Tuple[int, int, int]:
        gy = y - 1600
        gm = m - 1
        gd = d - 1
        g_day_no = 365 * gy + (gy + 3) // 4 - (gy + 99) // 100 + (gy + 399) // 400
        months = [31,28,31,30,31,30,31,31,30,31,30,31]
        for i in range(gm):
            g_day_no += months[i]
        if gm > 1 and ((y % 4 == 0 and y % 100 != 0) or (y % 400 == 0)):
            g_day_no += 1
        g_day_no += gd
        j_day_no = g_day_no - 79
        j_np = j_day_no // 12053
        j_day_no = j_day_no % 12053
        jy = 979 + 33 * j_np + 4 * (j_day_no // 1461)
        j_day_no %= 1461
        if j_day_no >= 366:
            jy += (j_day_no - 366) // 365
            j_day_no = (j_day_no - 366) % 365
        jalali_months = [31,31,31,31,31,31,30,30,30,30,30,29]
        jm = 0; jd = 0
        for i, v in enumerate(jalali_months):
            if j_day_no < v:
                jm = i + 1
                jd = j_day_no + 1
                break
            j_day_no -= v
        return jy, jm, jd

    def gregorian_to_jalali_str(dt: Optional[datetime]) -> str:
        if dt is None:
            return ""
        try:
            d = dt.date() if hasattr(dt, "date") else dt
            jy, jm, jd = _gregorian_to_jalali(d.year, d.month, d.day)
            return f"{jy:04d}{jm:02d}{jd:02d}"
        except Exception:
            return ""

# ------------------------
# توابع کمکی تاریخ
# ------------------------
def parse_recdate_int(rec: Any) -> Optional[datetime]:
    try:
        if rec is None:
            return None
        s = str(int(rec))
        if len(s) != 8:
            return None
        year = int(s[0:4]); month = int(s[4:6]); day = int(s[6:8])
        return datetime(year, month, day)
    except Exception:
        return None

def dEven_to_datetime_heuristic(dEven: Any, prefer_mode: Optional[str] = None) -> Optional[datetime]:
    """
    تبدیل هوشمند dEven به datetime:
    - اگر prefer_mode مشخص باشد ("ms" یا "s" یا "none") آن را در اولویت قرار می‌دهد.
    - در حالت auto: ابتدا ms، سپس s را امتحان می‌کند.
    - منفی بودن مقدار را بدون گرفتن قدرمطلق بررسی می‌کند؛ اگر تبدیل به تاریخ معقول (1970..2100) شد، قبول می‌شود.
    """
    try:
        if dEven is None:
            return None
        val = float(dEven)
        mode = prefer_mode or settings_store.get("dEven_offset_mode", "auto")

        def try_ms(v):
            try:
                dt = datetime.utcfromtimestamp(v / 1000.0)
                if 1970 <= dt.year <= 2100:
                    return dt
            except Exception:
                return None
            return None

        def try_s(v):
            try:
                dt = datetime.utcfromtimestamp(v)
                if 1970 <= dt.year <= 2100:
                    return dt
            except Exception:
                return None
            return None

        if mode == "ms":
            return try_ms(val)
        if mode == "s":
            return try_s(val)
        if mode == "none":
            return None

        # auto
        dt = try_ms(val)
        if dt:
            return dt
        dt = try_s(val)
        if dt:
            return dt
        return None
    except Exception:
        return None

# ------------------------
# فراخوانی HTTP و پارس JSON
# ------------------------
REQUEST_TIMEOUT = 30.0

def fetch_json(url: str, timeout: float = REQUEST_TIMEOUT) -> Tuple[bool, Optional[Any], Optional[str]]:
    try:
        r = requests.get(url, timeout=timeout)
        if r.status_code != 200:
            return False, None, f"HTTP {r.status_code}"
        try:
            j = r.json()
            return True, j, None
        except Exception as e:
            return False, None, f"JSON parse error: {e}"
    except Exception as e:
        return False, None, str(e)

# ------------------------
# فیلدها و نگاشت خروجی
# ------------------------
CLIENT_FIELDS = [
    "recDate", "insCode", "buy_I_Volume", "buy_N_Volume", "buy_I_Value", "buy_N_Value",
    "buy_N_Count", "sell_I_Volume", "buy_I_Count", "sell_N_Volume", "sell_I_Value",
    "sell_N_Value", "sell_N_Count", "sell_I_Count"
]
PRICE_FIELDS = ["dEven", "pDrCotVal", "qTotTran5J", "priceFirst", "priceMin", "priceMax"]

# ------------------------
# توابع کمکی برای نام فایل امن
# ------------------------
def safe_filename(name: str) -> str:
    # حذف کاراکترهای نامعتبر برای نام فایل در ویندوز/لینوکس/مک
    invalid = '<>:"/\\|?*\0'
    out = ''.join(c for c in name if c not in invalid)
    out = out.strip()
    if not out:
        out = "symbol"
    return out

# ------------------------
# ترکیب داده‌ها (نسخهٔ کامل و مقاوم)
# ------------------------
def merge_client_and_price(client_list: List[Dict[str, Any]], price_list: List[Dict[str, Any]], symbol: str) -> pd.DataFrame:
    """
    ترکیب clientType و closingPrice:
    - price_list معمولاً از قدیم->جدید است؛ آن را معکوس می‌کنیم تا جدیدترین اول شود.
    - تلاش می‌کنیم رکورد قیمت متناظر با recDate را با استفاده از dEven پیدا کنیم.
    - نام‌گذاری قیمت‌ها به pf, pl, pmin, pmax, vol تغییر می‌کند.
    - ستون‌های خروجی به ترتیب خواسته‌شده مرتب می‌شوند.
    """
    rows: List[Dict[str, Any]] = []

    price_list_rev = list(reversed(price_list or []))

    # ساخت نقشه قیمت بر اساس تاریخ (YYYYMMDD) اگر dEven قابل تبدیل باشد
    price_by_date: Dict[str, List[Dict[str, Any]]] = {}
    for p in price_list_rev:
        dEven = p.get("dEven")
        dt = dEven_to_datetime_heuristic(dEven)
        if dt is not None:
            key = dt.strftime("%Y%m%d")
            price_by_date.setdefault(key, []).append(p)

    for idx, c in enumerate(client_list or []):
        recDate = c.get("recDate")
        rec_dt = parse_recdate_int(recDate)
        rec_iso = rec_dt.strftime("%Y-%m-%d") if rec_dt else ""
        rec_jalali = gregorian_to_jalali_str(rec_dt) if rec_dt else ""
        rec_key = rec_dt.strftime("%Y%m%d") if rec_dt else ""

        matched_price = None
        if rec_key and rec_key in price_by_date and price_by_date[rec_key]:
            matched_price = price_by_date[rec_key].pop(0)
        else:
            if idx < len(price_list_rev):
                matched_price = price_list_rev[idx]
            else:
                matched_price = None

        row: Dict[str, Any] = {}
        row["ticker"] = symbol

        if matched_price:
            dt_price = dEven_to_datetime_heuristic(matched_price.get("dEven"))
            row["pf"] = matched_price.get("priceFirst", "")
            row["pl"] = matched_price.get("pDrCotVal", "")
            row["pmin"] = matched_price.get("priceMin", "")
            row["pmax"] = matched_price.get("priceMax", "")
            row["vol"] = matched_price.get("qTotTran5J", "")
            if dt_price:
                row["price_date_iso"] = dt_price.strftime("%Y-%m-%d")
                row["price_date_jalali"] = gregorian_to_jalali_str(dt_price)
            else:
                row["price_date_iso"] = rec_iso
                row["price_date_jalali"] = rec_jalali
        else:
            row["pf"] = ""
            row["pl"] = ""
            row["pmin"] = ""
            row["pmax"] = ""
            row["vol"] = ""
            row["price_date_iso"] = rec_iso
            row["price_date_jalali"] = rec_jalali

        row["recDate"] = int(recDate) if recDate is not None else ""
        row["jalalidate"] = rec_jalali

        # اضافه کردن فیلدهای client با همان نام مرجع (به جز recDate و insCode که جداگانه اضافه می‌شوند)
        for f in CLIENT_FIELDS:
            if f in ("recDate", "insCode"):
                continue
            row[f] = c.get(f, "")

        row["insCode"] = c.get("insCode", "")

        rows.append(row)

    df = pd.DataFrame(rows)

    # ترتیب ستون‌ها
    client_order = [f for f in CLIENT_FIELDS if f not in ("recDate", "insCode")]
    final_cols = ["ticker", "pf", "pl", "pmin", "pmax", "vol", "recDate", "jalalidate"] + client_order + ["insCode", "price_date_iso", "price_date_jalali"]
    final_cols = [c for c in final_cols if c in df.columns]

    # مرتب‌سازی بر اساس recDate نزولی (جدیدترین اول)
    if "recDate" in df.columns:
        try:
            df = df.sort_values(by="recDate", ascending=False).reset_index(drop=True)
        except Exception:
            pass

    return df[final_cols]

# ------------------------
# تابع اصلی دانلود و ذخیره CSV
# ------------------------
def fetch_and_save_for_symbol(ins_code: str, symbol: str, out_dir: str = ".", client_url_template: Optional[str] = None, price_url_template: Optional[str] = None) -> Tuple[bool, str]:
    """
    دانلود داده‌های حقیقی/حقوقی و قیمت برای یک ins_code و ذخیرهٔ CSV.
    نام فایل خروجی: <safe_symbol>.csv   (مثال: قیراط.csv)
    اگر فایل با همین نام وجود داشته باشد، بازنویسی می‌شود.
    """
    try:
        logging.info("شروع دانلود برای: %s نماد: %s", ins_code, symbol)
        client_url_template = client_url_template or settings_store.get("client_url_template")
        price_url_template = price_url_template or settings_store.get("price_url_template")

        url_client = client_url_template.format(inscode=ins_code)
        ok_c, json_c, err_c = fetch_json(url_client)
        if not ok_c or json_c is None:
            msg = f"حقیقی/حقوقی: FAILED {err_c}"
            logging.error(msg)
            return False, msg

        client_list: List[Dict[str, Any]] = []
        if isinstance(json_c, dict) and "clientType" in json_c:
            client_list = json_c.get("clientType", [])
        elif isinstance(json_c, list):
            client_list = json_c
        else:
            if isinstance(json_c, dict):
                for v in json_c.values():
                    if isinstance(v, list):
                        client_list = v
                        break

        logging.info("حقیقی/حقوقی: OK %s", json.dumps(client_list[:1], ensure_ascii=False) if client_list else "{}")

        url_price = price_url_template.format(inscode=ins_code)
        ok_p, json_p, err_p = fetch_json(url_price)
        price_list: List[Dict[str, Any]] = []
        if not ok_p or json_p is None:
            logging.warning("قیمت: FAILED %s", err_p)
            price_list = []
        else:
            if isinstance(json_p, dict) and "closingPriceChartData" in json_p:
                price_list = json_p.get("closingPriceChartData", [])
            elif isinstance(json_p, list):
                price_list = json_p
            else:
                price_list = []
            logging.info("قیمت: OK %s", json.dumps(price_list[:1], ensure_ascii=False) if price_list else "{}")

        df = merge_client_and_price(client_list, price_list, symbol)

        # نام فایل خروجی: فقط نماد (safe) + .csv
        safe_sym = safe_filename(symbol)
        filename = f"{safe_sym}.csv"
        out_path = os.path.join(out_dir, filename)

        # ذخیره CSV با utf-8-sig برای سازگاری با Excel فارسی
        df.to_csv(out_path, index=False, encoding="utf-8-sig")
        logging.info("CSV ذخیره شد: %s", out_path)
        return True, out_path
    except Exception as e:
        logging.exception("خطا هنگام دانلود یا ذخیره:")
        return False, str(e)

# ------------------------
# رابط کاربری Tkinter: ClientTypeExportWindow
# ------------------------
class ClientTypeExportWindow(tk.Toplevel):
    """
    پنجرهٔ خروجی حقیقی/حقوقی و قیمت با قابلیت:
    - بارگذاری خودکار لیست نمادهای فیلترشده از current_tree
    - انتخاب/عدم انتخاب نمادها و دانلود گروهی
    - نمایش نوار پیشرفت، وضعیت و ETA
    - امکان لغو عملیات
    """
    def __init__(self, master=None, current_tree=None, selection_iid=None):
        super().__init__(master)
        self.title("خروجی حقیقی/حقوقی و قیمت (گروهی)")
        self.geometry("980x720")
        self.resizable(True, True)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        self.current_tree = current_tree
        self.selection_iid = selection_iid

        # صف دانلود: لیستی از دیکشنری {insCode, symbol}
        self.download_queue: List[Dict[str, str]] = []
        self._is_downloading = False
        self._download_start_time: Optional[datetime] = None
        self._processed_count = 0
        self._cancel_requested = False
        self._after_job = None

        self._build_ui()
        self._load_settings_into_ui()

        # بارگذاری اولیه لیست نمادها بدون نیاز به دکمه
        try:
            self._populate_symbol_list_from_tree()
        except Exception:
            logging.exception("خطا هنگام بارگذاری اولیه لیست نمادها.")

        # ثبت callback برای به‌روزرسانی خودکار وقتی current_tree تغییر می‌کند
        try:
            if self.current_tree is not None and hasattr(self.current_tree, 'on_update_callbacks'):
                def _cb_refresh():
                    try:
                        self.after(50, self._populate_symbol_list_from_tree)
                    except Exception:
                        pass
                # جلوگیری از اضافه شدن چندباره
                if _cb_refresh not in self.current_tree.on_update_callbacks:
                    self.current_tree.on_update_callbacks.append(_cb_refresh)
        except Exception:
            logging.exception("خطا هنگام ثبت callback برای به‌روزرسانی خودکار لیست نمادها.")

    # ------------------------
    # UI
    # ------------------------
    def _build_ui(self):
        frm = ttk.Frame(self, padding=8)
        frm.pack(fill="both", expand=True)

        # بالای پنجره: ورودی‌ها و انتخاب مسیر خروجی
        top = ttk.Frame(frm)
        top.pack(fill="x", pady=(0,6))

        ttk.Label(top, text="پوشه خروجی:").pack(side="left")
        self.out_entry = ttk.Entry(top, width=60)
        self.out_entry.pack(side="left", padx=6)
        self.out_entry.insert(0, settings_store.get("last_out_dir", "."))
        ttk.Button(top, text="انتخاب...", command=self._choose_out_dir).pack(side="left", padx=6)

        # بخش میانی: لیست نمادها با اسکرول و چک‌باکس‌ها
        mid = ttk.Frame(frm)
        mid.pack(fill="both", expand=True)

        left_panel = ttk.Frame(mid)
        left_panel.pack(side="left", fill="both", expand=True, padx=(0,6))

        ttk.Label(left_panel, text="نمادهای فیلترشده (انتخاب برای دانلود):", font=("Tahoma", 10, "bold")).pack(anchor="w", pady=(0,4))

        # کانتینر اسکرول‌شونده برای چک‌باکس‌ها
        self.symbol_canvas = tk.Canvas(left_panel, highlightthickness=0)
        self.symbol_vscroll = ttk.Scrollbar(left_panel, orient="vertical", command=self.symbol_canvas.yview)
        self.symbol_canvas.configure(yscrollcommand=self.symbol_vscroll.set)
        self.symbol_canvas.pack(side="left", fill="both", expand=True)
        self.symbol_vscroll.pack(side="right", fill="y")
        self.symbol_inner = ttk.Frame(self.symbol_canvas)
        self.symbol_canvas.create_window((0,0), window=self.symbol_inner, anchor="nw")
        self.symbol_inner.bind("<Configure>", lambda e: self.symbol_canvas.configure(scrollregion=self.symbol_canvas.bbox("all")))

        # دکمه‌های انتخاب همه / هیچ
        btns = ttk.Frame(left_panel)
        btns.pack(fill="x", pady=6)
        ttk.Button(btns, text="انتخاب همه", command=self._select_all_symbols).pack(side="left", padx=4)
        ttk.Button(btns, text="لغو انتخاب همه", command=self._deselect_all_symbols).pack(side="left", padx=4)
        ttk.Button(btns, text="بازسازی لیست (از جدول)", command=self._populate_symbol_list_from_tree).pack(side="right", padx=4)

        # پنل راست: تنظیمات و وضعیت
        right_panel = ttk.Frame(mid, width=320)
        right_panel.pack(side="right", fill="y")

        ttk.Label(right_panel, text="قالب لینک‌ها (قابل ویرایش):", font=("Tahoma", 10, "bold")).pack(anchor="w", pady=(0,6))
        ttk.Label(right_panel, text="قالب لینک حقیقی/حقوقی:").pack(anchor="w")
        self.client_url_text = tk.Text(right_panel, height=2, width=40, wrap="none")
        self.client_url_text.pack(fill="x", pady=4)
        ttk.Label(right_panel, text="قالب لینک قیمت:").pack(anchor="w")
        self.price_url_text = tk.Text(right_panel, height=2, width=40, wrap="none")
        self.price_url_text.pack(fill="x", pady=4)

        # فعال‌سازی میانبرهای کپی/پیست برای تکست‌ها
        def _bind_copy_paste_text(widget):
            widget.bind("<Control-c>", lambda e: widget.event_generate("<<Copy>>"))
            widget.bind("<Control-x>", lambda e: widget.event_generate("<<Cut>>"))
            widget.bind("<Control-v>", lambda e: widget.event_generate("<<Paste>>"))
        _bind_copy_paste_text(self.client_url_text)
        _bind_copy_paste_text(self.price_url_text)

        ttk.Separator(right_panel, orient="horizontal").pack(fill="x", pady=8)
        # دکمه دانلود گروهی
        self.download_btn = ttk.Button(right_panel, text="دانلود انتخاب‌شده", command=self._on_download_selected)
        self.download_btn.pack(fill="x", pady=(4,6))

        self.cancel_btn = ttk.Button(right_panel, text="لغو دانلود", command=self._request_cancel)
        self.cancel_btn.pack(fill="x", pady=(0,6))
        self.cancel_btn.config(state="disabled")

        # وضعیت و نوار پیشرفت
        ttk.Label(right_panel, text="وضعیت:").pack(anchor="w", pady=(6,0))
        self.status_label = ttk.Label(right_panel, text="آماده", foreground="green")
        self.status_label.pack(anchor="w", pady=(0,6))

        self.progress = ttk.Progressbar(right_panel, orient="horizontal", mode="determinate")
        self.progress.pack(fill="x", pady=(0,6))

        self.eta_label = ttk.Label(right_panel, text="")
        self.eta_label.pack(anchor="w")

        # لاگ پایین
        bottom = ttk.Frame(frm)
        bottom.pack(fill="both", expand=False, pady=(8,0))
        ttk.Label(bottom, text="لاگ:").pack(anchor="w")
        self.log_text = tk.Text(bottom, height=8, wrap="none")
        self.log_text.pack(fill="both", expand=True)
        vscroll2 = ttk.Scrollbar(bottom, orient="vertical", command=self.log_text.yview)
        vscroll2.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=vscroll2.set)

    # ------------------------
    # مدیریت لیست نمادها (چک‌باکس‌ها)
    # ------------------------
    def _clear_symbol_widgets(self):
        for w in self.symbol_inner.winfo_children():
            w.destroy()
        self._symbol_vars: List[Tuple[str, tk.BooleanVar, str]] = []  # (symbol, var, insCode)

    def _populate_symbol_list_from_tree(self):
        """
        لیست نمادها را از current_tree.df می‌سازد (نمادها از ستون‌های ممکن استخراج می‌شوند).
        مرتب‌سازی الفبایی و نمایش چک‌باکس برای هر نماد (پیش‌فرض تیک‌خورده).
        """
        self._clear_symbol_widgets()
        if not self.current_tree:
            self._log("هیچ جدول فعالی برای بارگذاری نمادها ارسال نشده است.")
            return

        # تلاش برای گرفتن df از چند منبع ممکن
        df = None
        for attr in ('df', 'base_df', 'dataframe'):
            df = getattr(self.current_tree, attr, None)
            if df is not None:
                break
        if df is None or df.empty:
            self._log("DataFrame جدول خالی است یا وجود ندارد. لطفاً ابتدا جدول را بارگذاری یا فیلتر کنید.")
            return

        # پیدا کردن ستون نماد با جستجوی گسترده (case-insensitive, حذف نیم‌فاصله)
        def normalize_col(c): return str(c).replace('\u200c','').strip().lower()
        cols = list(df.columns)
        norm_map = {normalize_col(c): c for c in cols}
        candidates = ['نماد','symbol','ticker','tiker','نماد_']
        symbol_col = None
        for cand in candidates:
            nc = cand.replace('\u200c','').strip().lower()
            if nc in norm_map:
                symbol_col = norm_map[nc]; break
        if symbol_col is None:
            for c in cols:
                lc = normalize_col(c)
                if 'نماد' in lc or 'symbol' in lc or 'ticker' in lc:
                    symbol_col = c; break
        if symbol_col is None:
            self._log(f"ستون نماد در DataFrame پیدا نشد. ستون‌های موجود: {cols}")
            return

        # پیدا کردن insCode مشابه
        ins_col = None
        for cand in ['inscode','ins_code','کد_داخلی','کد داخلی','کد']:
            if cand in norm_map:
                ins_col = norm_map[cand]; break
        if ins_col is None:
            for c in cols:
                lc = normalize_col(c)
                if 'ins' in lc or 'کد' in lc:
                    ins_col = c; break

        # استخراج و مرتب‌سازی یکتا
        seen = {}
        for _, row in df.iterrows():
            sym = str(row.get(symbol_col, '')).strip()
            if not sym:
                continue
            ins = str(row.get(ins_col, '')).strip() if ins_col else ''
            if sym not in seen:
                seen[sym] = ins
        sorted_syms = sorted(seen.items(), key=lambda x: x[0])

        for sym, ins in sorted_syms:
            var = tk.BooleanVar(value=True)
            frame = ttk.Frame(self.symbol_inner)
            frame.pack(fill="x", pady=1, padx=2)
            cb = ttk.Checkbutton(frame, text=f"{sym}  [{ins}]" if ins else sym, variable=var)
            cb.pack(side="left", anchor="w")
            self._symbol_vars.append((sym, var, ins))

        self._log(f"{len(self._symbol_vars)} نماد از جدول بارگذاری شد.")

    def _select_all_symbols(self):
        for _, var, _ in getattr(self, "_symbol_vars", []):
            var.set(True)

    def _deselect_all_symbols(self):
        for _, var, _ in getattr(self, "_symbol_vars", []):
            var.set(False)

    # ------------------------
    # دانلود گروهی (صف و پردازش ترتیبی با after)
    # ------------------------
    def _on_download_selected(self):
        # ساخت صف دانلود از نمادهای تیک‌خورده
        selected = [(sym, ins) for (sym, var, ins) in getattr(self, "_symbol_vars", []) if var.get()]
        if not selected:
            messagebox.showwarning("هیچ نمادی انتخاب نشده", "لطفاً حداقل یک نماد را برای دانلود انتخاب کنید.")
            return
        out_dir = self.out_entry.get().strip() or "."
        settings_store["last_out_dir"] = out_dir
        save_settings(settings_store)

        # ساخت صف: هر آیتم دیکشنری {insCode, symbol}
        self.download_queue = []
        for sym, ins in selected:
            ins_code = ins
            if not ins_code and self.current_tree and getattr(self.current_tree, "df", None) is not None:
                df = self.current_tree.df
                # تلاش برای یافتن insCode متناظر با نماد
                mask = df.apply(lambda r: str(r.get('نماد', '')).strip() == sym or str(r.get('symbol', '')).strip() == sym, axis=1)
                if mask.any():
                    row = df[mask].iloc[0]
                    for c in ['insCode', 'کد_داخلی', 'کد داخلی', 'کد']:
                        if c in row and row[c]:
                            ins_code = str(row[c])
                            break
            if not ins_code:
                self._log(f"خطا: کد داخلی (insCode) برای نماد {sym} پیدا نشد؛ این نماد نادیده گرفته می‌شود.")
                continue
            self.download_queue.append({"symbol": sym, "insCode": ins_code, "out_dir": out_dir})

        if not self.download_queue:
            messagebox.showwarning("هیچ نمادی برای دانلود", "هیچ نمادی با insCode معتبر برای دانلود پیدا نشد.")
            return

        # آماده‌سازی وضعیت و شروع پردازش صف
        self._is_downloading = True
        self._cancel_requested = False
        self._processed_count = 0
        self._download_start_time = datetime.now()
        total = len(self.download_queue)
        self.progress["maximum"] = total
        self.progress["value"] = 0
        self.status_label.config(text=f"در حال دانلود 0/{total}", foreground="orange")
        self.download_btn.config(state="disabled")
        self.cancel_btn.config(state="normal")
        self._log(f"شروع دانلود گروهی برای {total} نماد.")
        # شروع اولین آیتم با after
        self._process_next_in_queue()

    def _process_next_in_queue(self):
        if self._cancel_requested:
            self._finish_downloads(cancelled=True)
            return
        if not self.download_queue:
            self._finish_downloads(cancelled=False)
            return

        item = self.download_queue.pop(0)
        sym = item["symbol"]
        ins = item["insCode"]
        out_dir = item["out_dir"]

        # لاگ شروع
        self._log(f"در حال دانلود برای {sym} ({ins}) ...")
        start = datetime.now()

        # فراخوانی تابع دانلود/ذخیره
        try:
            ok, msg = fetch_and_save_for_symbol(ins, sym, out_dir=out_dir,
                                               client_url_template=self.client_url_text.get("1.0", "end").strip() or None,
                                               price_url_template=self.price_url_text.get("1.0", "end").strip() or None)
            elapsed = (datetime.now() - start).total_seconds()
            self._processed_count += 1
            total_done = self._processed_count
            # به‌روزرسانی نوار پیشرفت و وضعیت
            self.progress["value"] = total_done
            remaining = int(self.progress["maximum"] - total_done)
            avg = (datetime.now() - self._download_start_time).total_seconds() / total_done if total_done > 0 else 0
            eta_seconds = int(avg * remaining)
            eta_text = self._format_eta(eta_seconds)
            self.status_label.config(text=f"در حال دانلود {total_done}/{int(self.progress['maximum'])}", foreground="orange")
            self.eta_label.config(text=f"باقی: {remaining}؛ حدوداً {eta_text}")
            if ok:
                self._log(f"ذخیره شد: {msg} (زمان: {elapsed:.1f}s)")
            else:
                self._log(f"خطا برای {sym}: {msg}")
        except Exception as e:
            logging.exception("خطا هنگام دانلود گروهی:")
            self._log(f"خطای داخلی برای {sym}: {e}")

        # زمان‌بندی پردازش آیتم بعدی با یک وقفهٔ کوتاه تا UI فرصت به‌روزرسانی داشته باشد
        self._after_job = self.after(200, self._process_next_in_queue)

    def _format_eta(self, seconds: int) -> str:
        if seconds <= 0:
            return "کمتر از یک دقیقه"
        m, s = divmod(seconds, 60)
        h, m = divmod(m, 60)
        if h > 0:
            return f"{h}س {m}د"
        if m > 0:
            return f"{m}د {s}ث"
        return f"{s}ث"

    def _finish_downloads(self, cancelled: bool = False):
        total_done = int(self.progress["value"])
        total = int(self.progress["maximum"])
        if cancelled:
            self._log(f"دانلودها لغو شد. انجام‌شده: {total_done} از {total}.")
            self.status_label.config(text=f"لغو شد ({total_done}/{total})", foreground="red")
        else:
            self._log(f"دانلود گروهی به پایان رسید. فایل‌های ساخته‌شده: {total_done}.")
            self.status_label.config(text=f"پایان ({total_done}/{total})", foreground="green")
            self.eta_label.config(text="")
        self._is_downloading = False
        self.download_btn.config(state="normal")
        self.cancel_btn.config(state="disabled")
        # پاکسازی صف و after job
        self.download_queue = []
        if self._after_job:
            try:
                self.after_cancel(self._after_job)
            except Exception:
                pass
            self._after_job = None

    def _request_cancel(self):
        if not self._is_downloading:
            return
        self._cancel_requested = True
        self._log("درخواست لغو دریافت شد...")

    # ------------------------
    # کمکی‌ها و لاگ
    # ------------------------
    def _choose_out_dir(self):
        d = filedialog.askdirectory(initialdir=settings_store.get("last_out_dir", "."))
        if d:
            self.out_entry.delete(0, tk.END)
            self.out_entry.insert(0, d)

    def _log(self, msg: str):
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        try:
            self.log_text.insert("end", f"[{ts}] {msg}\n")
            self.log_text.see("end")
        except Exception:
            pass
        logging.info(msg)

    def _load_settings_into_ui(self):
        try:
            default_client = settings_store.get("client_url_template",
                DEFAULTS["client_url_template"])
            default_price = settings_store.get("price_url_template",
                DEFAULTS["price_url_template"])

            self.client_url_text.delete("1.0", "end")
            self.client_url_text.insert("1.0", default_client)
            self.price_url_text.delete("1.0", "end")
            self.price_url_text.insert("1.0", default_price)

            self.out_entry.delete(0, tk.END)
            self.out_entry.insert(0, settings_store.get("last_out_dir", DEFAULTS["last_out_dir"]))
        except Exception:
            logging.exception("خطا هنگام بارگذاری تنظیمات در UI.")

    def _on_save_settings(self):
        try:
            settings_store["client_url_template"] = self.client_url_text.get("1.0", "end").strip()
            settings_store["price_url_template"] = self.price_url_text.get("1.0", "end").strip()
            settings_store["last_out_dir"] = self.out_entry.get().strip() or "."
            save_settings(settings_store)
            self._log("تنظیمات ذخیره شد.")
            messagebox.showinfo("ذخیره تنظیمات", "تنظیمات با موفقیت ذخیره شد.")
        except Exception:
            logging.exception("خطا هنگام ذخیره تنظیمات.")
            messagebox.showerror("خطا", "خطا هنگام ذخیره تنظیمات. لاگ را بررسی کنید.")

    def _clear_log(self):
        try:
            self.log_text.delete("1.0", "end")
        except Exception:
            pass

    def _on_close(self):
        # اگر در حال دانلود هستیم، از کاربر تایید بگیر
        if self._is_downloading:
            if not messagebox.askyesno("در حال دانلود", "دانلود در حال انجام است. آیا مطمئنید می‌خواهید پنجره را ببندید و دانلود را لغو کنید؟"):
                return
            self._request_cancel()
        try:
            if self._after_job:
                try:
                    self.after_cancel(self._after_job)
                except Exception:
                    pass
                self._after_job = None
        except Exception:
            pass
        try:
            self.destroy()
        except Exception:
            pass

# ------------------------
# صادر شده‌ها
# ------------------------
__all__ = ["ClientTypeExportWindow", "fetch_and_save_for_symbol", "merge_client_and_price"]

# اگر به صورت مستقیم اجرا شد، پنجرهٔ تست را باز کن
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    w = ClientTypeExportWindow(root)
    w.ins_entry.insert(0, "36844527173896115")
    w.sym_entry.insert(0, "زفجر")
    root.mainloop()
