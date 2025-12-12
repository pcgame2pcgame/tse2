# core_part1.py
# بخش اول از ماژول core برای نمایشگر TSETMC
# شامل تنظیمات، نگهداری تنظیمات، نرمال‌سازی متن، نگاشت‌ها و توابع کمکی پایه
# نیازمندی‌ها: pandas, requests
# نصب: pip install pandas requests

import os
import re
import json
import time
import traceback
import requests
import pandas as pd
import tkinter as tk
from tkinter import ttk, Menu, messagebox
from tkinter import font as tkfont
import webbrowser


# ------------------------
# ثابت‌ها و فایل تنظیمات
# ------------------------
URL_DEFAULT = "https://old.tsetmc.com/tsev2/data/MarketWatchPlus.aspx?h=0&r=0"
DEFAULT_EXPORT_NAME = "tsetmc.csv"
SETTINGS_FILE = "tsetmc_settings.json"

# ------------------------
# بارگذاری و ذخیره تنظیمات (مقاوم در برابر فایل خراب)
# ------------------------
def load_settings():
    """بارگذاری امن تنظیمات JSON. در صورت خراب بودن فایل، آن را جابجا می‌کند و دیکشنری خالی برمی‌گرداند."""
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except json.JSONDecodeError:
            try:
                bak = SETTINGS_FILE + ".corrupt"
                os.replace(SETTINGS_FILE, bak)
            except Exception:
                pass
            return {}
        except Exception:
            return {}
    return {}

def save_settings(d):
    try:
        with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
            json.dump(d, f, ensure_ascii=False, indent=2)
    except Exception:
        # بی‌صدا شکست می‌خورد تا UI قفل نشود
        pass

settings_store = load_settings()

# ------------------------
# نرمال‌سازی متن (حروف عربی -> فارسی، ارقام فارسی -> لاتین، حذف نیم‌فاصله)
# ------------------------
ARABIC_TO_PERSIAN = {
    'ي': 'ی', 'ى': 'ی', 'ك': 'ک', 'ﻻ': 'لا', 'ة': 'ه',
    'ؤ': 'و', 'إ': 'ا', 'أ': 'ا', 'آ': 'ا',
    'ئ': 'ی', '\u200c': '', '\u0640': ''
}
PERSIAN_DIGITS = str.maketrans('۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩', '01234567890123456789')
_ar_re = re.compile('|'.join(map(re.escape, ARABIC_TO_PERSIAN.keys())))

def normalize_text(s):
    """نرمال‌سازی متن فارسی/عربی و ارقام؛ خروجی رشتهٔ تمیز شده."""
    if s is None:
        return ''
    s = str(s)
    s = _ar_re.sub(lambda m: ARABIC_TO_PERSIAN[m.group(0)], s)
    s = s.translate(PERSIAN_DIGITS)
    s = s.replace('\u200c', ' ').strip()
    s = re.sub(r'\s+', ' ', s)
    return s

# ------------------------
# نگاشت صنایع و برچسب بازار
# ------------------------
INDUSTRY_MAP_LIST = [
    ['01', 'زراعت و خدمات وابسته'], ['02', 'جنگلداري و ماهيگيري'],
    ['10', 'استخراج زغال سنگ'], ['11', 'استخراج نفت گاز و خدمات جنبي جز اکتشاف'],
    ['13', 'استخراج کانه هاي فلزي'], ['14', 'استخراج ساير معادن'],
    ['15', 'حذف شده- فرآورده‌هاي غذايي و آشاميدني'], ['17', 'منسوجات'],
    ['19', 'دباغي، پرداخت چرم و ساخت انواع پاپوش'], ['20', 'محصولات چوبي'],
    ['21', 'محصولات كاغذي'], ['22', 'انتشار، چاپ و تکثير'],
    ['23', 'فراورده هاي نفتي، كک و سوخت هسته اي'], ['24', 'حذف شده-مواد و محصولات شيميايي'],
    ['25', 'لاستيك و پلاستيك'], ['26', 'توليد محصولات كامپيوتري الكترونيكي ونوري'],
    ['27', 'فلزات اساسي'], ['28', 'ساخت محصولات فلزي'],
    ['29', 'ماشين آلات و تجهيزات'], ['31', 'ماشين آلات و دستگاه‌هاي برقي'],
    ['32', 'ساخت دستگاه‌ها و وسايل ارتباطي'], ['33', 'ابزارپزشکي، اپتيکي و اندازه‌گيري'],
    ['34', 'خودرو و ساخت قطعات'], ['35', 'ساير تجهيزات حمل و نقل'],
    ['36', 'مبلمان و مصنوعات ديگر'], ['38', 'قند و شكر'],
    ['39', 'شرکتهاي چند رشته اي صنعتي'], ['40', 'عرضه برق، گاز، بخاروآب گرم'],
    ['41', 'جمع آوري، تصفيه و توزيع آب'], ['42', 'محصولات غذايي و آشاميدني به جز قند و شكر'],
    ['43', 'مواد و محصولات دارويي'], ['44', 'محصولات شيميايي'],
    ['45', 'پيمانكاري صنعتي'], ['46', 'تجارت عمده فروشي به جز وسايل نقليه موتور'],
    ['47', 'خرده فروشي،باستثناي وسايل نقليه موتوري'], ['49', 'كاشي و سراميك'],
    ['50', 'تجارت عمده وخرده فروشي وسائط نقليه موتور'], ['51', 'حمل و نقل هوايي'],
    ['52', 'انبارداري و حمايت از فعاليتهاي حمل و نقل'], ['53', 'سيمان، آهك و گچ'],
    ['54', 'ساير محصولات كاني غيرفلزي'], ['55', 'هتل و رستوران'],
    ['56', 'سرمايه گذاريها'], ['57', 'بانكها و موسسات اعتباري'],
    ['58', 'ساير واسطه گريهاي مالي'], ['59', 'اوراق حق تقدم استفاده از تسهيلات مسكن'],
    ['60', 'حمل ونقل، انبارداري و ارتباطات'], ['61', 'حمل و نقل آبی'],
    ['63', 'فعاليت های پشتیبانی و کمکی حمل و نقل'], ['64', 'مخابرات'],
    ['65', 'واسطه‌گری‌های مالی و پولی'], ['66', 'بیمه وصندوق بازنشستگی به جز تامین اجتماعی'],
    ['67', 'فعالیت‌هاي کمکی به نهادهای مالی واسط'], ['68', 'صندوق سرمایه گذاری قابل معامله'],
    ['69', 'اوراق تامین مالی'], ['70', 'انبوه سازی، املاک و مستغلات'],
    ['71', 'فعالیت مهندسی، تجزیه، تحلیل و آزمایش فنی'], ['72', 'رایانه و فعالیت‌های وابسته به آن'],
    ['73', 'اطلاعات و ارتباطات'], ['74', 'خدمات فنی و مهندسی'],
    ['76', 'اوراق بهادار مبتنی بر دارایی فکری'], ['77', 'فعالبت های اجاره و لیزینگ'],
    ['80', 'تبلیغات و بازارپژوهی'], ['82', 'فعالیت پشتیبانی اجرائی اداری و حمایت کسب'],
    ['84', 'سلامت انسان و مددکاری اجتماعی'], ['90', 'فعالیت های هنری، سرگرمی و خلاقانه'],
    ['93', 'فعالیت‌های فرهنگی و ورزشی'], ['98', 'گروه اوراق غیر فعال'],
    ['X1', 'شاخص']
]
INDUSTRY_MAP = {k: v for k, v in INDUSTRY_MAP_LIST}

MARKET_LABELS = {
    '300': 'بورس', '303': 'فرابورس', '309': 'پایه',
    '301': 'مشارکت', '304': 'آتی', '305': 'صندوق', '306': 'مرابحه و اجاره',
    '307': 'تسهیلات مسکن', '308': 'سلف', '311': 'اختیار خ ض', '312': 'اختیار ف ط',
    '313': 'بازار نوآفرین رشد پایه', '315': 'صندوق کالا', '320': 'اختیار خرید ض',
    '321': 'اختیار ف ط', '380': 'صندوق طلا و کالا', '400': 'حق بورس',
    '403': 'حق فرابورس', '404': 'حق پایه', '701': 'زعفران و سکه',
    '706': 'مرابحه دولت اراد', '803': 'بار برق', '804': 'بار برق',
    '200': 'سلف انرژی', '206': 'صکوک', '201': 'گواهی', '208': 'صکوک'
}

# ------------------------
# نگاشت نام ستون‌ها (پیش‌فرض)
# ------------------------
COLUMN_NAME_MAP = {}
line_map = {1: "خط1", 2: "خط2", 3: "خط3", 4: "خط4", 5: "خط5"}
c_map = {
    2: "تعداد فروشنده",
    3: "تعداد خریدار",
    4: "قیمت خریدار",
    5: "قیمت فروشنده",
    6: "حجم خریدار",
    7: "حجم فروشنده"
}
for lv in range(1, 6):
    for c in range(2, 8):
        key = f"S3_L{lv}_C{c}"
        display = f"{c_map[c]} {line_map[lv]}"
        COLUMN_NAME_MAP[key] = display
        COLUMN_NAME_MAP[key.lower()] = display
        COLUMN_NAME_MAP[key.upper()] = display

COLUMN_NAME_MAP.update({
    "کد_داخلی": "کد داخلی",
    "کد_بین_المللی": "کد بین المللی",
    "نماد": "نماد",
    "نام_شرکت": "نام شرکت",
    "قیمت_پایانی": "قیمت پایانی",
    "قیمت_آخرین_معامله": "قیمت آخرین معامله",
    "تعداد_معاملات": "تعداد معاملات",
    "حجم_معاملات": "حجم معاملات",
    "ارزش_معاملات": "ارزش معاملات",
    "کمترین_قیمت": "کمترین قیمت",
    "بیشترین_قیمت": "بیشترین قیمت",
    "قیمت_دیروز": "قیمت دیروز",
    "تعداد_کل_سهام": "تعداد کل سهام",
    "کد_بازار": "کد بازار",
    "گروه_صنعت": "گروه صنعت",
    "نوع_صنعت": "نوع صنعت",
    "ارزش بازار همت": "ارزش بازار همت",
    "PE": "PE",
    "صف خرید": "صف خرید",
    "صف فروش": "صف فروش"
})

# ادغام نگاشت ذخیره‌شده در تنظیمات (در صورت وجود)
if settings_store.get('column_name_map') is None:
    settings_store['column_name_map'] = COLUMN_NAME_MAP
else:
    saved_map = settings_store.get('column_name_map', {})
    merged = COLUMN_NAME_MAP.copy()
    merged.update(saved_map)
    COLUMN_NAME_MAP = merged
    settings_store['column_name_map'] = COLUMN_NAME_MAP

# ------------------------
# توابع دریافت و پارس اولیه داده‌ها
# ------------------------
def fetch_sections(url=URL_DEFAULT, timeout=30):
    """دریافت متن و تقسیم به بخش‌ها بر اساس @"""
    resp = requests.get(url, timeout=timeout)
    resp.encoding = 'utf-8'
    return resp.text.split('@')

def parse_section(section_text: str, mapping=None) -> pd.DataFrame:
    """
    پارس یک بخش از متن TSETMC که با ; جدا شده است.
    اگر mapping داده شود، ایندکس‌های مشخص را به نام ستون تبدیل می‌کند.
    """
    rows = [r for r in section_text.split(';') if r.strip()]
    data = []
    for i, row in enumerate(rows):
        fields = row.split(',')
        rec = {'ردیف': i+1}
        if mapping:
            for idx, name in mapping.items():
                rec[name] = fields[idx] if idx < len(fields) else ''
        else:
            for j, val in enumerate(fields):
                rec[f"ستون{j}"] = val
        data.append(rec)
    if not data:
        return pd.DataFrame()
    return pd.DataFrame(data)

def merge_section3_into2(df2: pd.DataFrame, df3: pd.DataFrame) -> pd.DataFrame:
    """
    ادغام اطلاعات بخش 3 (S3) به بخش 2 بر اساس کلید کد.
    خروجی: df2 با ستون‌های اضافی S3_L{1..5}_C{2..7}
    """
    if df2 is None or df3 is None or df2.empty or df3.empty:
        return df2.copy() if df2 is not None else pd.DataFrame()
    key_df3 = 'ستون0'
    key_df2 = 'کد_داخلی' if 'کد_داخلی' in df2.columns else ('ستون0' if 'ستون0' in df2.columns else None)
    if key_df2 is None or key_df3 not in df3.columns:
        return df2.copy()
    block_cols = [f'ستون{i}' for i in range(2, 8)]
    for c in block_cols:
        if c not in df3.columns:
            df3[c] = ''
    level_col = 'ستون1'
    if level_col not in df3.columns:
        df3[level_col] = ''
    s3_map = {}
    for _, row in df3.iterrows():
        k = str(row.get(key_df3, '')).strip()
        lv_str = str(row.get(level_col, '')).strip()
        if not k or not lv_str.isdigit():
            continue
        lv = int(lv_str)
        if lv < 1 or lv > 5:
            continue
        vals = [row.get(col, '') for col in block_cols]
        s3_map.setdefault(k, {})[lv] = vals
    extra_col_names = []
    for lv in range(1, 6):
        for j in range(2, 8):
            extra_col_names.append(f"S3_L{lv}_C{j}")
    extra_data = []
    for _, row in df2.iterrows():
        k = str(row.get(key_df2, '')).strip()
        lv_map = s3_map.get(k, {})
        row_extra = []
        for lv in range(1, 6):
            vals = lv_map.get(lv)
            if vals:
                row_extra.extend(vals)
            else:
                row_extra.extend([''] * 6)
        extra_data.append(row_extra)
    extra_df = pd.DataFrame(extra_data, columns=extra_col_names, index=df2.index)
    merged = pd.concat([df2.reset_index(drop=True), extra_df.reset_index(drop=True)], axis=1)
    return merged

# ------------------------
# کلید مرتب‌سازی برای مقادیر ترکیبی عدد/متن
# ------------------------
_token_re = re.compile(r'(\d+|\D+)')
def to_sort_key(s):
    """تبدیل رشته به کلید مرتب‌سازی که اعداد را عددی و متن را حروفی در نظر می‌گیرد."""
    s = normalize_text(s)
    if s == '':
        return (2, '')
    try:
        if re.fullmatch(r'[-+]?\d+(\.\d+)?', s):
            return (0, float(s))
    except Exception:
        pass
    parts = _token_re.findall(s)
    key_parts = []
    for p in parts:
        if p.isdigit():
            key_parts.append((0, int(p)))
        else:
            key_parts.append((1, p.lower()))
    flat = []
    for t in key_parts:
        flat.append(t[0]); flat.append(t[1])
    return (1, tuple(flat))

# ------------------------
# کمک‌کننده‌های کوچک و مقداردهی پیش‌فرض تنظیمات
# ------------------------
def _ensure_settings_defaults():
    changed = False
    if 'column_name_map' not in settings_store:
        settings_store['column_name_map'] = COLUMN_NAME_MAP
        changed = True
    if 'visible_columns' not in settings_store:
        settings_store['visible_columns'] = {}
        changed = True
    if 'saved_filters_full' not in settings_store:
        settings_store['saved_filters_full'] = []
        changed = True
    if 'bottom_visible_columns' not in settings_store:
        settings_store['bottom_visible_columns'] = None
    if changed:
        try:
            save_settings(settings_store)
        except Exception:
            pass

_ensure_settings_defaults()

def get_column_case_insensitive(df: pd.DataFrame, col_name: str):
    """خواندن ستون از DataFrame به صورت case-insensitive؛ None در صورت عدم وجود."""
    if df is None or col_name is None:
        return None
    if col_name in df.columns:
        return df[col_name]
    lower = col_name.lower()
    for c in df.columns:
        if c.lower() == lower:
            return df[c]
    return None

# ------------------------
# FIELD_MAPPING نمونه برای بخش 2
# ------------------------
FIELD_MAPPING = {
    0: "کد_داخلی",
    1: "کد_بین_المللی",
    2: "نماد",
    3: "نام_شرکت",
    4: "زمان_آخرین_معامله",
    5: "اولین_قیمت",
    6: "قیمت_پایانی",
    7: "قیمت_آخرین_معامله",
    8: "تعداد_معاملات",
    9: "حجم_معاملات",
    10: "ارزش_معاملات",
    11: "کمترین_قیمت",
    12: "بیشترین_قیمت",
    13: "قیمت_دیروز",
    14: "EPS",
    15: "حجم_مبنا",
    16: "تعداد_بازدید_کننده",
    17: "بازار_اصلی",
    18: "گروه_صنعت",
    19: "حداکثر_قیمت_مجاز",
    20: "حداقل_قیمت_مجاز",
    21: "تعداد_کل_سهام",
    22: "کد_بازار",
    23: "NAV",
    24: "موقعیت_های_باز",
    25: "دسته_بندی_تخصصی"
}

# ------------------------
# صادرات نمادها برای استفاده در فایل دوم
# ------------------------
__all__ = [
    "URL_DEFAULT", "DEFAULT_EXPORT_NAME", "SETTINGS_FILE",
    "load_settings", "save_settings", "settings_store",
    "normalize_text", "INDUSTRY_MAP", "MARKET_LABELS",
    "COLUMN_NAME_MAP", "fetch_sections", "parse_section",
    "merge_section3_into2", "to_sort_key", "FIELD_MAPPING",
    "get_column_case_insensitive"
]

# پایان بخش اول
# core_part2.py
# بخش دوم از ماژول core برای نمایشگر TSETMC
# شامل ویجت‌های رابط کاربری: AdvancedTreeview، BottomStatsTable، ScrollableToplevel،
# ColumnSettingsDialog و AppSettingsDialog
# این فایل به core_part1.py وابسته است و باید آن را در همان پوشه داشته باشید:
# from core_part1 import <functions and constants>



# ------------------------
# AdvancedTreeview
# ------------------------
class AdvancedTreeview(ttk.Treeview):
    """
    Treeview پیشرفته با:
    - نگهداری base_df و df فعلی
    - فیلترها (active_filters)
    - نمایش مقادیر محاسبه‌شده مانند 'ارزش بازار همت' و 'PE'
    - منوی راست کلیک برای کپی و فیلتر سریع
    """
    def __init__(self, parent, df: pd.DataFrame, app_runtime_log: dict = None, **kwargs):
        super().__init__(parent, show="headings", **kwargs)
        self.app_runtime_log = app_runtime_log if app_runtime_log is not None else {}
        self.base_df = df.copy() if df is not None else pd.DataFrame()
        self.df = self.base_df.copy()
        self.norm_df = pd.DataFrame()
        self.active_filters = []  # list of {'desc':..., 'func':..., 'enabled':True}
        self.visible_columns = {col: True for col in list(self.df.columns)}
        saved_vis = settings_store.get('visible_columns', {})
        for c, v in saved_vis.items():
            self.visible_columns[c] = v
        self.on_update_callbacks = []
        self._sort_state = {}
        # prepare data (compute derived cols) and build UI
        self._prepare_dataframe()
        self._setup_columns(auto_optimize=True)
        self._create_context_menu()
        self.tag_configure('search_match', background='yellow')
        self.bind("<Button-3>", self._on_right_click)
        self._load_batch()

    def _prepare_dataframe(self):
        """Normalize text columns and compute derived columns."""
        for col in list(self.base_df.columns):
            s = self.base_df[col].astype(str).fillna('')
            s_norm = s.map(normalize_text)
            self.base_df[col] = s_norm

        # ارزش بازار همت = قیمت_پایانی * تعداد_کل_سهام / 1e13
        if 'قیمت_پایانی' in self.base_df.columns and 'تعداد_کل_سهام' in self.base_df.columns:
            try:
                num_price = pd.to_numeric(self.base_df['قیمت_پایانی'], errors='coerce')
                num_shares = pd.to_numeric(self.base_df['تعداد_کل_سهام'], errors='coerce')
                mv = (num_price * num_shares) / 1e13
                self.base_df['ارزش بازار همت'] = mv
            except Exception:
                self.base_df['ارزش بازار همت'] = pd.NA

        # PE = قیمت_آخرین_معامله / EPS
        if 'قیمت_آخرین_معامله' in self.base_df.columns and 'EPS' in self.base_df.columns:
            try:
                num_last = pd.to_numeric(self.base_df['قیمت_آخرین_معامله'], errors='coerce')
                num_eps = pd.to_numeric(self.base_df['EPS'], errors='coerce')
                pe = num_last / num_eps.replace({0: pd.NA})
                self.base_df['PE'] = pe
            except Exception:
                self.base_df['PE'] = pd.NA

        # نوع_صنعت از گروه_صنعت
        if 'گروه_صنعت' in self.base_df.columns:
            def map_industry(x):
                k = normalize_text(str(x)).strip()
                if not k:
                    return ''
                m = re.match(r'^([A-Za-z0-9Xx]+)', k)
                key = m.group(1) if m else k
                key = key.strip()
                if key.isdigit() and len(key) == 1:
                    key = key.zfill(2)
                return INDUSTRY_MAP.get(key, INDUSTRY_MAP.get(key.zfill(2), ''))
            self.base_df['نوع_صنعت'] = self.base_df['گروه_صنعت'].apply(map_industry)

        # صف خرید / صف فروش بر اساس S3 سطح 1 (محاسبه محافظه‌کارانه)
        self.base_df['صف خرید'] = 0.0
        self.base_df['صف فروش'] = 0.0

        def get_s3_numeric(idx_name):
            if idx_name in self.base_df.columns:
                return pd.to_numeric(self.base_df[idx_name], errors='coerce')
            for c in self.base_df.columns:
                if c.lower() == idx_name.lower():
                    return pd.to_numeric(self.base_df[c], errors='coerce')
            return pd.Series([pd.NA] * len(self.base_df), index=self.base_df.index)

        s_price_buyer_l1 = get_s3_numeric('S3_L1_C4')
        s_price_seller_l1 = get_s3_numeric('S3_L1_C5')
        s_vol_buyer_l1 = get_s3_numeric('S3_L1_C6')
        s_vol_seller_l1 = get_s3_numeric('S3_L1_C7')

        max_allowed = pd.to_numeric(self.base_df.get('حداکثر_قیمت_مجاز', pd.Series([pd.NA] * len(self.base_df))), errors='coerce')
        min_allowed = pd.to_numeric(self.base_df.get('حداقل_قیمت_مجاز', pd.Series([pd.NA] * len(self.base_df))), errors='coerce')

        try:
            buy_series = pd.Series(0.0, index=self.base_df.index)
            sell_series = pd.Series(0.0, index=self.base_df.index)
            cond_buy = (s_price_buyer_l1.notna()) & (max_allowed.notna()) & (s_price_buyer_l1 == max_allowed)
            buy_series.loc[cond_buy] = (s_vol_buyer_l1.loc[cond_buy].fillna(0).astype(float) * s_price_buyer_l1.loc[cond_buy].astype(float)).fillna(0)
            cond_sell = (s_price_seller_l1.notna()) & (min_allowed.notna()) & (s_price_seller_l1 == min_allowed)
            sell_series.loc[cond_sell] = (s_vol_seller_l1.loc[cond_sell].fillna(0).astype(float) * s_price_seller_l1.loc[cond_sell].astype(float)).fillna(0)
            self.base_df['صف خرید'] = buy_series.fillna(0)
            self.base_df['صف فروش'] = sell_series.fillna(0)
        except Exception:
            self.base_df['صف خرید'] = self.base_df['صف خرید'].fillna(0)
            self.base_df['صف فروش'] = self.base_df['صف فروش'].fillna(0)

        # set df and normalized df
        self.df = self.base_df.copy()
        self.norm_df = self._build_normalized_df(self.df)

    def _build_normalized_df(self, df):
        if df is None or df.empty:
            return pd.DataFrame()
        norm = pd.DataFrame(index=df.index)
        for col in df.columns:
            norm[col] = df[col].astype(str).fillna('').apply(normalize_text)
        return norm

    def _format_value_for_display(self, col, val):
        """فرمت نمایش برای ستون‌های خاص"""
        if col == 'ارزش بازار همت':
            try:
                if pd.isna(val):
                    return ''
                v = float(val)
                return f"{v:.1f}"
            except:
                return str(val)
        if col == 'PE':
            try:
                if pd.isna(val):
                    return ''
                v = float(val)
                return f"{v:.1f}"
            except:
                return str(val)
        if col in ('صف خرید', 'صف فروش'):
            try:
                if pd.isna(val):
                    return ''
                v = float(val)
                if abs(v - int(v)) < 1e-6:
                    return str(int(v))
                return f"{v:.2f}"
            except:
                return str(val)
        return '' if pd.isna(val) else str(val)

    def _compute_optimal_widths(self, sample_rows=200, char_width=7, padding=20, max_width=600):
        widths = {}
        if self.df is None or self.df.empty:
            return widths
        sample = self.df.head(sample_rows).astype(str).fillna('')
        for col in self.df.columns:
            header_len = len(COLUMN_NAME_MAP.get(col, col))
            if col in ('ارزش بازار همت', 'PE', 'صف خرید', 'صف فروش'):
                sample_col = self.df[col].apply(lambda x: self._format_value_for_display(col, x)).astype(str)
            else:
                sample_col = sample[col]
            max_cell_len = sample_col.map(len).max() if not sample_col.empty else 0
            est_chars = max(header_len, max_cell_len)
            widths[col] = int(min(max(80, est_chars * char_width + padding), max_width))
        return widths

    def _setup_columns(self, auto_optimize=False):
        cols = list(self.df.columns)
        self["columns"] = cols
        if auto_optimize:
            widths = self._compute_optimal_widths()
            self.app_runtime_log.setdefault('column_widths', {}).update(widths)
        else:
            widths = self.app_runtime_log.get('column_widths', {})
        display_cols = [c for c, v in self.visible_columns.items() if v]
        self.configure(displaycolumns=display_cols)
        for col in cols:
            header_text = COLUMN_NAME_MAP.get(col, COLUMN_NAME_MAP.get(col.lower(), col))
            self.heading(col, text=header_text, command=lambda c=col: self._on_heading_click(c))
            self.column(col, width=widths.get(col, 120), anchor="center", stretch=False)

    def _on_heading_click(self, col):
        asc = self._sort_state.get(col, True)
        self._sort_state[col] = not asc
        try:
            key_series = self.df[col].astype(str).apply(to_sort_key)
            self.df = self.df.assign(_sort_col=key_series)
        except Exception:
            self.df = self.df.assign(_sort_col=self.df[col].astype(str).apply(normalize_text))
        self.df = self.df.sort_values(by='_sort_col', ascending=asc, na_position='last').drop(columns=['_sort_col']).reset_index(drop=True)
        if 'ردیف' in self.df.columns:
            self.df['ردیف'] = range(1, len(self.df) + 1)
        self.norm_df = self._build_normalized_df(self.df)
        self._load_batch()

    def _load_batch(self):
        self.delete(*self.get_children())
        if self.df is None or self.df.empty:
            return
        cols = list(self.df.columns)
        chunk = 500
        rows = []
        for _, row in self.df.iterrows():
            display_row = []
            for col in cols:
                display_row.append(self._format_value_for_display(col, row[col]) if col in ('ارزش بازار همت','PE','صف خرید','صف فروش') else ('' if pd.isna(row[col]) else str(row[col])))
            rows.append(tuple(display_row))
        for i in range(0, len(rows), chunk):
            for r in rows[i:i+chunk]:
                self.insert("", tk.END, values=r)
            self.update_idletasks()

    # ------------------------
    # Context menu and clipboard helpers
    # ------------------------
    def _create_context_menu(self):
        self.menu = Menu(self, tearoff=0)
        self.menu.add_command(label="کپی مقدار سلول", command=self.copy_cell)
        self.menu.add_command(label="کپی کل ردیف", command=self.copy_row)
        self.menu.add_separator()
        self.menu.add_command(label="باز کردن صفحه نماد", command=self.open_symbol_page)
        self.menu.add_command(label="فیلتر بر اساس این نماد", command=self.filter_by_symbol_from_selection)

    def _on_right_click(self, event):
        iid = self.identify_row(event.y)
        if iid:
            self.selection_set(iid)
            try:
                self.menu.tk_popup(event.x_root, event.y_root)
            finally:
                self.menu.grab_release()

    def copy_cell(self):
        sel = self.selection()
        if not sel:
            return
        vals = self.item(sel[0], 'values')
        try:
            self.clipboard_clear()
            self.clipboard_append(str(vals))
        except Exception:
            pass

    def copy_row(self):
        sel = self.selection()
        if not sel:
            return
        vals = self.item(sel[0], 'values')
        try:
            self.clipboard_clear()
            self.clipboard_append("\t".join(map(str, vals)))
        except Exception:
            pass

    def open_symbol_page(self):
        sel = self.selection()
        if not sel:
            return
        if "کد_داخلی" in self.df.columns:
            vals = self.item(sel[0], 'values')
            idx = list(self.df.columns).index("کد_داخلی")
            code = vals[idx]
            if code:
                webbrowser.open(f"https://www.tsetmc.com/instInfo/{code}")

    def filter_by_symbol_from_selection(self):
        sel = self.selection()
        if not sel:
            return
        vals = self.item(sel[0], 'values')
        cols = list(self.df.columns)
        if 'گروه_صنعت' not in cols or 'کد_بازار' not in cols:
            return
        industry_idx = cols.index('گروه_صنعت')
        market_idx = cols.index('کد_بازار')
        industry_val = vals[industry_idx]
        market_val = vals[market_idx]
        try:
            market_num = int(str(market_val))
        except:
            market_num = None
        def func(df):
            s_ind = df['گروه_صنعت'].astype(str).apply(normalize_text) == normalize_text(str(industry_val))
            def market_ok(x):
                try:
                    xi = int(str(x))
                except:
                    return False
                return xi in (300,303,309) or (market_num is not None and xi == market_num)
            mask_market = df['کد_بازار'].apply(market_ok)
            return df[s_ind & mask_market]
        desc = f"فیلتر نماد: گروه_صنعت={industry_val} و کد_بازار در [300,303,309] یا = {market_val}"
        payload = {'type':'pattern','column':'گروه_صنعت','mode':'contains','text':industry_val,'length':None,'exclude':False}
        self.add_filter_record(desc, func, enabled=True, persist_payload=payload)

    # ------------------------
    # Filter management
    # ------------------------
    def add_filter_record(self, desc, func, enabled=True, persist_payload=None, persist=True):
        """
        اضافه کردن فیلتر به لیست و در صورت نیاز ذخیرهٔ payload در settings_store
        """
        self.active_filters.append({'desc': desc, 'func': func, 'enabled': bool(enabled)})
        if persist and persist_payload is not None:
            settings_store.setdefault('saved_filters_full', [])
            settings_store['saved_filters_full'].append(persist_payload)
            save_settings(settings_store)
        self.apply_all_filters()

    def apply_all_filters(self):
        df = self.base_df.copy()
        for f in self.active_filters:
            if not f.get('enabled', True):
                continue
            try:
                df = f['func'](df)
            except Exception:
                pass
        self.df = df.reset_index(drop=True)
        if 'ردیف' in self.df.columns:
            self.df['ردیف'] = range(1, len(self.df) + 1)
        else:
            self.df.insert(0, 'ردیف', range(1, len(self.df) + 1))
        self.norm_df = self._build_normalized_df(self.df)
        widths = self._compute_optimal_widths()
        self.app_runtime_log.setdefault('column_widths', {}).update(widths)
        for col, w in widths.items():
            try:
                self.column(col, width=int(w))
            except Exception:
                pass
        self._load_batch()
        for cb in getattr(self, 'on_update_callbacks', []):
            try:
                cb()
            except Exception:
                pass

    def add_value_filter(self, column, values, exclude=False):
        if column not in self.base_df.columns:
            return
        norm_values = [normalize_text(v) for v in values]
        def func(df):
            s = df[column].astype(str).apply(normalize_text)
            mask = s.isin(norm_values)
            return df[~mask] if exclude else df[mask]
        desc = f"{column} {'شامل نشود' if exclude else 'شامل شود'}: {', '.join(values)}"
        payload = {'type': 'value', 'column': column, 'values': values, 'exclude': bool(exclude)}
        self.add_filter_record(desc, func, enabled=True, persist_payload=payload)

    def add_pattern_filter(self, column, mode, text, length=None, exclude=False):
        if column not in self.base_df.columns:
            return
        norm_text = normalize_text(text)
        def func(df):
            s = df[column].astype(str).apply(normalize_text)
            if mode == 'start':
                L = int(length) if length else len(norm_text)
                mask = s.str[:L] == norm_text
            elif mode == 'end':
                L = int(length) if length else len(norm_text)
                mask = s.str[-L:] == norm_text
            else:
                mask = s.str.contains(norm_text, na=False)
            return df[~mask] if exclude else df[mask]
        desc = f"{column} {'شامل نشود' if exclude else 'شامل شود'} الگو {mode}='{text}'"
        payload = {'type': 'pattern', 'column': column, 'mode': mode, 'text': text, 'length': length, 'exclude': bool(exclude)}
        self.add_filter_record(desc, func, enabled=True, persist_payload=payload)

    def add_relation_filter(self, left_col, op, right_expr):
        def func(df):
            L = pd.to_numeric(df[left_col], errors='coerce')
            expr = right_expr
            for c in df.columns:
                expr = re.sub(r'\b' + re.escape(c) + r'\b', f"df['{c}']", expr)
            try:
                R = pd.eval(expr, engine='python')
                R = pd.to_numeric(R, errors='coerce')
            except Exception:
                try:
                    R = float(right_expr)
                    R = pd.Series(R, index=df.index)
                except Exception:
                    R = pd.Series(pd.NA, index=df.index)
            if op == '>':
                mask = L > R
            elif op == '<':
                mask = L < R
            elif op == '>=':
                mask = L >= R
            elif op == '<=':
                mask = L <= R
            elif op == '==':
                mask = L == R
            elif op == '!=':
                mask = L != R
            else:
                mask = pd.Series(True, index=df.index)
            return df[mask.fillna(False)]
        desc = f"رابطه: {left_col} {op} {right_expr}"
        payload = {'type': 'relation', 'left': left_col, 'op': op, 'right': right_expr}
        self.add_filter_record(desc, func, enabled=True, persist_payload=payload)

    def clear_all_filters(self):
        self.active_filters.clear()
        settings_store['saved_filters_full'] = []
        save_settings(settings_store)
        self.apply_all_filters()

    # ------------------------
    # Search helper
    # ------------------------
    def search_live(self, term):
        for iid in self.get_children():
            self.item(iid, tags=())
        if not term:
            return []
        t = normalize_text(term)
        mask = pd.Series(False, index=self.norm_df.index)
        for col in self.norm_df.columns:
            try:
                mask = mask | self.norm_df[col].astype(str).str.contains(t, na=False)
            except Exception:
                continue
        matches = list(mask[mask].index)
        items = self.get_children()
        for i, iid in enumerate(items):
            if i in matches:
                self.item(iid, tags=('search_match',))
        return matches

    def export_current_view_to_csv(self, filepath):
        try:
            visible_cols = [c for c, v in self.visible_columns.items() if v]
            export_df = self.df.copy()
            for col in ('ارزش بازار همت', 'PE', 'صف خرید', 'صف فروش'):
                if col in self.base_df.columns and col in export_df.columns:
                    export_df[col] = self.base_df.loc[export_df.index, col]
            if visible_cols:
                export_df.to_csv(filepath, index=False, encoding='utf-8-sig', columns=visible_cols)
            else:
                export_df.to_csv(filepath, index=False, encoding='utf-8-sig')
            return True, None
        except Exception as e:
            return False, str(e)


# ------------------------
# BottomStatsTable
# ------------------------
# جایگزین کامل کلاس BottomStatsTable در core.py
class BottomStatsTable(ttk.Treeview):
    """
    جدول آمار پایین: جمع، میانگین، میانه، میانه مقاوم، کمترین، بیشترین
    با محافظت در برابر callback های after وقتی ویجت نابود شده باشد.
    """
    def __init__(self, parent, tree: AdvancedTreeview, visible_cols_for_bottom=None, **kwargs):
        cols = ['متریک'] + (visible_cols_for_bottom if visible_cols_for_bottom is not None else list(tree.df.columns))
        super().__init__(parent, columns=cols, show='headings', **kwargs)
        self.tree = tree
        self.visible_cols_for_bottom = cols[1:]
        for col in cols:
            if col == 'متریک':
                self.heading(col, text=col)
                self.column(col, width=120, anchor='w', stretch=False)
            else:
                display = COLUMN_NAME_MAP.get(col, COLUMN_NAME_MAP.get(col.lower(), col))
                self.heading(col, text=display)
                try:
                    w = self.tree.column(col, option='width') if col in self.tree['columns'] else 100
                except Exception:
                    w = 100
                self.column(col, width=w, anchor='center', stretch=False)
        self.metrics = ['جمع', 'میانگین', 'میانه', 'میانه مقاوم', 'کمترین', 'بیشترین']
        self._init_rows()
        self._after_id = None
        # وقتی ویجت نابود می‌شود، cleanup انجام شود
        self.bind("<Destroy>", self._on_destroy, add=True)

    def _init_rows(self):
        try:
            self.delete(*self.get_children())
        except Exception:
            pass
        for m in self.metrics:
            try:
                self.insert('', 'end', values=[m] + ['' for _ in range(len(self['columns']) - 1)])
            except Exception:
                pass

    def refresh_debounced(self, delay=300):
        # لغو هر after قبلی و زمان‌بندی جدید
        try:
            if self._after_id:
                self.after_cancel(self._after_id)
        except Exception:
            pass
        try:
            self._after_id = self.after(delay, self._compute_and_fill)
        except Exception:
            self._after_id = None

    def _compute_and_fill(self):
        # اگر ویجت دیگر وجود ندارد، کاری نکن
        try:
            if not getattr(self, 'winfo_exists', lambda: False)() or not self.winfo_exists():
                self._after_id = None
                return
        except Exception:
            # اگر هر خطایی در بررسی وجود ویجت رخ داد، ایمن عمل کن
            self._after_id = None
            return

        # محافظت در برابر خطاهای داخلی هنگام حذف/درج ردیف‌ها
        try:
            self.delete(*self.get_children())
        except Exception:
            # اگر ویجت حذف شده یا خطای Tcl رخ داد، فقط بازنشانی و خروج
            self._after_id = None
            return

        df = None
        try:
            df = self.tree.df
        except Exception:
            df = None

        if df is None or df.empty:
            self._init_rows()
            self._after_id = None
            return

        cols = self.visible_cols_for_bottom
        sums = []; means = []; medians = []; robusts = []; mins = []; maxs = []
        for col in cols:
            try:
                if col == 'ارزش معاملات به میلیارد تومن':
                    s = pd.to_numeric(df.get('ارزش_معاملات', pd.Series(dtype=float)), errors='coerce') / 1e10
                    s = s.dropna()
                elif col == 'ارزش بازار همت':
                    s = pd.to_numeric(df.get('ارزش بازار همت', pd.Series(dtype=float)), errors='coerce').dropna()
                else:
                    s = pd.to_numeric(df.get(col, pd.Series(dtype=float)), errors='coerce').dropna()
            except Exception:
                s = pd.Series(dtype=float)

            if s.empty:
                sums.append(''); means.append(''); medians.append(''); robusts.append(''); mins.append(''); maxs.append('')
                continue
            try:
                total = s.sum(); mean = s.mean(); median = s.median()
                q1, q3 = s.quantile(0.25), s.quantile(0.75)
                iqr = q3 - q1
                low, high = q1 - 1.5 * iqr, q3 + 1.5 * iqr
                trimmed = s[(s >= low) & (s <= high)]
                robust = trimmed.median() if not trimmed.empty else median
                mn = s.min(); mx = s.max()
            except Exception:
                total = mean = median = robust = mn = mx = pd.NA

            def fmt(x, colname):
                try:
                    if pd.isna(x):
                        return ''
                    if colname == 'ارزش معاملات به میلیارد تومن':
                        try:
                            return str(int(x))
                        except:
                            return str(x)
                    if colname == 'ارزش بازار همت':
                        try:
                            xv = float(x)
                            return f"{xv:.1f}"
                        except:
                            return str(x)
                    if abs(x - int(x)) < 1e-9:
                        return str(int(x))
                    return f"{x:.2f}"
                except Exception:
                    return str(x)

            sums.append(fmt(total, col)); means.append(fmt(mean, col)); medians.append(fmt(median, col)); robusts.append(fmt(robust, col))
            mins.append(fmt(mn, col)); maxs.append(fmt(mx, col))

        # درج ردیف‌ها با محافظت در برابر خطا
        try:
            for row_vals in [['جمع'] + sums, ['میانگین'] + means, ['میانه'] + medians, ['میانه مقاوم'] + robusts, ['کمترین'] + mins, ['بیشترین'] + maxs]:
                self.insert('', 'end', values=row_vals)
        except Exception:
            pass

        # تنظیم عرض ستون‌ها مطابق جدول اصلی (در صورت وجود)
        try:
            for col in cols:
                if col in self.tree['columns']:
                    w = self.tree.column(col, option='width')
                    try:
                        self.column(col, width=w)
                    except Exception:
                        pass
        except Exception:
            pass

        # پاکسازی شناسه after
        self._after_id = None

    def _on_destroy(self, event=None):
        # وقتی ویجت نابود می‌شود، هر after زمان‌بندی‌شده را لغو کن
        try:
            if self._after_id:
                try:
                    self.after_cancel(self._after_id)
                except Exception:
                    pass
                self._after_id = None
        except Exception:
            pass



# ------------------------
# ScrollableToplevel helper
# ------------------------
class ScrollableToplevel(tk.Toplevel):
    """
    پنجرهٔ Toplevel با قابلیت اسکرول عمودی و افقی برای محتوای داخلی.
    استفاده برای دیالوگ‌های بزرگ.
    """
    def __init__(self, parent, title="", width=900, height=600):
        super().__init__(parent)
        self.title(title)
        self.geometry(f"{width}x{height}")
        self.container = ttk.Frame(self)
        self.container.pack(fill=tk.BOTH, expand=True)
        self.canvas = tk.Canvas(self.container)
        self.v_scroll = ttk.Scrollbar(self.container, orient="vertical", command=self.canvas.yview)
        self.h_scroll = ttk.Scrollbar(self.container, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=self.v_scroll.set, xscrollcommand=self.h_scroll.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        self.inner = ttk.Frame(self.canvas)
        self.canvas_window = self.canvas.create_window((0,0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        # mouse wheel bindings
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel_windows)
        self.canvas.bind_all("<Button-4>", self._on_mousewheel_unix)
        self.canvas.bind_all("<Button-5>", self._on_mousewheel_unix)

    def _on_mousewheel_windows(self, event):
        try:
            self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        except Exception:
            pass

    def _on_mousewheel_unix(self, event):
        if event.num == 4:
            self.canvas.yview_scroll(-3, "units")
        elif event.num == 5:
            self.canvas.yview_scroll(3, "units")


# ------------------------
# ColumnSettingsDialog (فیلترها و تغییر نام ستون)
# ------------------------
class ColumnSettingsDialog(ScrollableToplevel):
    """
    دیالوگ مدیریت فیلترها، انتخاب ستون‌ها، اعمال فیلترهای الگو/مقدار/رابطه‌ای و تغییر نام ستون.
    این دیالوگ از AdvancedTreeview استفاده می‌کند و فیلترها را به صورت persistable ذخیره می‌کند.
    """
    def __init__(self, parent, tree: AdvancedTreeview):
        super().__init__(parent, title="فیلترها", width=1400, height=900)
        self.parent = parent
        self.tree = tree
        self.value_vars = []
        self.mode_var = tk.StringVar(value="include")
        self.pattern_text = tk.StringVar()
        self.pattern_mode = tk.StringVar(value="contains")
        self.pattern_length = tk.StringVar(value="")
        self.selected_column = None
        self.sort_mode = tk.StringVar(value='freq')
        self._build_ui()
        self._refresh_filters_list()

    def _bind_copy_paste(self, entry):
        entry.bind("<Control-c>", lambda e: entry.event_generate("<<Copy>>"))
        entry.bind("<Control-x>", lambda e: entry.event_generate("<<Cut>>"))
        entry.bind("<Control-v>", lambda e: entry.event_generate("<<Paste>>"))

    def _build_ui(self):
        frame = self.inner
        left = ttk.Frame(frame); left.grid(row=0, column=0, sticky='ns', padx=8, pady=8)
        mid = ttk.Frame(frame); mid.grid(row=0, column=1, sticky='nsew', padx=8, pady=8)
        right = ttk.Frame(frame); right.grid(row=0, column=2, sticky='ns', padx=8, pady=8)
        frame.grid_columnconfigure(1, weight=1)
        frame.grid_rowconfigure(0, weight=1)

        ttk.Label(left, text="ستون‌ها", font=("Tahoma", 11, "bold")).pack(anchor='w')
        self.col_listbox = tk.Listbox(left, exportselection=False, height=40, width=48)
        self.col_listbox.pack(fill='y', expand=True)
        self.col_index_to_key = []
        for c in self.tree.df.columns:
            display = COLUMN_NAME_MAP.get(c, COLUMN_NAME_MAP.get(c.lower(), c))
            self.col_listbox.insert(tk.END, f"{display}  [{c}]")
            self.col_index_to_key.append(c)
        self.col_listbox.bind("<<ListboxSelect>>", self.on_col_select)

        ttk.Label(mid, text="مقادیر ستون (فراوانی)", font=("Tahoma", 11, "bold")).pack(anchor='w')
        val_container = ttk.Frame(mid)
        val_container.pack(fill='both', expand=True)
        self.val_canvas = tk.Canvas(val_container, highlightthickness=0)
        self.val_vscroll = ttk.Scrollbar(val_container, orient='vertical', command=self.val_canvas.yview)
        self.val_canvas.configure(yscrollcommand=self.val_vscroll.set)
        self.val_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.val_vscroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.val_inner = ttk.Frame(self.val_canvas)
        self.val_canvas.create_window((0,0), window=self.val_inner, anchor='nw')
        self.val_inner.bind("<Configure>", lambda e: self.val_canvas.configure(scrollregion=self.val_canvas.bbox("all")))
        self.val_canvas.bind("<Enter>", lambda e: self._bind_mousewheel(self.val_canvas))
        self.val_canvas.bind("<Leave>", lambda e: self._unbind_mousewheel(self.val_canvas))

        btn_frame = ttk.Frame(mid); btn_frame.pack(fill='x', pady=6)
        ttk.Radiobutton(btn_frame, text="Include", variable=self.mode_var, value="include").pack(side='left', padx=6)
        ttk.Radiobutton(btn_frame, text="Exclude", variable=self.mode_var, value="exclude").pack(side='left', padx=6)
        ttk.Button(btn_frame, text="اعمال فیلتر انتخاب‌شده", command=self.apply_selected_values).pack(side='right', padx=6)

        sort_frame = ttk.Frame(mid); sort_frame.pack(fill='x', pady=(6,0))
        self.sort_btn = ttk.Button(sort_frame, text="مرتب‌سازی: بر اساس مقدار", command=self.toggle_sort_mode)
        self.sort_btn.pack(side='left', padx=4)

        ttk.Label(right, text="فیلتر الگو", font=("Tahoma", 11, "bold")).pack(anchor='w')
        ttk.Label(right, text="الگو:").pack(anchor='w', pady=(6,0))
        e1 = ttk.Entry(right, textvariable=self.pattern_text, width=30); e1.pack(anchor='w', pady=4); self._bind_copy_paste(e1)
        ttk.Combobox(right, values=["start","end","contains"], textvariable=self.pattern_mode, state='readonly', width=12).pack(anchor='w', pady=4)
        ttk.Label(right, text="طول اختیاری:").pack(anchor='w')
        e2 = ttk.Entry(right, textvariable=self.pattern_length, width=8); e2.pack(anchor='w', pady=4); self._bind_copy_paste(e2)
        ttk.Button(right, text="اعمال فیلتر الگو", command=self.apply_pattern_filter).pack(anchor='w', pady=6)

        ttk.Separator(right, orient='horizontal').pack(fill='x', pady=8)
        ttk.Label(right, text="فیلتر رابطه‌ای", font=("Tahoma", 11, "bold")).pack(anchor='w')
        ttk.Label(right, text="مثال: 2 * قیمت_پایانی یا قیمت_دیروز یا 100000").pack(anchor='w', pady=(4,0))
        self.left_col_entry = ttk.Entry(right, width=30); self.left_col_entry.pack(anchor='w', pady=4); self._bind_copy_paste(self.left_col_entry)
        self.op_entry = ttk.Combobox(right, values=['>','<','>=','<=','==','!='], state='readonly', width=6); self.op_entry.pack(anchor='w', pady=4)
        self.right_expr_entry = ttk.Entry(right, width=30); self.right_expr_entry.pack(anchor='w', pady=4); self._bind_copy_paste(self.right_expr_entry)
        ttk.Button(right, text="اعمال فیلتر رابطه‌ای", command=self.apply_relation_filter).pack(anchor='w', pady=6)

        ttk.Separator(right, orient='horizontal').pack(fill='x', pady=8)
        ttk.Label(right, text="تغییر نام ستون", font=("Tahoma", 11, "bold")).pack(anchor='w')
        self.rename_from = ttk.Entry(right, width=30); self.rename_from.pack(anchor='w', pady=4); self._bind_copy_paste(self.rename_from)
        self.rename_to = ttk.Entry(right, width=30); self.rename_to.pack(anchor='w', pady=4); self._bind_copy_paste(self.rename_to)
        ttk.Button(right, text="اعمال تغییر نام", command=self.apply_rename).pack(anchor='w', pady=6)

        ttk.Separator(right, orient='horizontal').pack(fill='x', pady=8)
        ttk.Label(right, text="فیلترهای اعمال‌شده", font=("Tahoma", 11, "bold")).pack(anchor='w', pady=(6,2))
        filters_container = ttk.Frame(right)
        filters_container.pack(fill='both', expand=True)
        self.filters_canvas = tk.Canvas(filters_container, height=200, highlightthickness=0)
        self.filters_vscroll = ttk.Scrollbar(filters_container, orient='vertical', command=self.filters_canvas.yview)
        self.filters_canvas.configure(yscrollcommand=self.filters_vscroll.set)
        self.filters_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.filters_vinner = ttk.Frame(self.filters_canvas)
        self.filters_canvas.create_window((0,0), window=self.filters_vinner, anchor='nw')
        self.filters_vinner.bind("<Configure>", lambda e: self.filters_canvas.configure(scrollregion=self.filters_canvas.bbox("all")))
        self.filters_canvas.bind("<Enter>", lambda e: self._bind_mousewheel(self.filters_canvas))
        self.filters_canvas.bind("<Leave>", lambda e: self._unbind_mousewheel(self.filters_canvas))

        bottom_buttons = ttk.Frame(frame)
        bottom_buttons.grid(row=1, column=0, columnspan=3, sticky='ew', padx=8, pady=8)
        ttk.Button(bottom_buttons, text="پاک کردن همه فیلترها", command=self._clear_all_filters).pack(side='left', padx=6)
        ttk.Button(bottom_buttons, text="بستن", command=self.destroy).pack(side='right', padx=6)

    def _bind_mousewheel(self, widget):
        widget.bind_all("<MouseWheel>", lambda e, w=widget: self._on_mousewheel(e, w))
        widget.bind_all("<Button-4>", lambda e, w=widget: self._on_mousewheel(e, w))
        widget.bind_all("<Button-5>", lambda e, w=widget: self._on_mousewheel(e, w))

    def _unbind_mousewheel(self, widget):
        widget.unbind_all("<MouseWheel>")
        widget.unbind_all("<Button-4>")
        widget.unbind_all("<Button-5>")

    def _on_mousewheel(self, event, widget):
        try:
            if hasattr(event, 'num') and event.num in (4,5):
                if event.num == 4:
                    widget.yview_scroll(-3, "units")
                else:
                    widget.yview_scroll(3, "units")
            else:
                delta = int(-1 * (event.delta / 120))
                widget.yview_scroll(delta * 3, "units")
        except Exception:
            pass

    def toggle_sort_mode(self):
        if self.sort_mode.get() == 'freq':
            self.sort_mode.set('value')
            self.sort_btn.config(text="مرتب‌سازی: بر اساس مقدار")
        else:
            self.sort_mode.set('freq')
            self.sort_btn.config(text="مرتب‌سازی: بر اساس فراوانی")
        self._rebuild_values_list()

    def on_col_select(self, _=None):
        sel = self.col_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        col = self.col_index_to_key[idx]
        self.selected_column = col
        self._rebuild_values_list()
        self.rename_from.delete(0, tk.END)
        self.rename_from.insert(0, col)

    def _rebuild_values_list(self):
        for w in self.val_inner.winfo_children():
            w.destroy()
        self.value_vars.clear()
        col = self.selected_column
        if not col or col not in self.tree.df.columns:
            return
        counts = self.tree.df[col].astype(str).map(normalize_text).value_counts(dropna=False)
        items = list(counts.items())
        if self.sort_mode.get() == 'value':
            def key_fn(x):
                v = x[0]
                try:
                    return (0, float(v))
                except:
                    return (1, str(v))
            items.sort(key=key_fn)
        else:
            items.sort(key=lambda x: (-x[1], str(x[0])))
        default_font = tkfont.nametofont("TkDefaultFont")
        bold_font = default_font.copy(); bold_font.configure(weight="bold")
        for val, cnt in items:
            label = "" if pd.isna(val) else str(val)
            display_label = label
            if col == 'کد_بازار' or col.lower() == 'کد_بازار':
                lbl = MARKET_LABELS.get(str(label), '')
                display_label = f"{label} {lbl}" if lbl else label
            var = tk.BooleanVar(value=False)
            if col == 'کد_بازار' and str(label) in ('300', '303', '309', '313'):
                cb_widget = tk.Checkbutton(self.val_inner, text=f"{display_label} ({cnt})", variable=var, font=bold_font, anchor='w')
                cb_widget.pack(anchor='w', fill='x', padx=4, pady=1)
                self.value_vars.append((label, var))
            else:
                cb = ttk.Checkbutton(self.val_inner, text=f"{display_label} ({cnt})", variable=var)
                cb.pack(anchor='w', fill='x', padx=4, pady=1)
                self.value_vars.append((label, var))

    def apply_selected_values(self):
        if not self.selected_column:
            return
        chosen = [val for (val, var) in self.value_vars if var.get()]
        if not chosen:
            return
        exclude = (self.mode_var.get() == "exclude")
        col = self.selected_column
        self.tree.add_value_filter(col, chosen, exclude=exclude)
        self._refresh_filters_list()

    def apply_pattern_filter(self):
        if not self.selected_column:
            return
        text = self.pattern_text.get().strip()
        if not text:
            return
        mode = self.pattern_mode.get()
        length = self.pattern_length.get().strip()
        L = int(length) if length.isdigit() else None
        exclude = (self.mode_var.get() == "exclude")
        col = self.selected_column
        self.tree.add_pattern_filter(col, mode, text, length=L, exclude=exclude)
        self._refresh_filters_list()

    def apply_relation_filter(self):
        left = self.left_col_entry.get().strip()
        op = self.op_entry.get().strip()
        right = self.right_expr_entry.get().strip()
        if not left or not op or not right:
            return
        self.tree.add_relation_filter(left, op, right)
        self._refresh_filters_list()

    def apply_rename(self):
        frm = self.rename_from.get().strip()
        to = self.rename_to.get().strip()
        if not frm or not to:
            return
        if frm in self.tree.base_df.columns:
            try:
                # rename in base and current df
                self.tree.base_df.rename(columns={frm: to}, inplace=True)
                self.tree.df.rename(columns={frm: to}, inplace=True)
            except Exception:
                pass
            # transfer visibility flag if present
            if frm in self.tree.visible_columns:
                self.tree.visible_columns[to] = self.tree.visible_columns.pop(frm)
            # update COLUMN_NAME_MAP and persist
            COLUMN_NAME_MAP[frm] = to
            settings_store['column_name_map'] = COLUMN_NAME_MAP
            save_settings(settings_store)
            # rebuild column listbox
            self.col_listbox.delete(0, tk.END)
            self.col_index_to_key.clear()
            for c in self.tree.df.columns:
                display = COLUMN_NAME_MAP.get(c, COLUMN_NAME_MAP.get(c.lower(), c))
                self.col_listbox.insert(tk.END, f"{display}  [{c}]")
                self.col_index_to_key.append(c)
            # refresh values and tree
            self._rebuild_values_list()
            try:
                self.tree._setup_columns(auto_optimize=False)
                self.tree._load_batch()
            except Exception:
                pass

    def _refresh_filters_list(self):
        for w in self.filters_vinner.winfo_children():
            w.destroy()
        self.filter_widgets = []
        for idx, f in enumerate(self.tree.active_filters):
            frame = ttk.Frame(self.filters_vinner)
            frame.pack(fill='x', pady=2, padx=2)
            var = tk.BooleanVar(value=f.get('enabled', True))
            chk = ttk.Checkbutton(frame, variable=var, command=lambda i=idx, v=var: self._toggle_filter(i, v.get()))
            chk.pack(side='left')
            lbl = ttk.Label(frame, text=f.get('desc', '')[:80], anchor='w')
            lbl.pack(side='left', fill='x', expand=True, padx=6)
            btn = ttk.Button(frame, text="حذف", width=6, command=lambda i=idx: self._remove_filter(i))
            btn.pack(side='right', padx=4)
            self.filter_widgets.append((frame, var, lbl, btn))
        self.filters_canvas.configure(scrollregion=self.filters_canvas.bbox("all"))

    def _toggle_filter(self, index, enabled):
        if 0 <= index < len(self.tree.active_filters):
            self.tree.active_filters[index]['enabled'] = bool(enabled)
            self.tree.apply_all_filters()
            self._refresh_filters_list()

    def _remove_filter(self, index):
        if 0 <= index < len(self.tree.active_filters):
            del self.tree.active_filters[index]
            # also remove from persisted list by index (best-effort)
            if settings_store.get('saved_filters_full'):
                try:
                    del settings_store['saved_filters_full'][index]
                    save_settings(settings_store)
                except Exception:
                    pass
            self.tree.apply_all_filters()
            self._refresh_filters_list()

    def _clear_all_filters(self):
        self.tree.clear_all_filters()
        self._refresh_filters_list()


# ------------------------
# AppSettingsDialog (تنظیمات برنامه)
# ------------------------
class AppSettingsDialog(ScrollableToplevel):
    def __init__(self, parent, app, tree: AdvancedTreeview = None):
        super().__init__(parent, title="تنظیمات برنامه", width=900, height=560)
        self.app = app
        self.tree = tree
        self.pending_main = {}
        self.pending_bottom = {}
        self._build_ui()

    def _build_ui(self):
        frame = self.inner
        ttk.Label(frame, text="آدرس دانلود داده (URL):").grid(row=0, column=0, columnspan=2, sticky='w', padx=8, pady=(8,4))
        self.url_var = tk.StringVar(value=getattr(self.app, 'data_url', URL_DEFAULT))
        eurl = ttk.Entry(frame, textvariable=self.url_var, width=80)
        eurl.grid(row=1, column=0, columnspan=2, sticky='ew', padx=8)
        eurl.bind("<Control-c>", lambda e: eurl.event_generate("<<Copy>>"))
        eurl.bind("<Control-v>", lambda e: eurl.event_generate("<<Paste>>"))

        ttk.Label(frame, text="ستون‌های جدول اصلی", font=("Tahoma", 11, "bold")).grid(row=2, column=0, sticky='w', padx=8, pady=(10,4))
        ttk.Label(frame, text="ستون‌های جدول پایین (آمار)", font=("Tahoma", 11, "bold")).grid(row=2, column=1, sticky='w', padx=8, pady=(10,4))

        cb_container1 = ttk.Frame(frame)
        cb_container1.grid(row=3, column=0, sticky='nsew', padx=8, pady=4)
        cb_canvas1 = tk.Canvas(cb_container1, height=320)
        cb_vscroll1 = ttk.Scrollbar(cb_container1, orient="vertical", command=cb_canvas1.yview)
        cb_canvas1.configure(yscrollcommand=cb_vscroll1.set)
        cb_canvas1.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        cb_vscroll1.pack(side=tk.RIGHT, fill=tk.Y)
        cb_inner1 = ttk.Frame(cb_canvas1)
        cb_canvas1.create_window((0,0), window=cb_inner1, anchor="nw")
        cb_inner1.bind("<Configure>", lambda e: cb_canvas1.configure(scrollregion=cb_canvas1.bbox("all")))

        cb_container2 = ttk.Frame(frame)
        cb_container2.grid(row=3, column=1, sticky='nsew', padx=8, pady=4)
        cb_canvas2 = tk.Canvas(cb_container2, height=320)
        cb_vscroll2 = ttk.Scrollbar(cb_container2, orient="vertical", command=cb_canvas2.yview)
        cb_canvas2.configure(yscrollcommand=cb_vscroll2.set)
        cb_canvas2.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        cb_vscroll2.pack(side=tk.RIGHT, fill=tk.Y)
        cb_inner2 = ttk.Frame(cb_canvas2)
        cb_canvas2.create_window((0,0), window=cb_inner2, anchor="nw")
        cb_inner2.bind("<Configure>", lambda e: cb_canvas2.configure(scrollregion=cb_canvas2.bbox("all")))

        self.col_vars_main = {}
        self.col_vars_bottom = {}
        cols = list(self.app.current_tree.df.columns) if getattr(self.app, 'current_tree', None) else []
        for col in cols:
            var = tk.BooleanVar(value=self.app.current_tree.visible_columns.get(col, True) if getattr(self.app, 'current_tree', None) else True)
            cb = ttk.Checkbutton(cb_inner1, text=COLUMN_NAME_MAP.get(col, COLUMN_NAME_MAP.get(col.lower(), col)), variable=var)
            cb.pack(anchor='w', padx=4, pady=2)
            self.col_vars_main[col] = var

        saved_bottom = settings_store.get('bottom_visible_columns')
        for col in cols:
            default_bottom = True
            if saved_bottom is not None:
                default_bottom = col in saved_bottom
            var2 = tk.BooleanVar(value=default_bottom)
            cb2 = ttk.Checkbutton(cb_inner2, text=COLUMN_NAME_MAP.get(col, COLUMN_NAME_MAP.get(col.lower(), col)), variable=var2)
            cb2.pack(anchor='w', padx=4, pady=2)
            self.col_vars_bottom[col] = var2

        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=4, column=0, columnspan=2, sticky='ew', padx=8, pady=8)
        ttk.Button(btn_frame, text="اعمال تغییرات نمایش/مخفی‌سازی", command=self.apply_visibility_changes).pack(side='left', padx=6)
        ttk.Button(btn_frame, text="ذخیره URL", command=self.save_url).pack(side='right', padx=6)
        ttk.Button(btn_frame, text="بستن", command=self.destroy).pack(side='right', padx=6)

    def apply_visibility_changes(self):
        if getattr(self.app, 'current_tree', None):
            for col, var in self.col_vars_main.items():
                self.app.current_tree.visible_columns[col] = bool(var.get())
            display_cols = [c for c, v in self.app.current_tree.visible_columns.items() if v]
            try:
                self.app.current_tree.configure(displaycolumns=display_cols)
                self.app.current_tree._load_batch()
            except Exception:
                pass

        selected_bottom = [c for c, var in self.col_vars_bottom.items() if var.get()]
        if not selected_bottom and getattr(self.app, 'current_tree', None):
            selected_bottom = list(self.app.current_tree.df.columns)
        settings_store['bottom_visible_columns'] = selected_bottom
        save_settings(settings_store)

        # rebuild bottom stats if present
        if getattr(self.app, 'bottom_frame', None):
            try:
                self.app.bottom_frame.destroy()
            except Exception:
                pass
            self.app.bottom_frame = None
            self.app.bottom_stats = None
        if getattr(self.app, 'current_tree', None):
            self.app.bottom_frame = ttk.Frame(self.app.root)
            self.app.bottom_frame.pack(fill='x', padx=8, pady=(0,8))
            self.app.bottom_stats = BottomStatsTable(self.app.bottom_frame, self.app.current_tree, visible_cols_for_bottom=selected_bottom)
            self.app.bottom_stats.pack(fill='x')
            if hasattr(self.app.current_tree, 'on_update_callbacks'):
                if self.app.bottom_stats.refresh_debounced not in self.app.current_tree.on_update_callbacks:
                    self.app.current_tree.on_update_callbacks.append(self.app.bottom_stats.refresh_debounced)
            self.app.bottom_stats.refresh_debounced()

    def save_url(self):
        new = self.url_var.get().strip()
        if new:
            self.app.data_url = new
            settings_store['data_url'] = new
            save_settings(settings_store)


# پایان بخش دوم
