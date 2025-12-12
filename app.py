# app.py
# Run: python app.py
# Requirements: pandas, requests
# pip install pandas requests

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading, time, json
from core import (
    settings_store, save_settings, URL_DEFAULT, DEFAULT_EXPORT_NAME, FIELD_MAPPING,
    fetch_sections, parse_section, merge_section3_into2, AdvancedTreeview,
    BottomStatsTable, ColumnSettingsDialog, AppSettingsDialog
)
from client_type_export import ClientTypeExportWindow

class MarketApp:
    def __init__(self, root):
        self.root = root
        self.root.title("TSETMC Viewer")
        self.data_url = settings_store.get('data_url', URL_DEFAULT)
        self.runtime_log = settings_store.get('runtime_log', {'column_widths': {}, 'converted_numeric_columns': [], 'visible_columns': {}})

        toolbar = ttk.Frame(root); toolbar.pack(fill=tk.X, padx=8, pady=6)
        btn_style = {'padx': 8, 'pady': 6}
        ttk.Button(toolbar, text="بارگذاری/به‌روزرسانی", command=self.load_sections_thread).pack(side=tk.LEFT, **btn_style)
        ttk.Button(toolbar, text="فیلترها", command=self.open_filters).pack(side=tk.LEFT, **btn_style)
        ttk.Button(toolbar, text="اعمال فیلتر ویژه", command=self.apply_special_filters).pack(side=tk.LEFT, **btn_style)
        ttk.Button(toolbar, text="خروجی CSV", command=self.export_current_view).pack(side=tk.LEFT, **btn_style)
        ttk.Button(toolbar, text="خروجی لاگ", command=self.export_log).pack(side=tk.LEFT, **btn_style)
        ttk.Button(toolbar, text="تنظیمات برنامه", command=self.open_app_settings).pack(side=tk.LEFT, **btn_style)
        ttk.Button(toolbar, text="خروجی حقیقی/حقوقی نماد", command=self.open_client_type_export).pack(side=tk.LEFT, **btn_style)

        search_frame = ttk.Frame(root); search_frame.pack(fill=tk.X, padx=8, pady=(0,6))
        ttk.Label(search_frame, text="جستجو:").pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=40); self.search_entry.pack(side=tk.LEFT, padx=6)
        self.search_entry.bind("<KeyRelease>", self.on_search_change_debounced)
        ttk.Button(search_frame, text="Next", command=self.search_next).pack(side=tk.LEFT, padx=6)
        self.search_count_label = ttk.Label(search_frame, text="(0)"); self.search_count_label.pack(side=tk.LEFT)

        self.notebook = ttk.Notebook(root); self.notebook.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        self.trees = []
        self.bottom_frame = None
        self.bottom_stats = None
        self._search_after_id = None
        self.current_tree = None
        self._load_thread = None

        self.load_sections_thread()

    def load_sections_thread(self):
        if self._load_thread and self._load_thread.is_alive():
            return
        self._load_thread = threading.Thread(target=self._load_sections_safe, daemon=True)
        self._load_thread.start()

    def _load_sections_safe(self):
        try:
            start = time.time()
            sections = fetch_sections(self.data_url)
            self.runtime_log['last_fetch_time'] = time.strftime("%Y-%m-%d %H:%M:%S")
            self.runtime_log['load_duration'] = round(time.time() - start, 3)
            self.root.after(0, lambda: self._populate_tabs(sections))
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("خطا در دریافت داده", str(e)))

    def _populate_tabs(self, sections):
        try:
            for tab in self.notebook.tabs():
                self.notebook.forget(tab)
            self.trees.clear()
            for i, sec in enumerate(sections):
                if not sec.strip(): continue
                if i == 2:
                    df2 = parse_section(sec, FIELD_MAPPING)
                    df3 = parse_section(sections[3], None) if len(sections) > 3 else None
                    df = merge_section3_into2(df2, df3 if df3 is not None else None)
                else:
                    df = parse_section(sec, None)
                frame = ttk.Frame(self.notebook)
                self.notebook.add(frame, text=f"بخش {i}")
                vscroll = ttk.Scrollbar(frame, orient="vertical")
                hscroll = ttk.Scrollbar(frame, orient="horizontal")
                tree = AdvancedTreeview(frame, df, app_runtime_log=self.runtime_log, yscrollcommand=vscroll.set, xscrollcommand=hscroll.set)
                tree.grid(row=0, column=0, sticky="nsew")
                vscroll.config(command=tree.yview); vscroll.grid(row=0, column=1, sticky="ns")
                hscroll.config(command=tree.xview); hscroll.grid(row=1, column=0, sticky="ew")
                frame.grid_rowconfigure(0, weight=1); frame.grid_columnconfigure(0, weight=1)
                self.trees.append(tree)
            # reapply persisted filters to last tree
            if self.trees:
                persisted = settings_store.get('saved_filters_full', [])
                if persisted:
                    try:
                        self.trees[-1]._reapply_persisted_filters(persisted)
                        self.trees[-1].apply_all_filters()
                    except Exception:
                        pass
            for idx, tab_id in enumerate(self.notebook.tabs()):
                if self.notebook.tab(tab_id, option='text') == "بخش 2":
                    self.notebook.select(idx)
                    break
            self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)
            self.on_tab_changed()
        except Exception as e:
            messagebox.showerror("خطا در ساخت تب‌ها", str(e))

    def on_tab_changed(self, _=None):
        idx = self.notebook.index(self.notebook.select())
        self.current_tree = self.trees[idx] if 0 <= idx < len(self.trees) else None
        self._attach_bottom_stats()

    def _attach_bottom_stats(self):
        if self.bottom_frame:
            try: self.bottom_frame.destroy()
            except: pass
            self.bottom_frame = None; self.bottom_stats = None

        self.bottom_frame = ttk.Frame(self.root); self.bottom_frame.pack(fill='x', padx=8, pady=(0,8))
        if not self.current_tree: return

        bottom_cols = settings_store.get('bottom_visible_columns')
        if bottom_cols is None:
            bottom_cols = list(self.current_tree.df.columns)
        if 'ارزش معاملات به میلیارد تومن' not in bottom_cols:
            bottom_cols += ['ارزش معاملات به میلیارد تومن']

        self.bottom_stats = BottomStatsTable(self.bottom_frame, self.current_tree, visible_cols_for_bottom=bottom_cols)
        self.bottom_stats.pack(fill='x')
        if hasattr(self.current_tree, 'on_update_callbacks'):
            if self.bottom_stats.refresh_debounced not in self.current_tree.on_update_callbacks:
                self.current_tree.on_update_callbacks.append(self.bottom_stats.refresh_debounced)
        self.bottom_stats.refresh_debounced()

    def on_search_change_debounced(self, _event):
        if self._search_after_id:
            self.root.after_cancel(self._search_after_id)
        self._search_after_id = self.root.after(300, self._on_search_change)

    def _on_search_change(self):
        term = self.search_var.get().strip()
        if not self.current_tree: return
        matches = self.current_tree.search_live(term)
        self.search_count_label.config(text=f"({len(matches)})")
        if self.bottom_stats: self.bottom_stats.refresh_debounced()

    def search_next(self):
        if not self.current_tree: return
        tree = self.current_tree
        items = tree.get_children()
        matches = [i for i in items if 'search_match' in tree.item(i, 'tags')]
        if not matches: return
        sel = tree.selection()
        cur = matches.index(sel[0]) if sel and sel[0] in matches else -1
        nxt = matches[(cur + 1) % len(matches)]
        tree.selection_set(nxt); tree.see(nxt); tree.focus(nxt)

    def open_filters(self):
        if not self.current_tree:
            messagebox.showwarning("هشدار", "ابتدا داده‌ها را بارگذاری کنید")
            return
        ColumnSettingsDialog(self.root, self.current_tree)

    def open_app_settings(self):
        AppSettingsDialog(self.root, self, tree=self.current_tree)

    def export_current_view(self):
        if not self.current_tree:
            messagebox.showwarning("هشدار", "تب فعالی وجود ندارد"); return
        filepath = filedialog.acksaveasfilename(defaultextension=".csv", initialfile=DEFAULT_EXPORT_NAME,
                                                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if not filepath: return
        ok, err = self.current_tree.export_current_view_to_csv(filepath)
        if ok: messagebox.showinfo("موفق", f"خروجی ذخیره شد:\n{filepath}")
        else: messagebox.showerror("خطا در ذخیره", err)

    def export_log(self):
        if not self.current_tree:
            messagebox.showwarning("هشدار", "تب فعالی وجود ندارد"); return
        cols = list(self.current_tree.df.columns)
        widths = {col: self.current_tree.column(col, option='width') for col in cols}
        self.runtime_log['column_widths'] = widths
        self.runtime_log['rows_shown'] = len(self.current_tree.df)
        self.runtime_log['filters_count'] = len(self.current_tree.active_filters)
        self.runtime_log['filters'] = [f['desc'] for f in self.current_tree.active_filters]
        self.runtime_log['visible_columns'] = self.current_tree.visible_columns.copy()
        filepath = filedialog.asksaveasfilename(defaultextension=".json", initialfile="tsetmc_log.json",
                                                filetypes=[("JSON files", "*.json"), ("All files", "*.*")])
        if not filepath: return
        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(self.runtime_log, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("موفق", f"لاگ ذخیره شد:\n{filepath}")
        except Exception as e:
            messagebox.showerror("خطا در ذخیره لاگ", str(e))

    def apply_special_filters(self):
        if not self.current_tree:
            messagebox.showwarning("هشدار", "ابتدا داده‌ها را بارگذاری کنید")
            return
        def filter_market(df):
            return df[df['کد_بازار'].astype(str).isin(['300','303','309','313'])]
        def filter_international(df):
            return df[df['کد_بین_المللی'].astype(str).str[-4:] == '0001']
        self.current_tree.add_filter_record('کد بازار در [300,303,309,313]', filter_market, enabled=True,
                                            persist_payload={'type':'value','column':'کد_بازار','values':['300','303','309','313'],'exclude':False})
        self.current_tree.add_filter_record('کد بین المللی انتهای 0001', filter_international, enabled=True,
                                            persist_payload={'type':'pattern','column':'کد_بین_المللی','mode':'end','text':'0001','length':4,'exclude':False})

    def open_client_type_export(self):
        if not self.current_tree:
            messagebox.showwarning("هشدار", "ابتدا داده‌ها را بارگذاری کنید"); return
        sel = self.current_tree.selection()
        if not sel:
            messagebox.showinfo("انتخاب نماد", "یک ردیف نماد را انتخاب کنید")
            return
        ClientTypeExportWindow(self.root, self.current_tree, selection_iid=sel[0])

def run():
    root = tk.Tk()
    root.geometry("1250x820")
    app = MarketApp(root)
    def on_close():
        try:
            settings_store['data_url'] = app.data_url
            settings_store['runtime_log'] = app.runtime_log
            save_settings(settings_store)
        except Exception:
            pass
        try: root.destroy()
        except Exception: pass
    root.protocol("WM_DELETE_WINDOW", on_close)
    root.mainloop()

if __name__ == "__main__":
    run()

# --- IDE variable snapshot ---
try:
    import json
    _vars_snapshot = {
        k: v for k, v in globals().items()
        if k not in ("__name__", "__file__", "__package__", "__loader__", "__spec__", "__builtins__")
        and isinstance(v, (int, float, str, list, dict, tuple, bool))
    }
    with open(r"E:/python bours3\vars_snapshot.json", "w", encoding="utf-8") as _f:
        json.dump(_vars_snapshot, _f, ensure_ascii=False, indent=2)
except Exception as _e:
    print("خطا در ذخیره متغیرها:", _e)

# --- IDE variable snapshot ---
try:
    import json
    _vars_snapshot = {
        k: v for k, v in globals().items()
        if k not in ("__name__", "__file__", "__package__", "__loader__", "__spec__", "__builtins__")
        and isinstance(v, (int, float, str, list, dict, tuple, bool))
    }
    with open(r"E:/python bours3\vars_snapshot.json", "w", encoding="utf-8") as _f:
        json.dump(_vars_snapshot, _f, ensure_ascii=False, indent=2)
except Exception as _e:
    print("خطا در ذخیره متغیرها:", _e)

# --- IDE variable snapshot ---
try:
    import json
    _vars_snapshot = {
        k: v for k, v in globals().items()
        if k not in ("__name__", "__file__", "__package__", "__loader__", "__spec__", "__builtins__")
        and isinstance(v, (int, float, str, list, dict, tuple, bool))
    }
    with open(r"E:/python bours3\vars_snapshot.json", "w", encoding="utf-8") as _f:
        json.dump(_vars_snapshot, _f, ensure_ascii=False, indent=2)
except Exception as _e:
    print("خطا در ذخیره متغیرها:", _e)
