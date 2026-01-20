# relic_score_ui_multi_json.py
# -*- coding: utf-8 -*-

import json
import os
import re
from typing import Dict, Optional, List, Tuple

import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import matplotlib
matplotlib.use("TkAgg")  # 强制交互后端
import matplotlib.pyplot as plt
from matplotlib import font_manager, rcParams

rcParams["font.sans-serif"] = ["SimHei"]      # 黑体
rcParams["axes.unicode_minus"] = False        # 解决负号显示问题

NULL_EFFECT_ID = 4294967295
FIXED_SHEET_NAME = "Relics"

# Column width tuning
CHAR_PIXELS = 9
PADDING_PIXELS = 24
MAX_COL_WIDTH = 520
EFFECT_NAME_FIXED_WIDTH = 320


# --------------------------
# helpers
# --------------------------
def normalize_int(x) -> Optional[int]:
    if x is None:
        return None
    if isinstance(x, int):
        return x
    if isinstance(x, float):
        return int(x) if x.is_integer() else int(x)
    s = str(x).strip()
    if s == "" or s.lower() == "nan":
        return None
    try:
        return int(float(s))
    except Exception:
        return None


def norm_colname(c: str) -> str:
    return re.sub(r"\s+", "", str(c)).lower()


def find_column(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    cols = list(df.columns)
    norm_cols = {c: norm_colname(c) for c in cols}
    for token in candidates:
        token_n = norm_colname(token)
        for c, cn in norm_cols.items():
            if token_n in cn:
                return c
    return None


def relic_type_from_item_id(item_id: Optional[int]) -> str:
    if item_id is None:
        return "未知"
    s = str(item_id)
    if 4 <= len(s) <= 5:
        return "Unique"
    if len(s) == 7 and s.startswith("2"):
        return "深夜"
    return "普通"


def score_effect_id(effect_id: Optional[int], score_map: Dict[str, int]) -> int:
    if effect_id is None or effect_id == NULL_EFFECT_ID:
        return 0
    return int(score_map.get(str(effect_id), 0))


def detect_effect_id_columns(df: pd.DataFrame) -> Tuple[str, str, str, str, str, str, Optional[str]]:
    item_col = find_column(df, ["itemid", "item_id", "item id", "物品id", "遗物id", "item"])

    e1 = find_column(df, ["effect1", "effect 1", "主词条1", "effect_1"])
    e2 = find_column(df, ["effect2", "effect 2", "主词条2", "effect_2"])
    e3 = find_column(df, ["effect3", "effect 3", "主词条3", "effect_3"])

    s1 = find_column(df, ["seceffect1", "sec effect 1", "secondaryeffect1", "副词条1", "sec_effect_1", "sec1"])
    s2 = find_column(df, ["seceffect2", "sec effect 2", "secondaryeffect2", "副词条2", "sec_effect_2", "sec2"])
    s3 = find_column(df, ["seceffect3", "sec effect 3", "secondaryeffect3", "副词条3", "sec_effect_3", "sec3"])

    missing = [name for name, col in [
        ("Effect1", e1), ("Effect2", e2), ("Effect3", e3),
        ("SecEffect1", s1), ("SecEffect2", s2), ("SecEffect3", s3),
    ] if col is None]

    if missing:
        raise ValueError(
            "没自动识别到这些列："
            + ", ".join(missing)
            + "\n\n我在 Excel 里看到的列名是：\n"
            + "\n".join([str(c) for c in df.columns])
        )
    return e1, e2, e3, s1, s2, s3, item_col


def detect_name_color_columns(df: pd.DataFrame) -> Tuple[Optional[str], Optional[str]]:
    name_col = find_column(df, ["relicname", "relic name", "name", "遗物名", "名称", "relic_name"])
    color_col = find_column(df, ["reliccolor", "relic color", "color", "颜色", "品质", "rarity", "quality", "relic_color"])
    return name_col, color_col


def guess_effect_name_columns(df: pd.DataFrame) -> List[str]:
    cols = []
    for c in df.columns:
        cn = norm_colname(c)
        if any(k in cn for k in ["effect", "seceffect", "主词条", "副词条", "sec"]):
            s = df[c].dropna().astype(str)
            if len(s) == 0:
                continue
            sample = s.head(50)
            numeric_like = sample.str.fullmatch(r"[-+]?\d+(\.0+)?").mean()
            if numeric_like < 0.6:
                cols.append(c)
    return cols


def compute_auto_width(series: pd.Series, header: str) -> int:
    s = series.fillna("").astype(str)
    max_len = max(len(str(header)), int(s.map(len).max()) if len(s) else len(str(header)))
    w = max_len * CHAR_PIXELS + PADDING_PIXELS
    return min(MAX_COL_WIDTH, max(70, w))


def load_score_json_folder(folder: str) -> Dict[str, Dict[str, int]]:
    """
    Return: {json_filename: {effect_id_str: score_int}}
    """
    out = {}
    if not folder or not os.path.isdir(folder):
        return out

    files = [f for f in os.listdir(folder) if f.lower().endswith(".json")]
    files.sort(key=lambda x: x.lower())

    for fn in files:
        path = os.path.join(folder, fn)
        try:
            with open(path, "r", encoding="utf-8") as f:
                d = json.load(f)
            # normalize
            d2 = {str(k): int(v) for k, v in d.items()}
            out[fn] = d2
        except Exception:
            # skip broken json
            continue
    return out


# --------------------------
# Sortable Treeview
# --------------------------
class SortableTreeview(ttk.Treeview):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self._sort_state = {}

    def _clear_arrows(self):
        for c in self["columns"]:
            text = self.heading(c, "text")
            text = re.sub(r"[ ▲▼]$", "", text)
            self.heading(c, text=text)

    def set_sort_indicator(self, col: str, asc: bool):
        self._clear_arrows()
        base = re.sub(r"[ ▲▼]$", "", self.heading(col, "text"))
        self.heading(col, text=f"{base} {'▲' if asc else '▼'}")

    def sort_by(self, col: str):
        asc = self._sort_state.get(col, True)

        data = []
        for iid in self.get_children(""):
            v = self.set(iid, col)
            data.append((v, iid))

        def to_sortable(x):
            s = "" if x is None else str(x).strip()
            if s == "":
                return (1, 0.0, "")
            try:
                return (0, float(s), "")
            except Exception:
                return (0, 0.0, s.lower())

        data.sort(key=lambda t: to_sortable(t[0]), reverse=not asc)
        for idx, (_, iid) in enumerate(data):
            self.move(iid, "", idx)

        self._sort_state[col] = not asc
        self.set_sort_indicator(col, asc)


# --------------------------
# UI App
# --------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("遗物自动评分")
        self.geometry("1480x860")

        self.excel_path = tk.StringVar(value="")
        self.json_folder = tk.StringVar(value="")

        # filters
        self.type_filter = tk.StringVar(value="全部")
        self.color_vars: Dict[str, tk.BooleanVar] = {}
        self.color_checkbuttons: List[ttk.Checkbutton] = []

        self.json_vars: Dict[str, tk.BooleanVar] = {}
        self.json_checkbuttons: List[ttk.Checkbutton] = []

        # data
        self.df_base: Optional[pd.DataFrame] = None   # after hiding effect IDs, before applying json set
        self.df_full: Optional[pd.DataFrame] = None   # after scoring (max)
        self.df_display: Optional[pd.DataFrame] = None

        self.color_col: Optional[str] = None
        self.effect_name_cols: List[str] = []
        self.score_maps: Dict[str, Dict[str, int]] = {}

        self._build_top_controls()
        self._build_filters()
        self._build_json_filters()
        self._build_table()
        
        

    # ---- UI building
    def _build_top_controls(self):
        frm = ttk.Frame(self)
        frm.pack(fill="x", padx=10, pady=10)
        frm.columnconfigure(1, weight=1)
        frm.columnconfigure(4, weight=1)

        ttk.Label(frm, text="Excel：").grid(row=0, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.excel_path, width=80).grid(row=0, column=1, sticky="we", padx=6)
        ttk.Button(frm, text="选择...", command=self.pick_excel).grid(row=0, column=2)

        ttk.Label(frm, text="分数JSON文件夹：").grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(frm, textvariable=self.json_folder, width=80).grid(row=1, column=1, sticky="we", padx=6, pady=(6, 0))
        ttk.Button(frm, text="选择...", command=self.pick_json_folder).grid(row=1, column=2, pady=(6, 0))

        ttk.Button(frm, text="加载并显示", command=self.load_and_show).grid(
            row=0, column=3, rowspan=1
        )

        ttk.Button(frm, text="Histogram", command=self.show_histogram).grid(
            row=1, column=3, rowspan=2, pady=(6, 0)
        )
        
        ttk.Button(frm, text="导出筛选后Excel", command=self.export_filtered_excel).grid(
            row=0, column=6, rowspan=2, padx=(10, 0), ipadx=10, ipady=10
        )
        
        self.status = tk.StringVar(value="")
        ttk.Label(frm, textvariable=self.status).grid(row=0, column=4, rowspan=2, sticky="nw", pady=(6, 0))
        
        self.count_text = tk.StringVar(value="")
        ttk.Label(frm, textvariable=self.count_text).grid(row=1, column=4, rowspan=2, sticky="nw", pady=(6, 0))


    def _build_filters(self):
        frm = ttk.LabelFrame(self, text="筛选：遗物类型 / 颜色")
        frm.pack(fill="x", padx=10, pady=(0, 10))

        ttk.Label(frm, text="遗物类型：").grid(row=0, column=0, sticky="w", padx=8, pady=8)
        type_combo = ttk.Combobox(frm, textvariable=self.type_filter, state="readonly",
                                  values=["全部", "深夜", "普通", "Unique"], width=10)
        type_combo.grid(row=0, column=1, sticky="w", pady=8)
        type_combo.bind("<<ComboboxSelected>>", lambda e: self.apply_filters())

        ttk.Label(frm, text="颜色：").grid(row=0, column=2, sticky="w", padx=(20, 8), pady=8)
        self.colors_container = ttk.Frame(frm)
        self.colors_container.grid(row=0, column=3, sticky="w", pady=8)

        ttk.Button(frm, text="全选颜色", command=lambda: self.set_all_colors(True)).grid(row=0, column=4, padx=(20, 6), pady=8)
        ttk.Button(frm, text="全不选", command=lambda: self.set_all_colors(False)).grid(row=0, column=5, padx=6, pady=8)

    def _build_json_filters(self):
        frm = ttk.LabelFrame(self, text="筛选：参与取最高分的JSON（勾选哪些就只用哪些计算）")
        frm.pack(fill="x", padx=10, pady=(0, 10))

        self.jsons_container = ttk.Frame(frm)
        self.jsons_container.pack(fill="x", padx=8, pady=8)

        btns = ttk.Frame(frm)
        btns.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Button(btns, text="全选JSON", command=lambda: self.set_all_jsons(True)).pack(side="left")
        ttk.Button(btns, text="全不选", command=lambda: self.set_all_jsons(False)).pack(side="left", padx=8)
        ttk.Button(btns, text="重新计算最高分", command=self.recompute_scores_only).pack(side="left", padx=8)
        
        self.del_threshold = tk.StringVar(value="60")  # 默认60
        ttk.Label(frm, text="删除阈值(<=)：").pack(side="left", padx=(20, 6))
        ttk.Entry(frm, textvariable=self.del_threshold, width=6).pack(side="left")
        ttk.Button(frm, text="删除这些遗物", command=self.delete_by_threshold).pack(side="left", padx=8)
        ttk.Label(frm, text="独特遗物不会被删除").pack(side="left", padx=4)
        
    def delete_by_threshold(self):
        """
        删除：总分 <= 阈值 的遗物（从当前数据源 df_full / df_display 对应的底层 df_with_effect_ids 同步删除）
        注意：这里只是删除 UI 内存中的数据，不会写回 Excel。
        """
        if not hasattr(self, "_df_with_effect_ids"):
            messagebox.showinfo("提示", "请先加载 Excel 和 JSON")
            return

        # 读取阈值
        try:
            thr = float(self.del_threshold.get().strip())
        except Exception:
            messagebox.showerror("错误", "阈值必须是数字，比如 60")
            return

        # 当前展示数据为空就不删
        if self.df_full is None or len(self.df_full) == 0:
            messagebox.showinfo("提示", "当前没有数据可删除")
            return

        if "总分" not in self.df_full.columns:
            messagebox.showerror("错误", "没找到“总分”列")
            return

        # 找到要删的 index（用 # 列做稳定定位）
        # 你的 # 是从1开始的序号，对应原 df 的行
        score_ok = pd.to_numeric(self.df_full["总分"], errors="coerce").fillna(-1) <= thr
        not_unique = self.df_full["遗物类型"] != "Unique"
        to_delete = self.df_full[score_ok & not_unique]["__orig_row"].tolist()
        unique_under_thr = self.df_full[score_ok & (self.df_full["遗物类型"] == "Unique")].shape[0]

        if not to_delete:
            messagebox.showinfo("提示", f"没有“总分 <= {thr}”的遗物")
            return

        # 确认
        if not messagebox.askyesno(
                "确认删除",
                f"将删除 {len(to_delete)} 条（总分 <= {thr}，不包含 Unique）。\n"
                f"另外有 {unique_under_thr} 条 Unique 低于阈值，但会被保留。\n继续？"
            ):
            return

        # 在底层 df 删除（_df_with_effect_ids 的 # 列就是 1..n）
        base = self._df_with_effect_ids
        base = base[~base["__orig_row"].isin(to_delete)].copy()

        # 重新生成 #（可选：你想保持原 # 不变就不要这段）
        base = base.reset_index(drop=True)
        base["#"] = base.index + 1

        self._df_with_effect_ids = base

        # 删除后需要重新计算分数 + 刷新 UI
        self.recompute_scores_only(first_time=False)

        messagebox.showinfo("完成", f"已删除 {len(to_delete)} 条（总分 <= {thr}）")
    
        
    def _build_table(self):
        container = ttk.Frame(self)
        container.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        container.rowconfigure(0, weight=1)
        container.columnconfigure(0, weight=1)

        self.tree = SortableTreeview(container, columns=[], show="headings")
        self.tree.grid(row=0, column=0, sticky="nsew")

        vsb = ttk.Scrollbar(container, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        
        
        
     
    def show_histogram(self):
        if self.df_display is None or len(self.df_display) == 0:
            messagebox.showinfo("提示", "当前没有可显示的数据（可能筛选后为空）")
            return
        if "总分" not in self.df_display.columns:
            messagebox.showerror("错误", "没找到“总分”列")
            return

        scores = pd.to_numeric(self.df_display["总分"], errors="coerce").dropna()
        if len(scores) == 0:
            messagebox.showinfo("提示", "总分列没有有效数值")
            return

        import matplotlib.pyplot as plt
        import numpy as np

        bins = np.arange(0, 110, 10)  # 0-100, step 10
        counts, edges = np.histogram(scores, bins=bins)

        total = counts.sum() if counts.sum() > 0 else 1
        perc = counts / total * 100

        plt.figure()
        n, b, patches = plt.hist(
            scores,
            bins=bins,
            edgecolor="black",   # 每个柱子的黑色边框
            linewidth=1.0
        )

        plt.title("总分 Histogram（百分制）")
        plt.xlabel("总分区间")
        plt.ylabel("数量")
        
        plt.grid(True, linestyle="--", alpha=0.5)

        # 给每个柱子标注：count + percentage
        for i, p in enumerate(patches):
            h = p.get_height()
            if h <= 0:
                continue
            # 柱子中心
            x = p.get_x() + p.get_width() / 2
            label = f"{int(counts[i])}\n({perc[i]:.1f}%)"
            plt.text(x, h, label, ha="center", va="bottom", fontsize=9)

        plt.show()



    # ---- pickers
    def pick_excel(self):
        p = filedialog.askopenfilename(title="选择 nightreign.xlsx", filetypes=[("Excel Files", "*.xlsx *.xls")])
        if p:
            self.excel_path.set(p)

    def pick_json_folder(self):
        p = filedialog.askdirectory(title="选择包含多个score json的文件夹")
        if p:
            self.json_folder.set(p)

    # ---- filter helpers
    def set_all_colors(self, value: bool):
        for v in self.color_vars.values():
            v.set(value)
        self.apply_filters()

    def set_all_jsons(self, value: bool):
        for v in self.json_vars.values():
            v.set(value)
        self.recompute_scores_only()

    def rebuild_color_filters(self, colors: List[str]):
        for cb in self.color_checkbuttons:
            cb.destroy()
        self.color_checkbuttons.clear()
        self.color_vars.clear()

        colors = [c for c in colors if str(c).strip() != ""]
        colors = list(dict.fromkeys(colors))[:4]  # 只保留前4种
        for i, c in enumerate(colors):
            var = tk.BooleanVar(value=True)
            self.color_vars[c] = var
            cb = ttk.Checkbutton(self.colors_container, text=str(c), variable=var, command=self.apply_filters)
            cb.grid(row=0, column=i, sticky="w", padx=6)
            self.color_checkbuttons.append(cb)

    def rebuild_json_filters(self, json_names: List[str]):
        for cb in self.json_checkbuttons:
            cb.destroy()
        self.json_checkbuttons.clear()
        self.json_vars.clear()

        # grid display, wrap
        cols = 4
        for idx, name in enumerate(json_names):
            var = tk.BooleanVar(value=True)
            self.json_vars[name] = var
            cb = ttk.Checkbutton(self.jsons_container, text=name, variable=var, command=self.recompute_scores_only)
            r = idx // cols
            c = idx % cols
            cb.grid(row=r, column=c, sticky="w", padx=8, pady=2)
            self.json_checkbuttons.append(cb)

    # --------------------------
    # Core logic
    # --------------------------
    def load_and_show(self):
        xlsx = self.excel_path.get().strip()
        folder = self.json_folder.get().strip()

        if not xlsx or not folder:
            messagebox.showwarning("提示", "请先选择 Excel 和 JSON文件夹")
            return

        # load json folder
        self.score_maps = load_score_json_folder(folder)
        if not self.score_maps:
            messagebox.showerror("错误", "该文件夹没有可用的 .json（或json解析失败）")
            return

        self.rebuild_json_filters(list(self.score_maps.keys()))

        try:
            df0 = pd.read_excel(xlsx, sheet_name=FIXED_SHEET_NAME, engine="openpyxl")
            self._df_original = df0.copy()
            e1, e2, e3, s1, s2, s3, item_col = detect_effect_id_columns(df0)
            name_col, color_col = detect_name_color_columns(df0)
            self.color_col = color_col

            df = df0.copy()
            df.insert(0, "#", df.index + 1)
            df["__orig_row"] = df["#"] 

            if item_col is not None:
                item_vals = df[item_col].apply(normalize_int)
                df.insert(1, "遗物类型", item_vals.apply(relic_type_from_item_id))
            else:
                df.insert(1, "遗物类型", "未知")

            # normalize effect ids
            for col in [e1, e2, e3, s1, s2, s3]:
                df[col] = df[col].apply(normalize_int)

            # base columns: we keep effect IDs internally for recompute, but we'll remove from display after scoring
            # store effect id col names for later
            self._effect_id_cols = (e1, e2, e3, s1, s2, s3)

            # move name/color forward later
            self._name_col = name_col
            self._color_col = color_col
            self._item_col = item_col

            # save base df with effect IDs so we can recompute without re-reading Excel
            self._df_with_effect_ids = df

            # compute initial max scoring and display
            self.recompute_scores_only(first_time=True)

        except Exception as e:
            messagebox.showerror("错误", f"加载Excel失败：\n{e}")

    def recompute_scores_only(self, first_time: bool = False):
        """
        Recompute max score based on checked jsons, then apply filters and refresh table.
        """
        if not hasattr(self, "_df_with_effect_ids"):
            return

        df = self._df_with_effect_ids.copy()
        e1, e2, e3, s1, s2, s3 = self._effect_id_cols

        chosen = [name for name, var in self.json_vars.items() if var.get()]
        if not chosen:
            # no json selected -> scores all 0 and source empty
            df["词条1"] = 0
            df["词条2"] = 0
            df["词条3"] = 0
            df["负面词条1"] = 0
            df["负面词条2"] = 0
            df["负面词条3"] = 0
            df["总分"] = 0
            df["分数来源"] = ""
            score_used = 0
        else:
            # compute best per row by looping jsons
            best_total = []
            best_src = []

            # also keep best breakdown for display
            b1, b2, b3, bs1, bs2, bs3 = [], [], [], [], [], []
            n = len(df)
            
            # Pre-get effect id arrays (faster)
            ids1 = df[e1].tolist()
            ids2 = df[e2].tolist()
            ids3 = df[e3].tolist()
            ids4 = df[s1].tolist()
            ids5 = df[s2].tolist()
            ids6 = df[s3].tolist()

            # Initialize with -inf
            
            cur_best = [-10**18] * n
            cur_best_raw = [0] * n
            cur_src = [""] * n
            cur_b = [[0]*n for _ in range(6)]

            for jn in chosen:
                smap = self.score_maps.get(jn, {})
                # compute totals for this json
                totals = []
                t1, t2, t3, t4, t5, t6 = [], [], [], [], [], []
                for i in range(n):
                    x1 = score_effect_id(ids1[i], smap)
                    x2 = score_effect_id(ids2[i], smap)
                    x3 = score_effect_id(ids3[i], smap)
                    x4 = score_effect_id(ids4[i], smap)
                    x5 = score_effect_id(ids5[i], smap)
                    x6 = score_effect_id(ids6[i], smap)
                    tot = x1 + x2 + x3 + x4 + x5 + x6
                    totals.append(tot)
                    t1.append(x1); t2.append(x2); t3.append(x3)
                    t4.append(x4); t5.append(x5); t6.append(x6)
                    
                # 2) minmax scale within this json
                mn = min(totals) if totals else 0
                mx = max(totals) if totals else 0
                if mx == mn:
                    scaled_totals = [0.0] * n
                else:
                    denom = (mx - mn)
                    scaled_totals = [(rt - mn) / denom for rt in totals]
                    
                # update best
                for i in range(n):
                    if scaled_totals[i] > cur_best[i]:
                        cur_best[i] = scaled_totals[i]
                        cur_best_raw[i] = totals[i]
                        cur_src[i] = jn
                        cur_b[0][i] = t1[i]
                        cur_b[1][i] = t2[i]
                        cur_b[2][i] = t3[i]
                        cur_b[3][i] = t4[i]
                        cur_b[4][i] = t5[i]
                        cur_b[5][i] = t6[i]

            df["词条1"] = cur_b[0]
            df["词条2"] = cur_b[1]
            df["词条3"] = cur_b[2]
            df["负面词条1"] = cur_b[3]
            df["负面词条2"] = cur_b[4]
            df["负面词条3"] = cur_b[5]
            df["Raw总分"] = cur_best_raw
            df["总分"] = [round(x*100, 2) for x in cur_best]
            df["分数来源"] = [os.path.splitext(x)[0] for x in cur_src]
            score_used = len(chosen)
            
        # Hide 6 effect ID columns from display (but keep other columns)
        hidden_id_cols = set(self._effect_id_cols)
        df_display = df[[c for c in df.columns if c not in hidden_id_cols]].copy()

        # effect name columns for width rules
        self.effect_name_cols = guess_effect_name_columns(df_display)

        # reorder columns
        front = ["#", "遗物类型"]
        name_col = self._name_col
        color_col = self._color_col
        if name_col is not None and name_col in df_display.columns:
            front.append(name_col)
        if color_col is not None and color_col in df_display.columns:
            front.append(color_col)

        front += ["词条1", "词条2", "词条3", "负面词条1", "负面词条2", "负面词条3","Raw总分", "总分", "分数来源"]
        front = [c for c in front if c in df_display.columns]
        rest = [c for c in df_display.columns if c not in front]
        df_display = df_display[front + rest]

        # Save full for filtering
        self.df_full = df_display

        # Build color filters once
        if first_time and self._color_col and self._color_col in self.df_full.columns:
            colors = self.df_full[self._color_col].dropna().astype(str).tolist()
            seen = set()
            uniq = []
            for c in colors:
                c = c.strip()
                if c and c not in seen:
                    uniq.append(c)
                    seen.add(c)
            self.rebuild_color_filters(uniq)

        self.status.set(f"已加载JSON: {len(self.score_maps)} 个 | 参与计算: {score_used} 个")
        self.apply_filters()

    def apply_filters(self):
        if self.df_full is None:
            self.count_text.set("显示：0 / 0")
            return

        df = self.df_full.copy()

        # type filter
        t = self.type_filter.get()
        if t != "全部" and "遗物类型" in df.columns:
            df = df[df["遗物类型"] == t]

        # color filter
        if self._color_col and self._color_col in df.columns and self.color_vars:
            allowed = {c for c, v in self.color_vars.items() if v.get()}
            if allowed:
                df = df[df[self._color_col].astype(str).isin(allowed)]
            else:
                df = df.iloc[0:0]

        self.df_display = df
        self.show_dataframe(df)
        
        total_n = len(self.df_full) if self.df_full is not None else 0
        shown_n = len(df)
        self.count_text.set(f"显示：{shown_n} / {total_n}")


    def show_dataframe(self, df: pd.DataFrame):
        self.tree.delete(*self.tree.get_children(""))
        show_cols = [c for c in df.columns if c != "__orig_row"]  # 不显示原始行号
        self.tree["columns"] = show_cols


        # compute widths
        col_widths = {}
        for c in show_cols:
            if c in self.effect_name_cols:
                col_widths[c] = EFFECT_NAME_FIXED_WIDTH
            else:
                col_widths[c] = compute_auto_width(df[c], header=str(c))

        for c in show_cols:
            self.tree.heading(c, text=str(c), command=lambda cc=c: self.tree.sort_by(cc))
            self.tree.column(c, width=col_widths[c], anchor="center", stretch=False)

        for _, row in df.iterrows():
            vals = [("" if pd.isna(row[c]) else row[c]) for c in df.columns]
            self.tree.insert("", "end", values=vals)
    
    
    def export_filtered_excel(self):
        if self.df_full is None or len(self.df_full) == 0:
            messagebox.showinfo("提示", "当前没有可导出的数据（请先加载并显示）")
            return
        if not hasattr(self, "_df_original"):
            messagebox.showinfo("提示", "没有原始Excel数据（请重新加载Excel）")
            return

        out_path = filedialog.asksaveasfilename(
            title="保存筛选后的Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if not out_path:
            return

        try:
            # 当前“保留”的行号集合（# 从1开始，对应原始 df 的行）
            kept_indices = set(pd.to_numeric(self.df_full["__orig_row"], errors="coerce").dropna().astype(int).tolist())

            df_out = self._df_original.copy()
            mask = [(i + 1) in kept_indices for i in range(len(df_out))]
            df_out = df_out.loc[mask].copy()

            with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
                df_out.to_excel(writer, sheet_name=FIXED_SHEET_NAME, index=False)

            messagebox.showinfo("完成", f"已导出：\n{out_path}")
        except Exception as e:
            messagebox.showerror("导出失败", str(e))



def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
