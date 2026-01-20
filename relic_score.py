# relic_score_ui.py
# -*- coding: utf-8 -*-

import json
import re
from typing import Dict, Optional, List, Tuple

import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


NULL_EFFECT_ID = 4294967295
FIXED_SHEET_NAME = "Relics"

# 列宽参数（你想更宽/更窄就改这里）
CHAR_PIXELS = 9          # 每个字符大约像素
PADDING_PIXELS = 24      # 额外 padding
MAX_COL_WIDTH = 520      # 单列最大宽度（防止超级长文本撑爆）
EFFECT_NAME_FIXED_WIDTH = 320  # 6个 effect name 列固定宽度


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
    """
    你说的“6个effect name”如果在 Excel 里是文本列（不是数字ID），这里会尽量自动抓出来：
    - 列名包含 effect/sec/主词条/副词条 等关键词
    - 且列内容更像文本（非纯数字）
    """
    cols = []
    for c in df.columns:
        cn = norm_colname(c)
        if any(k in cn for k in ["effect", "seceffect", "主词条", "副词条", "sec"]):
            # 判断这列是否“像文本列”
            s = df[c].dropna().astype(str)
            if len(s) == 0:
                continue
            # 取一小部分样本，数字占比低 -> 当作 name 列
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
        self.title("遗物自动评分 UI")
        self.geometry("1400x820")

        self.excel_path = tk.StringVar(value="")
        self.json_path = tk.StringVar(value="")

        # filters
        self.type_filter = tk.StringVar(value="全部")  # 全部/深夜/普通/Unique
        self.color_vars: Dict[str, tk.BooleanVar] = {}
        self.color_checkbuttons: List[ttk.Checkbutton] = []

        self.df_full: Optional[pd.DataFrame] = None
        self.df_display: Optional[pd.DataFrame] = None
        self.color_col: Optional[str] = None
        self.effect_name_cols: List[str] = []

        self._build_top_controls()
        self._build_filters()
        self._build_table()

    def _build_top_controls(self):
        frm = ttk.Frame(self)
        frm.pack(fill="x", padx=10, pady=10)
        frm.columnconfigure(1, weight=1)

        ttk.Label(frm, text="Excel：").grid(row=0, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.excel_path, width=80).grid(row=0, column=1, sticky="we", padx=6)
        ttk.Button(frm, text="选择...", command=self.pick_excel).grid(row=0, column=2)

        ttk.Label(frm, text="分数JSON：").grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(frm, textvariable=self.json_path, width=80).grid(row=1, column=1, sticky="we", padx=6, pady=(6, 0))
        ttk.Button(frm, text="选择...", command=self.pick_json).grid(row=1, column=2, pady=(6, 0))

        ttk.Button(frm, text="加载并显示", command=self.load_and_show).grid(
            row=0, column=3, rowspan=2, padx=(14, 0), ipadx=12, ipady=10
        )

    def _build_filters(self):
        frm = ttk.LabelFrame(self, text="筛选")
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

    def pick_excel(self):
        p = filedialog.askopenfilename(title="选择 nightreign.xlsx", filetypes=[("Excel Files", "*.xlsx *.xls")])
        if p:
            self.excel_path.set(p)

    def pick_json(self):
        p = filedialog.askopenfilename(title="选择 分数JSON", filetypes=[("JSON Files", "*.json")])
        if p:
            self.json_path.set(p)

    def set_all_colors(self, value: bool):
        for v in self.color_vars.values():
            v.set(value)
        self.apply_filters()

    def rebuild_color_filters(self, colors: List[str]):
        # 清空旧的
        for cb in self.color_checkbuttons:
            cb.destroy()
        self.color_checkbuttons.clear()
        self.color_vars.clear()

        # 只取前4种（你说固定四种）
        colors = [c for c in colors if str(c).strip() != ""]
        colors = list(dict.fromkeys(colors))  # 保序去重
        colors = colors[:4]

        for i, c in enumerate(colors):
            var = tk.BooleanVar(value=True)
            self.color_vars[c] = var
            cb = ttk.Checkbutton(self.colors_container, text=str(c), variable=var, command=self.apply_filters)
            cb.grid(row=0, column=i, sticky="w", padx=6)
            self.color_checkbuttons.append(cb)

    def load_and_show(self):
        xlsx = self.excel_path.get().strip()
        jpath = self.json_path.get().strip()

        if not xlsx or not jpath:
            messagebox.showwarning("提示", "请先选择 Excel 和 分数JSON")
            return

        # load score map
        try:
            with open(jpath, "r", encoding="utf-8") as f:
                score_map = json.load(f)
            score_map = {str(k): int(v) for k, v in score_map.items()}
        except Exception as e:
            messagebox.showerror("错误", f"读取 JSON 失败：\n{e}")
            return

        try:
            df0 = pd.read_excel(xlsx, sheet_name=FIXED_SHEET_NAME, engine="openpyxl")
            e1, e2, e3, s1, s2, s3, item_col = detect_effect_id_columns(df0)
            name_col, color_col = detect_name_color_columns(df0)
            self.color_col = color_col

            df = df0.copy()
            df.insert(0, "#", df.index + 1)

            # type
            if item_col is not None:
                item_vals = df[item_col].apply(normalize_int)
                df.insert(1, "遗物类型", item_vals.apply(relic_type_from_item_id))
            else:
                df.insert(1, "遗物类型", "未知")

            # normalize effect ids
            for col in [e1, e2, e3, s1, s2, s3]:
                df[col] = df[col].apply(normalize_int)

            # scores
            df["Effect1Score"] = df[e1].apply(lambda x: score_effect_id(x, score_map))
            df["Effect2Score"] = df[e2].apply(lambda x: score_effect_id(x, score_map))
            df["Effect3Score"] = df[e3].apply(lambda x: score_effect_id(x, score_map))
            df["SecEffect1Score"] = df[s1].apply(lambda x: score_effect_id(x, score_map))
            df["SecEffect2Score"] = df[s2].apply(lambda x: score_effect_id(x, score_map))
            df["SecEffect3Score"] = df[s3].apply(lambda x: score_effect_id(x, score_map))
            
            df["TotalScore"] = (
                df["Effect1Score"] + df["Effect2Score"] + df["Effect3Score"]
                + df["SecEffect1Score"] + df["SecEffect2Score"] + df["SecEffect3Score"]
            )
            
            df.rename(columns={
                "Effect1Score": "词条1",
                "Effect2Score": "词条2",
                "Effect3Score": "词条3",
                "SecEffect1Score": "负面词条1",
                "SecEffect2Score": "负面词条2",
                "SecEffect3Score": "负面词条3",
                "TotalScore": "总分",
            }, inplace=True)


            # hide 6 effect ID cols
            hidden_id_cols = {e1, e2, e3, s1, s2, s3}
            df_display = df[[c for c in df.columns if c not in hidden_id_cols]].copy()

            # effect name cols (用于列宽规则)
            self.effect_name_cols = guess_effect_name_columns(df_display)

            # move name/color forward
            front = ["#", "遗物类型"]
            if name_col is not None and name_col in df_display.columns:
                front.append(name_col)
            if color_col is not None and color_col in df_display.columns:
                front.append(color_col)

            front += [
                "词条1", "词条2", "词条3",
                "负面词条1", "负面词条2", "负面词条3",
                "总分",
            ]
            front = [c for c in front if c in df_display.columns]
            rest = [c for c in df_display.columns if c not in front]
            df_display = df_display[front + rest]
           
            # 记录全量，用于筛选
            self.df_full = df_display.copy()

            # build color filters
            if self.color_col and self.color_col in self.df_full.columns:
                colors = self.df_full[self.color_col].dropna().astype(str).tolist()
                # 保序去重
                seen = set()
                uniq = []
                for c in colors:
                    c = c.strip()
                    if c and c not in seen:
                        uniq.append(c)
                        seen.add(c)
                self.rebuild_color_filters(uniq)

            # 初次显示
            self.apply_filters()

        except Exception as e:
            messagebox.showerror("错误", f"加载/打分失败：\n{e}")

    def apply_filters(self):
        if self.df_full is None:
            return

        df = self.df_full.copy()

        # type filter
        t = self.type_filter.get()
        if t != "全部" and "遗物类型" in df.columns:
            df = df[df["遗物类型"] == t]

        # color filter
        if self.color_col and self.color_col in df.columns and self.color_vars:
            allowed = {c for c, v in self.color_vars.items() if v.get()}
            if allowed:
                df = df[df[self.color_col].astype(str).isin(allowed)]
            else:
                df = df.iloc[0:0]  # 全不选 -> 空表

        self.df_display = df
        self.show_dataframe(df)

    def show_dataframe(self, df: pd.DataFrame):
        self.tree.delete(*self.tree.get_children(""))
        self.tree["columns"] = list(df.columns)

        # 先算列宽（自适应：除了 effect name 列）
        col_widths = {}
        for c in df.columns:
            if c in self.effect_name_cols:
                col_widths[c] = EFFECT_NAME_FIXED_WIDTH
            else:
                col_widths[c] = compute_auto_width(df[c], header=str(c))

        # headings + columns
        for c in df.columns:
            self.tree.heading(c, text=str(c), command=lambda cc=c: self.tree.sort_by(cc))
            self.tree.column(c, width=col_widths[c], anchor="center", stretch=False)

        # insert rows
        for _, row in df.iterrows():
            vals = [("" if pd.isna(row[c]) else row[c]) for c in df.columns]
            self.tree.insert("", "end", values=vals)


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
