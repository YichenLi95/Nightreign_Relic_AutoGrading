# export_relic_scores.py
# -*- coding: utf-8 -*-
"""
从 Excel 指定 sheet 导出 {ID: Score} 的 JSON。
兼容一些 Excel 样式损坏导致 openpyxl/pandas 读取报错的情况：使用“解压 xlsx + 解析 xml”方式读取单元格值。
用法示例：
  python export_relic_scores.py "遗物筛选.xlsx" --sheet "小蜗（兽爪）"
  python export_relic_scores.py "遗物筛选.xlsx" --list-sheets
  python export_relic_scores.py "遗物筛选.xlsx" --auto --out relic_scores.json
  python export_relic_scores.py "遗物筛选.xlsx" --all --out relic_scores_by_sheet.json
"""
import os
import argparse
import json
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from typing import Dict, List, Tuple, Optional


NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"


def col_to_idx(col_letters: str) -> int:
    idx = 0
    for ch in col_letters:
        idx = idx * 26 + (ord(ch.upper()) - ord("A") + 1)
    return idx


def load_shared_strings(z: zipfile.ZipFile) -> List[str]:
    try:
        ss_xml = z.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    root = ET.fromstring(ss_xml)
    ns = {"ns": NS_MAIN}
    strings = []
    for si in root.findall("ns:si", ns):
        texts = [t.text or "" for t in si.findall(".//ns:t", ns)]
        strings.append("".join(texts))
    return strings


def read_workbook_sheets(z: zipfile.ZipFile) -> List[Dict[str, str]]:
    wb_xml = ET.fromstring(z.read("xl/workbook.xml"))
    ns = {"ns": NS_MAIN, "r": NS_REL}
    sheets = []
    for sh in wb_xml.findall("ns:sheets/ns:sheet", ns):
        name = sh.attrib.get("name")
        rid = sh.attrib.get(f"{{{NS_REL}}}id")
        sheet_id = sh.attrib.get("sheetId")
        sheets.append({"name": name, "rid": rid, "sheetId": sheet_id})

    rels_xml = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rel_ns = {"rel": NS_PKG_REL}
    rid_to_target = {}
    for rel in rels_xml.findall("rel:Relationship", rel_ns):
        rid_to_target[rel.attrib["Id"]] = rel.attrib["Target"]

    for sh in sheets:
        target = rid_to_target.get(sh["rid"])
        if target and not target.startswith("xl/"):
            target = "xl/" + target
        sh["path"] = target
    return sheets


def parse_sheet_cells(
    z: zipfile.ZipFile,
    sheet_path: str,
    shared_strings: List[str],
    max_rows: Optional[int] = None
) -> Tuple[Dict[Tuple[int, int], Optional[str]], int, int]:
    """
    返回：
      cells[(row, col)] = value(str or None)
      max_row, max_col
    """
    xml = ET.fromstring(z.read(sheet_path))
    ns = {"ns": NS_MAIN}

    cells: Dict[Tuple[int, int], Optional[str]] = {}
    max_r = 0
    max_c = 0

    for c in xml.findall(".//ns:sheetData/ns:row/ns:c", ns):
        ref = c.attrib.get("r")  # e.g. A1
        if not ref:
            continue
        m = re.match(r"([A-Z]+)(\d+)", ref)
        if not m:
            continue
        col_letters, row_num = m.group(1), int(m.group(2))
        if max_rows is not None and row_num > max_rows:
            continue

        col_num = col_to_idx(col_letters)
        cell_type = c.attrib.get("t")  # s / inlineStr / None

        value: Optional[str] = None
        if cell_type == "s":
            v_node = c.find("ns:v", ns)
            if v_node is not None and v_node.text is not None:
                s_idx = int(v_node.text)
                value = shared_strings[s_idx] if 0 <= s_idx < len(shared_strings) else None
        elif cell_type == "inlineStr":
            t_node = c.find(".//ns:t", ns)
            value = t_node.text if t_node is not None else None
        else:
            v_node = c.find("ns:v", ns)
            value = v_node.text if v_node is not None else None

        cells[(row_num, col_num)] = value
        max_r = max(max_r, row_num)
        max_c = max(max_c, col_num)

    return cells, max_r, max_c


def read_rows(
    z: zipfile.ZipFile,
    sheet_path: str,
    shared_strings: List[str],
    max_rows: Optional[int] = None
) -> List[List[Optional[str]]]:
    cells, max_r, max_c = parse_sheet_cells(z, sheet_path, shared_strings, max_rows=max_rows)
    if max_rows is not None:
        max_r = min(max_r, max_rows)

    rows: List[List[Optional[str]]] = []
    for r in range(1, max_r + 1):
        row = [cells.get((r, c)) for c in range(1, max_c + 1)]
        rows.append(row)
    return rows


def find_id_score_columns(header_row: List[Optional[str]]) -> Optional[Tuple[int, int]]:
    headers = [(h or "").strip() for h in header_row]
    id_idx = None
    score_idx = None
    for i, h in enumerate(headers):
        if h.upper() == "ID":
            id_idx = i
        if h.lower() == "score" or h == "Score" or "评分" in h:
            score_idx = i
    if id_idx is None or score_idx is None:
        return None
    return id_idx, score_idx


def to_int_like(x: Optional[str]) -> int:
    if x is None:
        return 0
    s = str(x).strip()
    if s == "":
        return 0
    try:
        return int(float(s))
    except Exception:
        return 0


def export_one_sheet(xlsx_path: str, sheet_name: str) -> Dict[str, Optional[int]]:
    with zipfile.ZipFile(xlsx_path) as z:
        shared_strings = load_shared_strings(z)
        sheets = read_workbook_sheets(z)

        target = None
        for sh in sheets:
            if sh["name"] == sheet_name:
                target = sh
                break
        if target is None:
            raise ValueError(f"找不到 sheet: {sheet_name}. 可用: {[s['name'] for s in sheets]}")

        rows_preview = read_rows(z, target["path"], shared_strings, max_rows=5)
        if not rows_preview:
            return {}

        col_pair = find_id_score_columns(rows_preview[0])
        if col_pair is None:
            raise ValueError(f"sheet '{sheet_name}' 第一行没找到 ID/Score 列（header）")

        id_idx, score_idx = col_pair

        rows_all = read_rows(z, target["path"], shared_strings, max_rows=None)
        out: Dict[str, Optional[int]] = {}
        for row in rows_all[1:]:
            if id_idx >= len(row):
                continue
            rid = row[id_idx]
            if rid is None or str(rid).strip() == "":
                continue
            # ID 强制转成 int-like 再转字符串，避免 "7000000.0"
            rid_int = to_int_like(str(rid))  # type: ignore
            rid_key = str(rid_int) if rid_int is not None else str(rid).strip()

            score_val = None
            if score_idx < len(row):
                score_val = to_int_like(row[score_idx])
            out[rid_key] = score_val
        return out


def export_all_sheets(xlsx_path: str) -> Dict[str, Dict[str, Optional[int]]]:
    with zipfile.ZipFile(xlsx_path) as z:
        shared_strings = load_shared_strings(z)
        sheets = read_workbook_sheets(z)

    result: Dict[str, Dict[str, Optional[int]]] = {}
    for sh in sheets:
        try:
            d = export_one_sheet(xlsx_path, sh["name"])
            # 只有当确实包含数据时再收录（避免空表）
            if len(d) > 0:
                result[sh["name"]] = d
        except Exception:
            # 不是每个 sheet 都有 ID/Score header，忽略
            continue
    return result


def detect_sheets_with_id_score(xlsx_path: str) -> List[str]:
    with zipfile.ZipFile(xlsx_path) as z:
        shared_strings = load_shared_strings(z)
        sheets = read_workbook_sheets(z)

        ok_names = []
        for sh in sheets:
            rows_preview = read_rows(z, sh["path"], shared_strings, max_rows=3)
            if rows_preview and find_id_score_columns(rows_preview[0]) is not None:
                ok_names.append(sh["name"])
        return ok_names


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx", help="Excel 路径，如 遗物筛选.xlsx")
    ap.add_argument("--sheet", help="指定导出某个 sheet")
    ap.add_argument("--out", default=None, help="输出 json 路径（默认自动命名）")
    ap.add_argument("--list-sheets", action="store_true", help="列出所有 sheet 名字")
    ap.add_argument("--auto", action="store_true", help="自动选择第一个含 ID/Score 的 sheet 导出")
    ap.add_argument("--all", action="store_true", help="导出所有含 ID/Score 的 sheet，按 sheet 分组")
    args = ap.parse_args()

    if args.list_sheets:
        with zipfile.ZipFile(args.xlsx) as z:
            sheets = read_workbook_sheets(z)
        print("Sheets:")
        for s in sheets:
            print("-", s["name"])
        return

    if args.all:
        # 输出目录：默认当前目录；你也可以用 --out 指定目录
        out_dir = args.out or "relic_scores_out"
        os.makedirs(out_dir, exist_ok=True)

        ok_sheets = detect_sheets_with_id_score(args.xlsx)
        if not ok_sheets:
            raise SystemExit("没有找到包含 ID/Score 的 sheet")

        # 逐个 sheet 导出
        for sh in ok_sheets[1:]:
            data = export_one_sheet(args.xlsx, sh)
            # 清理文件名（避免 Windows 不允许的字符）
            safe_name = re.sub(r'[\\/:*?"<>|]+', "_", sh).strip()
            out_path = os.path.join(out_dir, f"{safe_name}.json")
            with open(out_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            print(f"已导出 sheet '{sh}' -> {out_path} ({len(data)} 条)")

        return

    if args.auto and not args.sheet:
        ok = detect_sheets_with_id_score(args.xlsx)
        if not ok:
            raise SystemExit("没有找到包含 ID/Score 的 sheet")
        args.sheet = ok[0]
        print(f"[auto] 选择 sheet: {args.sheet}")

    if not args.sheet:
        raise SystemExit("请用 --sheet 指定 sheet，或用 --auto 自动选择，或用 --all 导出全部。")

    data = export_one_sheet(args.xlsx, args.sheet)
    out_path = args.out or f"relic_scores_{args.sheet}.json"
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"已导出 {len(data)} 条 ID->Score -> {out_path}")


if __name__ == "__main__":
    main()
