import pandas as pd
from pathlib import Path
from datetime import datetime
import re
import sys


# ==================================================
# 自动识别关键词
# ==================================================

BESTELL_KEYWORDS = [
    "Daimler_Bestellunterlagen"
]

FG_KEYWORDS = [
    "Daimler_FGPruefung",
    "Daimler_FGPrüfung",
    "Daimler_FGPrufung"
]


# ==================================================
# 工具函数
# ==================================================

def clean_value(value):
    """清理单元格内容"""
    if pd.isna(value):
        return ""

    value = str(value).replace("\xa0", " ")
    value = re.sub(r"\s+", " ", value)

    return value.strip()


def find_latest_xlsx_by_keywords(keywords):
    """
    自动查找当前文件夹中名称包含关键词的 xlsx 文件。
    如果有多个，选择最后修改时间最新的一个。
    """
    candidates = []

    for p in Path(".").glob("*.xlsx"):

        if p.name.startswith("~$"):
            continue

        name_lower = p.name.lower()

        for kw in keywords:
            if kw.lower() in name_lower:
                candidates.append(p)
                break

    if not candidates:
        raise FileNotFoundError(
            f"未找到名称包含以下关键词的xlsx文件: {keywords}"
        )

    # 选择最新修改的文件
    candidates = sorted(
        candidates,
        key=lambda x: x.stat().st_mtime,
        reverse=True
    )

    print("找到候选文件:")
    for p in candidates:
        print(" -", p.name)

    print("使用文件:", candidates[0].name)
    print()

    return candidates[0]


def find_column(df, possible_names):
    """
    自动识别列名
    """
    for name in possible_names:
        if name in df.columns:
            return name

    raise ValueError(
        f"未找到列名。候选列名: {possible_names}，实际列名: {list(df.columns)}"
    )


def build_code_dict(df, code_col, desc_col):
    """
    建立 Code -> Description 字典
    保留第一次出现的 Code 顺序
    """
    result = {}

    for _, row in df.iterrows():

        code = clean_value(row.get(code_col, ""))

        if not code:
            continue

        desc = clean_value(row.get(desc_col, ""))

        if code not in result:
            result[code] = desc

    return result


# ==================================================
# 主程序
# ==================================================

def main():

    # ==================================================
    # 自动识别两个 Excel 文件
    # ==================================================

    bestell_file = find_latest_xlsx_by_keywords(
        BESTELL_KEYWORDS
    )

    fg_file = find_latest_xlsx_by_keywords(
        FG_KEYWORDS
    )

    print("Bestellunterlagen Excel:", bestell_file.name)
    print("FG Prüfung Excel:", fg_file.name)
    print()

    # ==================================================
    # 读取 Excel
    # ==================================================

    df_bestell = pd.read_excel(
        bestell_file,
        dtype=str,
        engine="openpyxl"
    )

    df_fg = pd.read_excel(
        fg_file,
        dtype=str,
        engine="openpyxl"
    )

    # ==================================================
    # 自动识别列名
    # ==================================================

    bestell_code_col = find_column(
        df_bestell,
        ["Code", "VCode"]
    )

    bestell_desc_col = find_column(
        df_bestell,
        ["Description", "Beschreibung", "Name"]
    )

    fg_code_col = find_column(
        df_fg,
        ["Code", "VCode"]
    )

    fg_desc_col = find_column(
        df_fg,
        ["Description", "Beschreibung", "Name"]
    )

    print("Bestell Code列:", bestell_code_col)
    print("Bestell Description列:", bestell_desc_col)
    print("FG Code列:", fg_code_col)
    print("FG Description列:", fg_desc_col)
    print()

    # ==================================================
    # 建立字典
    # ==================================================

    bestell_dict = build_code_dict(
        df_bestell,
        bestell_code_col,
        bestell_desc_col
    )

    fg_dict = build_code_dict(
        df_fg,
        fg_code_col,
        fg_desc_col
    )

    bestell_codes = set(bestell_dict.keys())
    fg_codes = set(fg_dict.keys())

    # ==================================================
    # 比较差异
    # ==================================================

    only_in_bestell = bestell_codes - fg_codes
    only_in_fg = fg_codes - bestell_codes
    matched = bestell_codes & fg_codes

    # ==================================================
    # 输出表：Only in Bestellunterlagen
    # ==================================================

    only_bestell_rows = []

    for code in bestell_dict.keys():

        if code in only_in_bestell:

            only_bestell_rows.append({
                "Code": code,
                "Description": bestell_dict.get(code, "")
            })

    df_only_bestell = pd.DataFrame(
        only_bestell_rows,
        columns=[
            "Code",
            "Description"
        ]
    )

    # ==================================================
    # 输出表：Only in FG Prüfung
    # ==================================================

    only_fg_rows = []

    for code in fg_dict.keys():

        if code in only_in_fg:

            only_fg_rows.append({
                "Code": code,
                "Description": fg_dict.get(code, "")
            })

    df_only_fg = pd.DataFrame(
        only_fg_rows,
        columns=[
            "Code",
            "Description"
        ]
    )

    # ==================================================
    # 输出表：Matched
    # ==================================================

    matched_rows = []

    for code in bestell_dict.keys():

        if code in matched:

            matched_rows.append({
                "Code": code,
                "Bestellunterlagen_Description": bestell_dict.get(code, ""),
                "FGPruefung_Description": fg_dict.get(code, "")
            })

    df_matched = pd.DataFrame(
        matched_rows,
        columns=[
            "Code",
            "Bestellunterlagen_Description",
            "FGPruefung_Description"
        ]
    )

    # ==================================================
    # 详细对比表
    # ==================================================

    detail_rows = []

    all_codes_ordered = []

    for code in fg_dict.keys():
        if code not in all_codes_ordered:
            all_codes_ordered.append(code)

    for code in bestell_dict.keys():
        if code not in all_codes_ordered:
            all_codes_ordered.append(code)

    for code in all_codes_ordered:

        in_fg = code in fg_codes
        in_bestell = code in bestell_codes

        if in_fg and in_bestell:
            status = "MATCHED"
        elif in_fg and not in_bestell:
            status = "ONLY_IN_FG"
        elif in_bestell and not in_fg:
            status = "ONLY_IN_BESTELLUNTERLAGEN"
        else:
            status = "UNKNOWN"

        detail_rows.append({
            "Code": code,
            "FGPruefung_Description": fg_dict.get(code, ""),
            "Bestellunterlagen_Description": bestell_dict.get(code, ""),
            "In_FGPruefung": "Ja" if in_fg else "Nein",
            "In_Bestellunterlagen": "Ja" if in_bestell else "Nein",
            "Status": status
        })

    df_details = pd.DataFrame(
        detail_rows,
        columns=[
            "Code",
            "FGPruefung_Description",
            "Bestellunterlagen_Description",
            "In_FGPruefung",
            "In_Bestellunterlagen",
            "Status"
        ]
    )

    # ==================================================
    # Summary
    # ==================================================

    df_summary = pd.DataFrame([
        ["Bestellunterlagen文件", bestell_file.name],
        ["FG Prüfung文件", fg_file.name],
        ["Bestellunterlagen Code数量", len(bestell_codes)],
        ["FG Prüfung Code数量", len(fg_codes)],
        ["Matched数量", len(matched)],
        ["Only in Bestellunterlagen数量", len(only_in_bestell)],
        ["Only in FG Prüfung数量", len(only_in_fg)],
        ["生成时间", datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
    ], columns=["项目", "值"])

    # ==================================================
    # 输出 Excel
    # ==================================================

    output_file = Path(
        f"Vergleich_{fg_file.stem}_vs_{bestell_file.stem}.xlsx"
    )

    with pd.ExcelWriter(
        output_file,
        engine="openpyxl"
    ) as writer:

        df_summary.to_excel(
            writer,
            sheet_name="Summary",
            index=False
        )

        df_only_fg.to_excel(
            writer,
            sheet_name="Only_in_FG",
            index=False
        )

        df_only_bestell.to_excel(
            writer,
            sheet_name="Only_in_Bestell",
            index=False
        )

        df_matched.to_excel(
            writer,
            sheet_name="Matched",
            index=False
        )

        df_details.to_excel(
            writer,
            sheet_name="Details",
            index=False
        )

    # ==================================================
    # 控制台输出
    # ==================================================

    print("比较完成")
    print("Excel已生成:", output_file)
    print()
    print("Bestellunterlagen Code数量:", len(bestell_codes))
    print("FG Prüfung Code数量:", len(fg_codes))
    print("Matched:", len(matched))
    print("Only in Bestellunterlagen:", len(only_in_bestell))
    print("Only in FG Prüfung:", len(only_in_fg))


if __name__ == "__main__":

    try:
        main()

    except Exception as e:
        print("程序出错:")
        print(e)
        sys.exit(1)