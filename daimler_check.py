import pandas as pd
from pathlib import Path


FG_FILE = "FG Prüfung Am Ettersberg_DaimLER_FGPruefung.xlsx"
DAIMLER_FILE = "daimler_Daimler_Bestellunterlagen.xlsx"

OUTPUT_FILE = "FG_vs_Daimler_Vergleich.xlsx"


# ==================================================
# 读取
# ==================================================

df_fg = pd.read_excel(
    FG_FILE,
    dtype=str,
    engine="openpyxl"
)

df_daimler = pd.read_excel(
    DAIMLER_FILE,
    dtype=str,
    engine="openpyxl"
)

# ==================================================
# Code列
# ==================================================

fg_codes = set(
    df_fg["Code"]
    .fillna("")
    .astype(str)
    .str.strip()
)

daimler_codes = set(
    df_daimler["Code"]
    .fillna("")
    .astype(str)
    .str.strip()
)

fg_codes.discard("")
daimler_codes.discard("")

# ==================================================
# 差异
# ==================================================

only_fg = fg_codes - daimler_codes

only_daimler = daimler_codes - fg_codes

matched = fg_codes & daimler_codes

# ==================================================
# 详细检查
# ==================================================

details = []

for _, row in df_fg.iterrows():

    bemerkung = str(
        row.get("Bemerkung von AZG", "")
    ).strip()

    code = str(
        row.get("VCode", "")
    ).strip()

    name = str(
        row.get("Name", "")
    ).strip()

    if not code:
        continue

    if bemerkung in ["OK", "hinzufügen"]:

        if code in daimler_codes:

            status = "PASS"

        else:

            status = "Missing in Daimler"

    elif bemerkung == "Entfall":

        if code in daimler_codes:

            status = "Entfall but exists"

        else:

            status = "PASS"

    else:

        status = "SKIP"

    details.append([
        bemerkung,
        code,
        name,
        status
    ])

# ==================================================
# 输出
# ==================================================

with pd.ExcelWriter(
    OUTPUT_FILE,
    engine="openpyxl"
) as writer:

    pd.DataFrame(
        sorted(list(only_fg)),
        columns=["Code"]
    ).to_excel(
        writer,
        sheet_name="Only_in_FG",
        index=False
    )

    pd.DataFrame(
        sorted(list(only_daimler)),
        columns=["Code"]
    ).to_excel(
        writer,
        sheet_name="Only_in_Daimler",
        index=False
    )

    pd.DataFrame(
        sorted(list(matched)),
        columns=["Code"]
    ).to_excel(
        writer,
        sheet_name="Matched",
        index=False
    )

    pd.DataFrame(
        details,
        columns=[
            "Bemerkung",
            "Code",
            "Name",
            "Status"
        ]
    ).to_excel(
        writer,
        sheet_name="Detailed_Comparison",
        index=False
    )

print("完成")
print("输出文件:", OUTPUT_FILE)

print()
print("Only in FG:", len(only_fg))
print("Only in Daimler:", len(only_daimler))
print("Matched:", len(matched))
