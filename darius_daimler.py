import pdfplumber
import pandas as pd
from pathlib import Path
import re

# =====================================
# 输入文件
# =====================================

INPUT_FILE = "FG Prüfung Am Ettersberg.pdf"

pdf_path = Path(INPUT_FILE)

OUTPUT_FILE = pdf_path.with_name(
    pdf_path.stem + "_DaimLER_FGPruefung.xlsx"
)

# =====================================
# Daimler Code格式
# =====================================

CODE_RE = re.compile(
    r'^[A-Z][A-Z0-9]{2,4}$'
)

HEADER_WORDS = [
    "Vor Übergabe an Kunden",
    "Seite",
    "Albert Ziegler GmbH",
    "Bestellung von"
]

END_WORDS = [
    "Gesamtpreis Fahrzeug netto",
    "Kaufpreis"
]


def clean(txt):

    txt = txt.replace("\xa0", " ")
    txt = re.sub(r"\s+", " ", txt)

    return txt.strip()


rows = []

with pdfplumber.open(INPUT_FILE) as pdf:

    # 配置页
    for page in pdf.pages[1:20]:

        text = page.extract_text()

        if not text:
            continue

        lines = text.split("\n")

        for line in lines:

            line = clean(line)

            if not line:
                continue

            if any(h in line for h in HEADER_WORDS):
                continue

            if any(e in line for e in END_WORDS):
                break

            # 去掉章节标题
            if line.startswith("**"):
                continue

            # 识别：
            # F1X S-Fahrerhaus ClassicSpace
            m = re.match(
                r'^([A-Z][A-Z0-9]{2,4})\s+(.+)$',
                line
            )

            if not m:
                continue

            code = m.group(1)
            desc = m.group(2)

            if CODE_RE.match(code):

                rows.append([
                    code,
                    desc
                ])


df = pd.DataFrame(
    rows,
    columns=[
        "Code",
        "Description"
    ]
)


# =====================================
# 输出
# =====================================

df.to_excel(
    OUTPUT_FILE,
    index=False
)

print(df.head())
print()
print("共识别:", len(df))
print("Excel已生成:", OUTPUT_FILE)