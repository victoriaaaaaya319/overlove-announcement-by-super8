"""把威許匯出的「會員基本資料」Excel 轉成只有 Line Uid 的 csv。

從下載資料夾找出最新的真金、白銀檔各一份（依 F 欄「會員級別」判斷），
擷取 O 欄「Line Uid」，存成「真金 YYMMDD.csv」「白銀 YYMMDD.csv」。
"""
import csv
import sys
from datetime import date
from pathlib import Path

import openpyxl

sys.stdout.reconfigure(encoding="utf-8")

SOURCE_DIR = Path.home() / "Downloads"
SOURCE_PATTERN = "會員基本資料*.xlsx"
OUTPUT_DIR = Path(r"C:\Users\USER\Documents\02-癡情工會\3-會員")
LEVELS = {"1": "白銀", "2": "真金"}
LEVEL_COL = 5  # F 欄「會員級別」
UID_COL = 14  # O 欄「Line Uid」


def read_members(path):
    """回傳 (級別名稱, uid 清單)；檔案格式不對或級別混雜時回傳 (None, 原因)。"""
    wb = openpyxl.load_workbook(path, read_only=True)
    try:
        rows = wb.worksheets[0].iter_rows(values_only=True)
        header = next(rows, None)
        if (not header or len(header) <= UID_COL
                or header[LEVEL_COL] != "會員級別" or header[UID_COL] != "Line Uid"):
            return None, "F 欄不是「會員級別」或 O 欄不是「Line Uid」"

        levels = set()
        uids = []
        for row in rows:
            if len(row) <= UID_COL or row[LEVEL_COL] is None:
                continue
            levels.add(str(row[LEVEL_COL]).strip())
            uid = str(row[UID_COL] or "").strip()
            if uid:
                uids.append(uid)
    finally:
        wb.close()

    if len(levels) != 1 or not levels <= LEVELS.keys():
        return None, f"會員級別不是單一的 1 或 2（讀到：{sorted(levels)}）"
    return LEVELS[levels.pop()], list(dict.fromkeys(uids))


def main():
    files = sorted(
        (p for p in SOURCE_DIR.glob(SOURCE_PATTERN) if not p.name.startswith("~$")),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    if not files:
        print(f"在 {SOURCE_DIR} 找不到「會員基本資料」的 Excel 檔")
        return 1

    # 由新到舊找，每個級別只取最新的一份
    found = {}
    for path in files:
        if len(found) == len(LEVELS):
            break
        name, result = read_members(path)
        if name is None:
            print(f"略過 {path.name}：{result}")
        elif name not in found:
            found[name] = (path, result)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    today = date.today().strftime("%y%m%d")
    for name in LEVELS.values():
        if name not in found:
            print(f"找不到{name}的檔案")
            continue
        path, uids = found[name]
        out = OUTPUT_DIR / f"{name} {today}.csv"
        with open(out, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["uid"])
            writer.writerows([uid] for uid in uids)
        print(f"{name}：{path.name} → {out.name}（{len(uids)} 筆 uid）")

    return 0 if len(found) == len(LEVELS) else 1


if __name__ == "__main__":
    sys.exit(main())
