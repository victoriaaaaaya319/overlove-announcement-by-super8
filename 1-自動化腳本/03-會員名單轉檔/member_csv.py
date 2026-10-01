"""把威許匯出的「會員基本資料」Excel 轉成只有 Line Uid 的 csv。

從下載資料夾找出最新的真金、白銀檔各一份（依 F 欄「會員級別」判斷），
擷取 O 欄「Line Uid」，存成「真金 YYMMDD.csv」「白銀 YYMMDD.csv」。

另外兩個動作（都在轉完 csv、確認沒問題之後才手動執行）：
  move   把這兩份 Excel 搬到 csv 的資料夾，改名成「會員基本資料_真金_YYYYMMDD.xlsx」
  clean  把下載資料夾裡所有會員基本資料 Excel 丟到資源回收筒
"""
import csv
import ctypes
import shutil
import sys
from ctypes import wintypes
from datetime import date
from pathlib import Path

import openpyxl

sys.stdout.reconfigure(encoding="utf-8")

SOURCE_DIR = Path.home() / "Downloads"
SOURCE_PATTERN = "會員基本資料*.xlsx"
ENV_FILE = Path(__file__).with_name(".env")
LEVELS = {"1": "白銀", "2": "真金"}
LEVEL_COL = 5  # F 欄「會員級別」
UID_COL = 14  # O 欄「Line Uid」


def read_output_dir():
    """從 .env 讀 OUTPUT_DIR；沒設定時回傳 None。"""
    if not ENV_FILE.exists():
        return None
    for line in ENV_FILE.read_text(encoding="utf-8-sig").splitlines():
        key, sep, value = line.partition("=")
        if sep and key.strip() == "OUTPUT_DIR":
            value = value.strip().strip('"').strip("'")
            return Path(value) if value else None
    return None


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


def source_files():
    """下載資料夾裡的會員基本資料 Excel，由新到舊。"""
    return sorted(
        (p for p in SOURCE_DIR.glob(SOURCE_PATTERN) if not p.name.startswith("~$")),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )


def find_latest(files):
    """由新到舊找，每個級別只取最新的一份。回傳 {級別名稱: (檔案, uid 清單)}。"""
    found = {}
    for path in files:
        if len(found) == len(LEVELS):
            break
        name, result = read_members(path)
        if name is None:
            print(f"略過 {path.name}：{result}")
        elif name not in found:
            found[name] = (path, result)
    for name in LEVELS.values():
        if name not in found:
            print(f"找不到{name}的檔案")
    return found


def convert(files, output_dir):
    """轉成「真金 YYMMDD.csv」「白銀 YYMMDD.csv」。"""
    found = find_latest(files)
    today = date.today().strftime("%y%m%d")
    for name, (path, uids) in found.items():
        out = output_dir / f"{name} {today}.csv"
        with open(out, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["uid"])
            writer.writerows([uid] for uid in uids)
        print(f"{name}：{path.name} → {out.name}（{len(uids)} 筆 uid）")
    print(f"存在 {output_dir}")
    return 0 if len(found) == len(LEVELS) else 1


def move(files, output_dir):
    """把最新的真金、白銀 Excel 搬成「會員基本資料_真金_YYYYMMDD.xlsx」。"""
    found = find_latest(files)
    today = date.today().strftime("%Y%m%d")
    for name, (path, _) in found.items():
        out = output_dir / f"會員基本資料_{name}_{today}.xlsx"
        out.unlink(missing_ok=True)
        shutil.move(path, out)
        print(f"{name}：{path.name} → {out.name}")
    print(f"搬到 {output_dir}")
    return 0 if len(found) == len(LEVELS) else 1


def recycle(paths):
    """丟到資源回收筒（不是永久刪除）。成功回傳 True。"""

    class SHFILEOPSTRUCTW(ctypes.Structure):
        _fields_ = [
            ("hwnd", wintypes.HWND),
            ("wFunc", wintypes.UINT),
            ("pFrom", wintypes.LPCWSTR),
            ("pTo", wintypes.LPCWSTR),
            ("fFlags", ctypes.c_uint16),
            ("fAnyOperationsAborted", wintypes.BOOL),
            ("hNameMappings", wintypes.LPVOID),
            ("lpszProgressTitle", wintypes.LPCWSTR),
        ]

    FO_DELETE = 3
    FOF_SILENT, FOF_NOCONFIRMATION, FOF_ALLOWUNDO, FOF_NOERRORUI = 0x4, 0x10, 0x40, 0x400
    op = SHFILEOPSTRUCTW(
        wFunc=FO_DELETE,
        pFrom="\0".join(str(p) for p in paths) + "\0\0",
        fFlags=FOF_SILENT | FOF_NOCONFIRMATION | FOF_ALLOWUNDO | FOF_NOERRORUI,
    )
    return ctypes.windll.shell32.SHFileOperationW(ctypes.byref(op)) == 0


def clean(files):
    """把下載資料夾裡所有會員基本資料 Excel 丟到資源回收筒，先列出來讓人確認。"""
    print(f"{SOURCE_DIR} 裡的這些檔案會被丟到資源回收筒：")
    for path in files:
        print(f"  {path.name}")
    if input(f"共 {len(files)} 個，確定要刪除嗎？輸入 y 再按 Enter：").strip().lower() != "y":
        print("沒有刪除任何檔案")
        return 0
    if not recycle(files):
        print("刪除失敗，可能有檔案正被 Excel 開著")
        return 1
    print(f"已把 {len(files)} 個檔案丟到資源回收筒")
    return 0


def main():
    action = sys.argv[1] if len(sys.argv) > 1 else "convert"
    if action not in ("convert", "move", "clean"):
        print(f"不認得的動作：{action}（可用 convert、move、clean）")
        return 1

    files = source_files()
    if not files:
        print(f"在 {SOURCE_DIR} 找不到「會員基本資料」的 Excel 檔")
        return 1
    if action == "clean":
        return clean(files)

    output_dir = read_output_dir()
    if output_dir is None:
        print("還沒設定檔案要存在哪裡：複製 .env.example 改名成 .env，填上 OUTPUT_DIR")
        return 1
    output_dir.mkdir(parents=True, exist_ok=True)
    return move(files, output_dir) if action == "move" else convert(files, output_dir)


if __name__ == "__main__":
    sys.exit(main())
