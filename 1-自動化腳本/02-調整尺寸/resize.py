# -*- coding: utf-8 -*-
"""
圖片批次縮放工具
- 長寬比符合目標（誤差 <= 5px）的圖片：建立指定尺寸的副本，統一存到腳本旁的「輸出」資料夾
- 其餘圖片：跳出通知顯示尺寸比例不符合

用法：
  把圖片拖曳到對應的 .bat 上，或直接雙擊 .bat 開啟檔案選擇視窗
    square_resize.bat    → 960×960（方形）
    portrait_resize.bat  → 960×1553（直式）

  指令列：python resize.py --size 960x1553 圖片1.jpg 圖片2.png ...
"""
import sys
import os
import tkinter as tk
from tkinter import filedialog, messagebox

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

try:
    from PIL import Image, ImageOps
except ImportError:
    tk.Tk().withdraw()
    messagebox.showerror("缺少套件", "找不到 Pillow，請先執行：\npython -m pip install pillow")
    sys.exit(1)

DEFAULT_SIZE = (960, 960)   # 未指定 --size 時的輸出尺寸（寬, 高）
TOLERANCE = 5               # 長寬比換算後相差在此 px 以內就直接縮放
OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "輸出")   # 副本統一存放處
EXTS = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".gif", ".tif", ".tiff"}


def parse_args(argv):
    """回傳 ((寬, 高), 檔案清單)"""
    size = DEFAULT_SIZE
    files = []
    it = iter(argv)
    for a in it:
        if a == "--size":
            w, h = next(it).lower().split("x")
            size = (int(w), int(h))
        elif a.lower().startswith("--size="):
            w, h = a.split("=", 1)[1].lower().split("x")
            size = (int(w), int(h))
        elif os.path.isfile(a):
            files.append(a)
    return size, files


def pick_files(size):
    root = tk.Tk()
    root.withdraw()
    files = filedialog.askopenfilenames(
        title=f"選擇要縮成 {size[0]}×{size[1]} 的圖片",
        filetypes=[("圖片", "*.jpg *.jpeg *.png *.webp *.bmp *.gif *.tif *.tiff"), ("所有檔案", "*.*")],
    )
    root.destroy()
    return list(files)


def output_path(src, size):
    os.makedirs(OUT_DIR, exist_ok=True)
    base, ext = os.path.splitext(os.path.basename(src))
    base = base.replace("_", "")        # Super 8 會擋底線，檔名內的底線全部拿掉
    suffix = f"{size[0]}" if size[0] == size[1] else f"{size[0]}x{size[1]}"
    dst = os.path.join(OUT_DIR, f"{base}{suffix}{ext}")
    n = 2
    while os.path.exists(dst):          # 避免覆蓋既有副本
        dst = os.path.join(OUT_DIR, f"{base}{suffix}-{n}{ext}")
        n += 1
    return dst


def ratio_matches(w, h, size):
    """以原圖寬度換算目標比例下應有的高度，相差 <= TOLERANCE px 視為符合"""
    expected_h = w * size[1] / size[0]
    return abs(h - expected_h) <= TOLERANCE


def process(src, size):
    """回傳 (狀態, 訊息)；狀態為 'ok' / 'skip' / 'error'"""
    try:
        with Image.open(src) as im:
            im = ImageOps.exif_transpose(im)   # 依 EXIF 方向轉正，避免手機照片長寬顛倒
            w, h = im.size
            if not ratio_matches(w, h, size):
                return "skip", f"{os.path.basename(src)}  ({w}×{h})"

            ext = os.path.splitext(src)[1].lower()
            if ext in (".jpg", ".jpeg", ".bmp") and im.mode not in ("RGB", "L"):
                im = im.convert("RGB")          # JPEG/BMP 不支援透明與調色盤模式

            out = im.resize(size, Image.LANCZOS)
            dst = output_path(src, size)
            save_kwargs = {}
            if ext in (".jpg", ".jpeg"):
                save_kwargs.update(quality=95, subsampling=0)
                exif = im.info.get("exif")
                if exif:
                    save_kwargs["exif"] = exif
            out.save(dst, **save_kwargs)
            return "ok", f"{os.path.basename(src)}  ({w}×{h}) → {os.path.basename(dst)}"
    except Exception as e:  # noqa: BLE001
        return "error", f"{os.path.basename(src)}：{e}"


def main():
    size, files = parse_args(sys.argv[1:])
    if not files:
        files = pick_files(size)
    files = [f for f in files if os.path.splitext(f)[1].lower() in EXTS]
    if not files:
        return

    label = f"{size[0]}×{size[1]}"
    ok, skip, err = [], [], []
    for f in files:
        status, msg = process(f, size)
        {"ok": ok, "skip": skip, "error": err}[status].append(msg)
        print(f"[{status}] {msg}")

    root = tk.Tk()
    root.withdraw()
    if skip:
        messagebox.showwarning(
            "尺寸比例不符合",
            f"以下圖片的長寬比不符合 {label}（誤差超過 {TOLERANCE}px），未處理：\n\n" + "\n".join(skip),
        )
    if err:
        messagebox.showerror("處理失敗", "\n".join(err))
    if ok:
        messagebox.showinfo(
            "完成",
            f"已建立 {len(ok)} 張 {label} 副本，存在：\n{OUT_DIR}\n\n" + "\n".join(ok),
        )
    root.destroy()


if __name__ == "__main__":
    main()
