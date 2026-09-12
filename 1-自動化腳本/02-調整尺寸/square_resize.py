# -*- coding: utf-8 -*-
"""
方形圖片批次縮放工具
- 正方形（或長寬差 <= 5px）的圖片：在原檔旁建立 960x960 的副本（檔名加 _960）
- 其餘圖片：跳出通知顯示尺寸比例不符合

用法：
  1. 把圖片拖曳到 square_resize.bat 上
  2. 或直接執行 square_resize.bat / square_resize.py，會開啟檔案選擇視窗
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

TARGET = 960          # 輸出尺寸
TOLERANCE = 5         # 長寬差在此 px 以內視為正方形
SUFFIX = "_960"
EXTS = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".gif", ".tif", ".tiff"}


def pick_files():
    root = tk.Tk()
    root.withdraw()
    files = filedialog.askopenfilenames(
        title="選擇要處理的圖片",
        filetypes=[("圖片", "*.jpg *.jpeg *.png *.webp *.bmp *.gif *.tif *.tiff"), ("所有檔案", "*.*")],
    )
    root.destroy()
    return list(files)


def output_path(src):
    base, ext = os.path.splitext(src)
    dst = f"{base}{SUFFIX}{ext}"
    n = 2
    while os.path.exists(dst):          # 避免覆蓋既有副本
        dst = f"{base}{SUFFIX}({n}){ext}"
        n += 1
    return dst


def process(src):
    """回傳 (狀態, 訊息)；狀態為 'ok' / 'skip' / 'error'"""
    try:
        with Image.open(src) as im:
            im = ImageOps.exif_transpose(im)   # 依 EXIF 方向轉正，避免手機照片長寬顛倒
            w, h = im.size
            if abs(w - h) > TOLERANCE:
                return "skip", f"{os.path.basename(src)}  ({w}×{h})"

            ext = os.path.splitext(src)[1].lower()
            if ext in (".jpg", ".jpeg", ".bmp") and im.mode not in ("RGB", "L"):
                im = im.convert("RGB")          # JPEG/BMP 不支援透明與調色盤模式

            out = im.resize((TARGET, TARGET), Image.LANCZOS)
            dst = output_path(src)
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
    files = [f for f in sys.argv[1:] if os.path.isfile(f)]
    if not files:
        files = pick_files()
    files = [f for f in files if os.path.splitext(f)[1].lower() in EXTS]
    if not files:
        return

    ok, skip, err = [], [], []
    for f in files:
        status, msg = process(f)
        {"ok": ok, "skip": skip, "error": err}[status].append(msg)
        print(f"[{status}] {msg}")

    root = tk.Tk()
    root.withdraw()
    if skip:
        messagebox.showwarning(
            "尺寸比例不符合",
            "以下圖片不是正方形（長寬差超過 5px），未處理：\n\n" + "\n".join(skip),
        )
    if err:
        messagebox.showerror("處理失敗", "\n".join(err))
    if ok:
        messagebox.showinfo(
            "完成",
            f"已建立 {len(ok)} 張 {TARGET}×{TARGET} 副本：\n\n" + "\n".join(ok),
        )
    root.destroy()


if __name__ == "__main__":
    main()
