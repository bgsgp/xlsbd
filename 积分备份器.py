import os
import shutil
from datetime import datetime
import time

SRC = r"D:\Program Files\bggp\1\scorer\score.xlsx"
DSTDIR = os.path.join(os.environ['USERPROFILE'], 'OneDrive', '文档', 'xslbd')

print("=" * 40)
print("      小组积分表日期编号备份·丐帮一院™ 荣誉出品")
print("=" * 40)
print()

if not os.path.exists(SRC):
    print("[错误] 找不到源文件（包括隐藏文件）")
    time.sleep(2)
    exit()

print("已找到源文件")

if not os.path.exists(DSTDIR):
    print("目标目录不存在，正在创建...")
    try:
        os.makedirs(DSTDIR)
    except Exception as e:
        print("[错误] 创建目录失败:", e)
        time.sleep(2)
        exit()

today = datetime.now().strftime("%Y%m%d")
base = f"score_{today}"
ext = ".xlsx"

n = 1
while True:
    newfile = os.path.join(DSTDIR, f"{base}_{n}{ext}")
    if not os.path.exists(newfile):
        break
    n += 1

print("正在复制文件...")
try:
    shutil.copy2(SRC, newfile)
except Exception as e:
    print("[错误] 复制失败:", e)
    time.sleep(2)
    exit()

try:
    import ctypes
    FILE_ATTRIBUTE_HIDDEN = 0x02
    FILE_ATTRIBUTE_SYSTEM = 0x04
    attrs = ctypes.windll.kernel32.GetFileAttributesW(newfile)
    if attrs != -1:
        ctypes.windll.kernel32.SetFileAttributesW(
            newfile,
            attrs & ~FILE_ATTRIBUTE_HIDDEN & ~FILE_ATTRIBUTE_SYSTEM
        )
except:
    pass

print("[完成] 已成功备份为：")
print(newfile)
print("原文件保持隐藏，副本可见。")
print()

try:
    import win32com.client
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    wb = excel.Workbooks.Open(newfile)
    try:
        ws = wb.Sheets("LOG")
        ws.Activate()
    except:
        pass
    try:
        excel.DisplayFullScreen = True
    except:
        pass
    try:
        last_row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
        ws.Cells(last_row, 1).Select()
        visible_rows = excel.ActiveWindow.VisibleRange.Rows.Count
        scroll_to = max(1, last_row - visible_rows + 1)
        excel.ActiveWindow.ScrollRow = scroll_to
        excel.ActiveWindow.ScrollColumn = 1
    except:
        pass
    print("已打开文件，全屏显示 LOG 簿，并定位到最新记录。")
    time.sleep(2)
except ImportError:
    os.startfile(newfile)
    print("已用默认程序打开文件。")
    time.sleep(2)
except Exception as e:
    print("[提示] 自动打开全屏时出现问题:", e)
    os.startfile(newfile)
    time.sleep(2)