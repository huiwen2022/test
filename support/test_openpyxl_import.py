import os
import sys

# 加入 libs 路徑（確保是 openpyxl 模組所在的那層）
base_path = os.path.dirname(os.path.abspath(__file__))
libs_path = os.path.join(base_path, 'libs')

if libs_path not in sys.path:
    sys.path.insert(0, libs_path)

# 嘗試匯入 openpyxl
try:
    import openpyxl
    print("成功匯入 openpyxl！")
    print("openpyxl 版本：", openpyxl.__version__)
except ModuleNotFoundError as e:
    print("無法匯入 openpyxl：", e)
except Exception as e:
    print("其他錯誤：", e)
