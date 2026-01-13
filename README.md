# 出勤 / 專案加班清洗系統 (Flask 版)

## 專案說明
本專案為內部出勤資料清洗與專案加班處理系統，使用 Python + Flask 製作。  
使用者可透過瀏覽器上傳 **出勤 Excel** 與 **加班 Excel**，程式會自動計算專案加班時段、調整上下班時間，並輸出格式化 Excel 報表供下載。

---

## 系統需求
- 作業系統：Windows 或 Linux（建議 Linux VM）
- Python 版本：>=3.9
- 套件需求：
  ```bash
  pip install flask pandas openpyxl
