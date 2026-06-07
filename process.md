# clean-ms-words 工作記錄

## 專案目的

檢查並清除 `.docx` 檔案中的「不乾淨」格式，讓 Word 文件只使用內建樣式與基本字元格式。

---

## 環境

- Python：`/Users/olddrhhtang/Documents/Gemini/gemini_env/bin/python3.12`
- 套件：`python-docx 1.2.0`（2026-05-21 安裝）

---

## 腳本說明

### `check_word.py` — 格式檢查

接受單一 `.docx`，掃描所有段落與 run，輸出 Markdown 報告（`<原檔名>_report.md`）。

**允許**：任何 Word 內建樣式、Bold、Italic

**標記為問題**：

| 問題類型 | 說明 |
|---------|------|
| 自訂樣式 | 段落或字元套用非 Word 內建樣式 |
| 底線 | 直接套用的底線格式 |
| 刪除線 | strikethrough |
| 文字顏色 | 非 AUTO 的明確設色 |
| 螢光筆 | 任何 highlight 顏色 |
| 字體覆蓋 | 行內直接指定 font family |
| 字號覆蓋 | 行內直接指定字號 |

**用法**：
```bash
python check_word.py <檔案.docx>
```

---

### `clean_word.py` — 格式清除

接受單一 `.docx`，清除所有問題格式，輸出 `<原檔名>_clean.docx`（不覆蓋原檔）。

**清除內容**：

| 項目 | 處理方式 |
|------|---------|
| 自訂段落樣式 | 改為 Normal |
| 自訂字元樣式 | 移除 |
| 底線、刪除線、文字顏色、螢光筆 | 移除 |
| 行內字體 / 字號覆蓋 | 移除 |
| 手動換行（Soft Return）| 直接刪除 |
| 欄位換行（Column Break）| 直接刪除 |
| 空段落 | 直接刪除 |

**保留**：Bold、Italic、Page Break

**Normal 樣式預設值**：
- 英文字體：Times New Roman
- 中文字體：Noto Sans TC
- 段落對齊：左右對齊（Justify）

**用法**：
```bash
python clean_word.py <檔案.docx>
```

---

## 工作流程

```
原始 .docx
  ↓
check_word.py → <原檔名>_report.md   （確認問題）
  ↓
clean_word.py → <原檔名>_clean.docx  （清除格式）
  ↓
check_word.py → 0 個問題             （驗證）
```

---

## 開發日誌

| 日期 | 項目 |
|------|------|
| 2026-05-21 | 建立 `check_word.py`、`clean_word.py` |
| 2026-05-21 | 新增：移除手動換行（soft return、column break） |
| 2026-05-21 | 新增：刪除空段落 |
| 2026-05-21 | Normal 樣式設定 Times New Roman / Noto Sans TC / Justify |
