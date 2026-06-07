"""
clean_word.py — 移除 .docx 中的直接格式，輸出乾淨的新檔案

保留：Word 內建樣式、Bold、Italic
移除：自訂樣式（→ Normal）、底線、刪除線、文字顏色、螢光筆、行內字體/字號覆蓋
"""

import sys
from pathlib import Path
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Normal 樣式的預設字體
FONT_LATIN = "Times New Roman"   # 英文
FONT_EAST_ASIA = "Noto Sans TC"  # 中文

# 要從 w:rPr 移除的直接格式標籤（保留 w:b / w:bCs / w:i / w:iCs）
REMOVE_TAGS = [
    "w:color",
    "w:highlight",
    "w:u",
    "w:strike",
    "w:rFonts",
    "w:sz",
    "w:szCs",
]


def is_custom_style(style) -> bool:
    return style.element.get(qn("w:customStyle")) == "1"


def clean_run(run) -> int:
    """清除 run 的直接格式與換行符，回傳移除的項目數。"""
    el = run._element
    removed = 0

    # 移除手動換行（soft return）與欄位換行（column break）
    # 保留 page break（w:br w:type="page"）
    for br in el.findall(qn("w:br")):
        br_type = br.get(qn("w:type"), "textWrapping")
        if br_type in ("textWrapping", "column"):
            el.remove(br)
            removed += 1

    rpr = el.find(qn("w:rPr"))
    if rpr is None:
        return removed

    # 移除自訂字元樣式
    if run.style and is_custom_style(run.style):
        rStyle = rpr.find(qn("w:rStyle"))
        if rStyle is not None:
            rpr.remove(rStyle)
            removed += 1

    # 移除各項直接格式
    for tag in REMOVE_TAGS:
        for e in rpr.findall(qn(tag)):
            rpr.remove(e)
            removed += 1

    return removed


def set_normal_style_fonts(doc, latin: str, east_asia: str):
    """把 Normal 樣式的字體設為指定的英文與中文字型。"""
    normal = doc.styles["Normal"]
    # 取得或建立樣式的 w:rPr
    rpr = normal.element.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        normal.element.append(rpr)

    # 移除既有 rFonts
    for el in rpr.findall(qn("w:rFonts")):
        rpr.remove(el)

    # 插入新的 rFonts
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:ascii"), latin)
    rFonts.set(qn("w:hAnsi"), latin)
    rFonts.set(qn("w:eastAsia"), east_asia)
    rFonts.set(qn("w:cs"), east_asia)
    rpr.insert(0, rFonts)

    # 段落對齊設為左右對齊
    normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY


def clean_document(doc_path: Path, out_path: Path) -> dict:
    doc = Document(str(doc_path))

    stats = {"custom_styles": 0, "run_items_removed": 0, "empty_paras_removed": 0}

    for para in doc.paragraphs:
        # 刪除空段落（只有段落符號，無任何文字）
        if para.text.strip() == "":
            para._element.getparent().remove(para._element)
            stats["empty_paras_removed"] += 1
            continue

        if para.style and is_custom_style(para.style):
            para.style = doc.styles["Normal"]
            stats["custom_styles"] += 1

        for run in para.runs:
            stats["run_items_removed"] += clean_run(run)

    set_normal_style_fonts(doc, FONT_LATIN, FONT_EAST_ASIA)
    doc.save(str(out_path))
    return stats


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法：python clean_word.py <檔案.docx>")
        sys.exit(1)

    doc_path = Path(sys.argv[1]).expanduser().resolve()
    if not doc_path.exists():
        print(f"錯誤：找不到檔案 {doc_path}")
        sys.exit(1)

    out_path = doc_path.with_name(doc_path.stem + "_clean.docx")
    stats = clean_document(doc_path, out_path)

    print(f"完成！輸出：{out_path}")
    print(f"  自訂段落樣式重設：{stats['custom_styles']} 個")
    print(f"  直接格式項目移除：{stats['run_items_removed']} 個")
    print(f"  空段落刪除：{stats['empty_paras_removed']} 個")
