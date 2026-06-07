#!/usr/bin/env python3
"""
apply_styles.py
───────────────
把「標準文件」的樣式定義與頁面設定套用到「目標文件」上：

  1. 用標準文件的 styles.xml 取代目標文件的 styles.xml
  2. 用標準文件的頁面設定（sectPr）取代目標文件的 sectPr
  3. 有套樣式的段落 → 清除直接格式，讓樣式完全接管
  4. 無樣式的段落   → 標紅底色，讓作者手動指定樣式

用法：
    python3 apply_styles.py 標準.docx 目標.docx 輸出.docx
"""

import sys, zipfile, copy
from pathlib import Path
from lxml import etree

# ── Namespaces ────────────────────────────────────────────────────────────────
W  = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
def qn(tag): return '{%s}%s' % (W, tag)

# ── pPr 裡要「保留」的標籤（其餘視為直接格式，刪除）────────────────────────
PPR_KEEP = {
    qn('pStyle'),           # 樣式名稱 ← 最重要，必須保留
    qn('numPr'),            # 清單編號
    qn('pageBreakBefore'),  # 段前分頁
    qn('keepLines'),        # 保持行在同頁
    qn('keepNext'),         # 與下段同頁
    qn('outlineLvl'),       # 大綱層級
    qn('sectPr'),           # 節屬性
    qn('framePr'),          # 框架
    qn('suppressLineNumbers'),
    qn('widowControl'),
}

# ── rPr 裡要「保留」的標籤（其餘視為直接格式，刪除）────────────────────────
RPR_KEEP = {
    qn('b'), qn('bCs'),     # 粗體
    qn('i'), qn('iCs'),     # 斜體
    qn('rStyle'),           # 字元樣式
    qn('lang'),             # 語言
    qn('noProof'),          # 不拼寫檢查
    qn('vertAlign'),        # 上下標（上標/下標需保留）
}

# ── 清除段落的直接格式 ────────────────────────────────────────────────────────
def clean_paragraph(para):
    """
    清除段落 pPr 中的直接格式，以及所有 run 的 rPr 直接格式。
    只保留白名單內的標籤。
    """
    pPr = para.find(qn('pPr'))
    if pPr is not None:
        for child in list(pPr):
            if child.tag not in PPR_KEEP:
                pPr.remove(child)

    for run in para.findall(qn('r')):
        rPr = run.find(qn('rPr'))
        if rPr is not None:
            for child in list(rPr):
                if child.tag not in RPR_KEEP:
                    rPr.remove(child)

# ── 標記無樣式段落（紅底）────────────────────────────────────────────────────
def mark_red(para):
    """
    在段落的 pPr 加入淡紅色背景（作者可清除直接格式後指定樣式）。
    """
    pPr = para.find(qn('pPr'))
    if pPr is None:
        pPr = etree.SubElement(para, qn('pPr'))
        para.insert(0, pPr)

    # 移除既有 shd
    for old_shd in pPr.findall(qn('shd')):
        pPr.remove(old_shd)

    shd = etree.SubElement(pPr, qn('shd'))
    shd.set(qn('val'),   'clear')
    shd.set(qn('color'), 'auto')
    shd.set(qn('fill'),  'FFCCCC')   # 淡紅底色

# ── 取代目標文件的 sectPr ─────────────────────────────────────────────────────
def replace_sectPr(target_body, source_sectPr):
    """把 target 的最後一個 sectPr 換成 source 的 sectPr。"""
    # 移除既有 sectPr（可能在最後一個 p 的 pPr 裡，或直接在 body 下）
    for old in target_body.findall(qn('sectPr')):
        target_body.remove(old)
    for p in target_body.findall(qn('p')):
        pPr = p.find(qn('pPr'))
        if pPr is not None:
            for old in pPr.findall(qn('sectPr')):
                pPr.remove(old)

    # 加入 source 的 sectPr
    target_body.append(copy.deepcopy(source_sectPr))

# ── 主程式 ────────────────────────────────────────────────────────────────────
def apply(source_path: str, target_path: str, output_path: str):
    source_path = Path(source_path)
    target_path = Path(target_path)
    output_path = Path(output_path)

    print(f"標準文件：{source_path.name}")
    print(f"目標文件：{target_path.name}")

    # ── 讀取兩份 docx ────────────────────────────────────────────────────────
    with zipfile.ZipFile(source_path) as zs:
        source_styles_xml = zs.read('word/styles.xml')
        source_doc_root   = etree.fromstring(zs.read('word/document.xml'))

    with zipfile.ZipFile(target_path) as zt:
        target_files = {name: zt.read(name) for name in zt.namelist()}

    # ── 取得 source 的 sectPr ────────────────────────────────────────────────
    source_sectPr = source_doc_root.find('.//' + qn('sectPr'))

    # ── 修改 target 的 document.xml ─────────────────────────────────────────
    target_doc_root = etree.fromstring(target_files['word/document.xml'])
    target_body     = target_doc_root.find(qn('body'))

    paras = target_body.findall('.//' + qn('p'))
    total = len(paras)
    cleaned = 0
    marked  = 0

    for para in paras:
        pPr    = para.find(qn('pPr'))
        pStyle = pPr.find(qn('pStyle')) if pPr is not None else None

        if pStyle is not None:
            # 有套樣式 → 清除直接格式
            clean_paragraph(para)
            cleaned += 1
        else:
            # 無樣式 → 標紅
            mark_red(para)
            marked += 1

    # ── 取代 sectPr ──────────────────────────────────────────────────────────
    if source_sectPr is not None:
        replace_sectPr(target_body, source_sectPr)

    # ── 序列化修改後的 document.xml ──────────────────────────────────────────
    target_files['word/document.xml'] = etree.tostring(
        target_doc_root, xml_declaration=True, encoding='UTF-8', standalone=True
    )

    # ── 用 source 的 styles.xml 取代 target 的 ───────────────────────────────
    target_files['word/styles.xml'] = source_styles_xml

    # ── 寫出新 docx ──────────────────────────────────────────────────────────
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in target_files.items():
            zout.writestr(name, data)

    print(f"\n結果：")
    print(f"  已清除直接格式：{cleaned} 個段落（有樣式）")
    print(f"  標紅待修：      {marked} 個段落（無樣式）")
    print(f"  總段落：        {total}")
    print(f"\n✓ 已輸出：{output_path}")

if __name__ == '__main__':
    if len(sys.argv) != 4:
        print("用法：python3 apply_styles.py 標準.docx 目標.docx 輸出.docx")
        sys.exit(1)
    apply(sys.argv[1], sys.argv[2], sys.argv[3])
