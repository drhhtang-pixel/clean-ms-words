#!/usr/bin/env python3
"""
extract_styles.py
─────────────────
萃取 .docx 的頁面設定與段落樣式，輸出自含式 HTML 報告。

用法：
    python3 extract_styles.py 論文.docx
    python3 extract_styles.py 論文.docx 輸出.html
"""

import sys, zipfile
from pathlib import Path
from lxml import etree

# ── Namespace ─────────────────────────────────────────────────────────────────
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
def qn(tag): return '{%s}%s' % (W, tag)

# ── 單位換算 ──────────────────────────────────────────────────────────────────
def twip_cm(v):  return f"{int(v)/1440*2.54:.2f} cm"  if v else "—"
def half_pt(v):  return f"{int(v)/2:.1f} pt"          if v else "—"
def line_sp(pPr):
    sp = pPr.find(qn('spacing')) if pPr is not None else None
    if sp is None: return "—"
    line = sp.get(qn('line')); rule = sp.get(qn('lineRule'), 'auto')
    if not line: return "—"
    if rule == 'auto':   return f"{int(line)/240:.2f}x"
    elif rule == 'exact': return f"{int(line)/20:.1f}pt（固定）"
    else:                return f"{int(line)/20:.1f}pt（最小）"

def align_label(v):
    return {'both':'兩端','left':'左','right':'右','center':'置中',
            'distribute':'分散'}.get(v or '', v or '—')

# ── 三種狀態統計 ──────────────────────────────────────────────────────────────
# pPr 裡「不算直接格式」的標籤（只記錄結構資訊，不影響外觀）
_PPR_STRUCTURAL = {
    qn('pStyle'), qn('numPr'), qn('sectPr'), qn('pageBreakBefore'),
    qn('keepLines'), qn('keepNext'), qn('outlineLvl'),
    qn('widowControl'), qn('framePr'), qn('suppressLineNumbers'),
}
# rPr 裡「不算直接格式」的標籤
_RPR_STRUCTURAL = {
    qn('b'), qn('bCs'), qn('i'), qn('iCs'),
    qn('rStyle'), qn('lang'), qn('noProof'), qn('vertAlign'),
}

def count_paragraph_states(doc_root):
    """
    回傳 (clean, mixed, no_style) 三種狀態的段落數量。
    - clean   ：有樣式、無直接格式
    - mixed   ：有樣式、又有直接格式蓋在上面
    - no_style：完全沒有套樣式
    """
    clean = mixed = no_style = 0
    for para in doc_root.iter(qn('p')):
        pPr    = para.find(qn('pPr'))
        pStyle = pPr.find(qn('pStyle')) if pPr is not None else None

        if pStyle is None:
            no_style += 1
            continue

        # 檢查 pPr 有沒有非結構性的直接格式
        has_direct = any(
            child.tag not in _PPR_STRUCTURAL
            for child in (pPr if pPr is not None else [])
        )

        # 檢查任一 run 的 rPr 有非結構性的直接格式
        if not has_direct:
            for run in para.findall(qn('r')):
                rPr = run.find(qn('rPr'))
                if rPr is not None and any(
                    child.tag not in _RPR_STRUCTURAL for child in rPr
                ):
                    has_direct = True
                    break

        if has_direct:
            mixed += 1
        else:
            clean += 1

    return clean, mixed, no_style

# ── 萃取頁面設定 ───────────────────────────────────────────────────────────────
def extract_page(doc_root):
    sectPr = doc_root.find('.//' + qn('sectPr'))
    if sectPr is None: return {}
    pgSz  = sectPr.find(qn('pgSz'))
    pgMar = sectPr.find(qn('pgMar'))
    d = {}
    if pgSz  is not None:
        d['紙張寬度'] = twip_cm(pgSz.get(qn('w')))
        d['紙張高度'] = twip_cm(pgSz.get(qn('h')))
        d['方向']    = '直式' if pgSz.get(qn('orient'), 'portrait') == 'portrait' else '橫式'
    if pgMar is not None:
        d['上邊距']   = twip_cm(pgMar.get(qn('top')))
        d['下邊距']   = twip_cm(pgMar.get(qn('bottom')))
        d['左邊距']   = twip_cm(pgMar.get(qn('left')))
        d['右邊距']   = twip_cm(pgMar.get(qn('right')))
        d['頁首距邊'] = twip_cm(pgMar.get(qn('header')))
        d['頁尾距邊'] = twip_cm(pgMar.get(qn('footer')))
        gutter = pgMar.get(qn('gutter'))
        if gutter and gutter != '0': d['裝訂邊'] = twip_cm(gutter)
    return d

# ── 萃取段落樣式 ───────────────────────────────────────────────────────────────
def extract_styles(styles_root, used_ids):
    rows = []
    for style in styles_root.findall(qn('style')):
        if style.get(qn('type')) != 'paragraph': continue
        sid     = style.get(qn('styleId'), '')
        name_el = style.find(qn('name'))
        name    = name_el.get(qn('val'), sid) if name_el is not None else sid
        custom  = style.get(qn('customStyle'), '0') == '1'
        used    = sid in used_ids

        pPr = style.find(qn('pPr'))
        rPr = style.find(qn('rPr'))

        # 段落格式
        jc      = pPr.find(qn('jc'))    if pPr is not None else None
        spacing = pPr.find(qn('spacing')) if pPr is not None else None
        ind     = pPr.find(qn('ind'))   if pPr is not None else None

        s_before = twip_cm(spacing.get(qn('before'))) if spacing is not None else "—"
        s_after  = twip_cm(spacing.get(qn('after')))  if spacing is not None else "—"
        s_line   = line_sp(pPr)
        s_align  = align_label(jc.get(qn('val')) if jc is not None else None)
        s_left   = twip_cm(ind.get(qn('left')))      if ind is not None else "—"
        s_first  = twip_cm(ind.get(qn('firstLine'))) if ind is not None else "—"
        s_hang   = twip_cm(ind.get(qn('hanging')))   if ind is not None else "—"

        # 字元格式
        fonts    = rPr.find(qn('rFonts')) if rPr is not None else None
        sz_el    = rPr.find(qn('sz'))     if rPr is not None else None
        color_el = rPr.find(qn('color'))  if rPr is not None else None
        f_latin  = (fonts.get(qn('ascii')) or fonts.get(qn('asciiTheme'))) if fonts is not None else "—"
        f_ea     = (fonts.get(qn('eastAsia')) or fonts.get(qn('eastAsiaTheme'))) if fonts is not None else "—"
        f_size   = half_pt(sz_el.get(qn('val'))) if sz_el is not None else "—"
        bold     = '✓' if (rPr is not None and rPr.find(qn('b')) is not None) else "—"
        italic   = '✓' if (rPr is not None and rPr.find(qn('i')) is not None) else "—"
        color    = color_el.get(qn('val')) if color_el is not None else None

        rows.append({
            'sid': sid, 'name': name, 'custom': custom, 'used': used,
            'f_latin': f_latin or '—', 'f_ea': f_ea or '—',
            'f_size': f_size, 'bold': bold, 'italic': italic,
            'color': color,
            'align': s_align, 'line': s_line,
            'before': s_before, 'after': s_after,
            'left': s_left, 'first': s_first, 'hang': s_hang,
        })

    # 排序：先顯示實際使用的，再按名稱
    rows.sort(key=lambda r: (not r['used'], r['name']))
    return rows

# ── 產生 HTML ─────────────────────────────────────────────────────────────────
def to_html(docx_name, page, styles, states=None):
    used   = [s for s in styles if s['used']]
    unused = [s for s in styles if not s['used']]

    def color_swatch(hex_val):
        if not hex_val or hex_val.upper() in ('AUTO', 'FFFFFF', 'NONE'): return ''
        return f'<span style="display:inline-block;width:12px;height:12px;background:#{hex_val};border:1px solid #999;vertical-align:middle;margin-right:4px;"></span>#{hex_val}'

    def style_row(s, idx):
        bg = '#fff' if idx % 2 == 0 else '#f9f9f9'
        tag = '自訂' if s['custom'] else '內建'
        tag_color = '#c0392b' if s['custom'] else '#2980b9'
        color_html = color_swatch(s['color']) or '—'
        return f"""
        <tr style="background:{bg}">
          <td><strong>{s['name']}</strong><br><small style="color:#888">{s['sid']}</small></td>
          <td><span style="color:{tag_color};font-size:0.8em;border:1px solid {tag_color};padding:1px 5px;border-radius:3px">{tag}</span></td>
          <td>{s['f_ea']}</td>
          <td>{s['f_latin']}</td>
          <td style="text-align:center">{s['f_size']}</td>
          <td style="text-align:center">{s['bold']}</td>
          <td style="text-align:center">{s['italic']}</td>
          <td style="text-align:center">{s['align']}</td>
          <td style="text-align:center">{s['line']}</td>
          <td style="text-align:center">{s['before']}</td>
          <td style="text-align:center">{s['after']}</td>
          <td style="text-align:center">{s['left']}</td>
          <td style="text-align:center">{s['first']}</td>
          <td>{color_html}</td>
        </tr>"""

    thead = """
    <tr style="background:#2c3e50;color:white">
      <th>樣式名稱</th><th>類型</th>
      <th>字體（中）</th><th>字體（英）</th><th>字號</th>
      <th>粗體</th><th>斜體</th>
      <th>對齊</th><th>行距</th><th>段前</th><th>段後</th>
      <th>左縮排</th><th>首行縮排</th><th>顏色</th>
    </tr>"""

    used_rows   = ''.join(style_row(s, i) for i, s in enumerate(used))
    unused_rows = ''.join(style_row(s, i) for i, s in enumerate(unused))

    page_rows = ''.join(
        f'<tr><td style="color:#555;width:130px">{k}</td><td><strong>{v}</strong></td></tr>'
        for k, v in page.items()
    )

    # ── 三種狀態區塊 ──────────────────────────────────────────────────────────
    if states:
        clean, mixed, no_style = states
        total = clean + mixed + no_style
        pct = lambda n: f"{n/total*100:.0f}%" if total else "0%"
        states_html = f"""
<h2>🔍 文件段落診斷</h2>
<p style="color:#555;margin-bottom:1rem">套用 <code>apply_styles.py</code> 後的預期效果說明：</p>

<div style="display:flex;gap:1rem;flex-wrap:wrap;margin-bottom:1.5rem">

  <div style="flex:1;min-width:260px;border:2px solid #27ae60;border-radius:8px;padding:1rem">
    <div style="font-size:1.8rem;font-weight:bold;color:#27ae60">{clean} <small style="font-size:.9rem">段落</small></div>
    <div style="font-size:.75rem;color:#888;margin-bottom:.5rem">{pct(clean)}（套用後直接完成）</div>
    <div style="font-weight:bold;margin-bottom:.4rem">✅ 狀態一：乾淨</div>
    <div style="font-family:monospace;font-size:.8rem;background:#f4f4f4;padding:.6rem;border-radius:4px;line-height:1.8">
      段落 → 套「論文：內文字」樣式<br>
      &nbsp;&nbsp;&nbsp;└─ 字體、行距、字號 全部<br>
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="color:#27ae60">由樣式決定 ✓</span>
    </div>
    <p style="font-size:.85rem;color:#555;margin-top:.6rem">無須額外處理，套用後外觀即符合標準。</p>
  </div>

  <div style="flex:1;min-width:260px;border:2px solid #e67e22;border-radius:8px;padding:1rem">
    <div style="font-size:1.8rem;font-weight:bold;color:#e67e22">{mixed} <small style="font-size:.9rem">段落</small></div>
    <div style="font-size:.75rem;color:#888;margin-bottom:.5rem">{pct(mixed)}（直接格式會被自動清除）</div>
    <div style="font-weight:bold;margin-bottom:.4rem">⚠️ 狀態二：有樣式＋直接格式蓋住</div>
    <div style="font-family:monospace;font-size:.8rem;background:#f4f4f4;padding:.6rem;border-radius:4px;line-height:1.8">
      段落 → 套「論文：內文字」樣式<br>
      &nbsp;&nbsp;&nbsp;└─ 字體 ✓&nbsp;&nbsp;行距 ✓<br>
      &nbsp;&nbsp;&nbsp;但又直接設定：<br>
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="color:#e67e22">字號 = 12pt ← 蓋掉樣式</span><br>
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="color:#e67e22">顏色 = 黑色 ← 蓋掉樣式</span>
    </div>
    <p style="font-size:.85rem;color:#555;margin-top:.6rem">套用時自動清除直接格式，讓樣式完全接管。</p>
  </div>

  <div style="flex:1;min-width:260px;border:2px solid #e74c3c;border-radius:8px;padding:1rem">
    <div style="font-size:1.8rem;font-weight:bold;color:#e74c3c">{no_style} <small style="font-size:.9rem">段落</small></div>
    <div style="font-size:.75rem;color:#888;margin-bottom:.5rem">{pct(no_style)}（需作者手動指定樣式）</div>
    <div style="font-weight:bold;margin-bottom:.4rem">🔴 狀態三：完全無樣式</div>
    <div style="font-family:monospace;font-size:.8rem;background:#f4f4f4;padding:.6rem;border-radius:4px;line-height:1.8">
      段落 → 沒有套任何樣式<br>
      &nbsp;&nbsp;&nbsp;全靠直接格式：<br>
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="color:#e74c3c">字號 12pt、字體 Times…</span><br>
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="color:#e74c3c">改樣式定義不會影響這段</span>
    </div>
    <p style="font-size:.85rem;color:#555;margin-top:.6rem">套用後標<span style="background:#FFCCCC;padding:1px 4px">紅底色</span>，請作者在 Word 裡指定對應樣式。</p>
  </div>

</div>

<div style="background:#f8f9fa;border-left:4px solid #3498db;padding:.8rem 1rem;margin-bottom:1.5rem;font-size:.9rem">
  <strong>套用後預計結果：</strong>
  {clean} 個段落直接完成 ／
  {mixed} 個段落清除直接格式後完成 ／
  {no_style} 個段落標紅待修
</div>"""
    else:
        states_html = ''

    return f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<title>樣式表 — {docx_name}</title>
<style>
  body {{ font-family: 'Noto Sans TC', sans-serif; margin: 2rem; color: #2c3e50; font-size:14px }}
  h1   {{ font-size: 1.4rem; border-bottom: 3px solid #2c3e50; padding-bottom:.5rem }}
  h2   {{ font-size: 1.1rem; margin-top:2rem; color:#34495e }}
  table {{ border-collapse: collapse; width: 100%; margin-bottom: 1.5rem }}
  th, td {{ border: 1px solid #ddd; padding: 6px 10px; text-align: left; white-space: nowrap }}
  .badge-used   {{ background:#27ae60; color:#fff; padding:2px 8px; border-radius:3px; font-size:.8em }}
  .badge-unused {{ background:#95a5a6; color:#fff; padding:2px 8px; border-radius:3px; font-size:.8em }}
  .summary {{ display:flex; gap:1rem; margin:1rem 0 }}
  .stat {{ background:#ecf0f1; padding:.8rem 1.2rem; border-radius:6px; text-align:center }}
  .stat b {{ display:block; font-size:1.6rem; color:#2c3e50 }}
</style>
</head>
<body>
<h1>📄 樣式表 — {docx_name}</h1>

<div class="summary">
  <div class="stat"><b>{len(used)}</b> 實際使用的樣式</div>
  <div class="stat"><b>{len(unused)}</b> 未使用的樣式</div>
  <div class="stat"><b>{len(used)+len(unused)}</b> 總樣式數</div>
</div>

{states_html}

<h2>📐 頁面設定</h2>
<table style="width:auto">
  <tbody>{page_rows}</tbody>
</table>

<h2>✅ 實際使用的樣式 <span class="badge-used">{len(used)} 個</span></h2>
<div style="overflow-x:auto">
<table>
  <thead>{thead}</thead>
  <tbody>{used_rows}</tbody>
</table>
</div>

<h2>⬜ 未使用的樣式 <span class="badge-unused">{len(unused)} 個</span></h2>
<div style="overflow-x:auto">
<table>
  <thead>{thead}</thead>
  <tbody>{unused_rows}</tbody>
</table>
</div>

<p style="color:#aaa;font-size:.8em;margin-top:2rem">
  由 extract_styles.py 自動產生 ｜ 來源：{docx_name}
</p>
</body>
</html>"""

# ── 主程式 ────────────────────────────────────────────────────────────────────
def extract(docx_path: str, output_html: str = None):
    docx_path = Path(docx_path)
    if output_html is None:
        output_html = docx_path.with_suffix('_styles.html').name

    with zipfile.ZipFile(docx_path) as z:
        styles_root = etree.fromstring(z.read('word/styles.xml'))
        doc_root    = etree.fromstring(z.read('word/document.xml'))

    # 找出文件裡實際使用的樣式 ID
    used_ids = set()
    for el in doc_root.iter(qn('pStyle'), qn('rStyle')):
        v = el.get(qn('val'))
        if v: used_ids.add(v)

    page   = extract_page(doc_root)
    styles = extract_styles(styles_root, used_ids)
    states = count_paragraph_states(doc_root)
    html_content = to_html(docx_path.name, page, styles, states)

    with open(output_html, 'w', encoding='utf-8') as f:
        f.write(html_content)

    used_count = sum(1 for s in styles if s['used'])
    clean, mixed, no_style = states
    print(f"✓ 找到 {len(styles)} 個段落樣式（{used_count} 個實際使用）")
    print(f"✓ 段落診斷：乾淨 {clean} ／ 混合 {mixed} ／ 無樣式 {no_style}")
    print(f"✓ 報告已儲存：{output_html}")
    return output_html

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("用法：python3 extract_styles.py 論文.docx [輸出.html]")
        sys.exit(1)
    out = sys.argv[2] if len(sys.argv) > 2 else None
    extract(sys.argv[1], out)
