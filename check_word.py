"""
check_word.py — 檢查 .docx 格式是否乾淨

允許：Word 內建樣式、Bold、Italic
標記：自訂樣式、底線、刪除線、文字顏色、螢光筆、行內字體/字號覆蓋
"""

import sys
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn
from docx.enum.dml import MSO_THEME_COLOR


def is_custom_style(style) -> bool:
    return style.element.get(qn("w:customStyle")) == "1"


def check_document(docx_path: Path) -> list[dict]:
    doc = Document(str(docx_path))
    issues = []

    for para_idx, para in enumerate(doc.paragraphs):
        para_num = para_idx + 1
        preview = para.text[:60].replace("\n", " ")

        # 段落樣式
        if para.style and is_custom_style(para.style):
            issues.append(
                {
                    "para": para_num,
                    "types": [f"自訂樣式（{para.style.name}）"],
                    "text": preview,
                }
            )

        for run in para.runs:
            run_issues = []

            # 字元樣式
            if run.style and is_custom_style(run.style):
                run_issues.append(f"自訂字元樣式（{run.style.name}）")

            # 底線（True = 有底線；None/False = 繼承或無）
            if run.underline is True:
                run_issues.append("底線")

            # 刪除線
            if run.font.strike:
                run_issues.append("刪除線")

            # 文字顏色（type 不是 None 且不是 AUTO 就算明確設色）
            color = run.font.color
            try:
                if color.type is not None:
                    # WD_COLOR_TYPE.AUTO = 自動色，不算問題
                    from docx.dml.color import RGBColor
                    from docx.oxml.ns import nsmap

                    color_xml = run.font._element.find(
                        qn("w:rPr") + "/" + qn("w:color")
                    )
                    if color_xml is None:
                        color_xml = run._element.find(
                            qn("w:rPr") + "/" + qn("w:color")
                        )
                    if color_xml is not None:
                        val = color_xml.get(qn("w:val"))
                        if val and val.upper() != "AUTO":
                            run_issues.append(f"文字顏色（#{val}）")
            except Exception:
                pass

            # 螢光筆
            if run.font.highlight_color is not None:
                run_issues.append("螢光筆")

            # 行內字體覆蓋
            if run.font.name is not None:
                run_issues.append(f"字體覆蓋（{run.font.name}）")

            # 行內字號覆蓋
            if run.font.size is not None:
                pt = run.font.size.pt
                run_issues.append(f"字號覆蓋（{pt:.1f}pt）")

            if run_issues:
                run_preview = run.text[:60].replace("\n", " ")
                issues.append(
                    {
                        "para": para_num,
                        "types": run_issues,
                        "text": run_preview,
                    }
                )

    return issues


def generate_report(docx_path: Path, issues: list[dict]) -> str:
    lines = [
        "# Word 格式檢查報告",
        "",
        f"**檔案**：`{docx_path.name}`",
        f"**問題數量**：{len(issues)}",
        "",
    ]

    if not issues:
        lines.append("✅ 格式乾淨，無問題。")
        return "\n".join(lines)

    lines += [
        "## 問題清單",
        "",
        "| 段落 | 問題 | 文字片段 |",
        "|:----:|------|---------|",
    ]

    for issue in issues:
        types = "、".join(issue["types"])
        text = issue["text"].replace("|", "\\|")
        lines.append(f"| {issue['para']} | {types} | `{text}` |")

    return "\n".join(lines)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法：python check_word.py <檔案.docx>")
        sys.exit(1)

    docx_path = Path(sys.argv[1]).expanduser().resolve()
    if not docx_path.exists():
        print(f"錯誤：找不到檔案 {docx_path}")
        sys.exit(1)

    issues = check_document(docx_path)
    report = generate_report(docx_path, issues)

    report_path = docx_path.with_name(docx_path.stem + "_report.md")
    report_path.write_text(report, encoding="utf-8")

    print(f"完成！共 {len(issues)} 個問題")
    print(f"報告：{report_path}")
