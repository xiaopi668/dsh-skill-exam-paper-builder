#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""exam-paper-builder — 一键把 JSON 试卷配置渲染成 试卷 / 标准答题卡 / 参考答案（DOCX，可选 PDF）。

用法：
    python3 scripts/build_exam.py --spec assets/demo_linux_exam.json --out out
    python3 scripts/build_exam.py --demo --out out          # 用内置示例
    python3 scripts/build_exam.py --spec s.json --out out --no-pdf --only paper,card

依赖：python-docx（用 load_workspace_dependencies 返回的 Python 执行）。PDF 依赖 LibreOffice Kit CLI，缺失时自动跳过。
"""
from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
from pathlib import Path

try:
    from docx import Document
    from docx.shared import Pt, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
except ImportError:  # pragma: no cover
    sys.exit("python-docx 不可用：请用 load_workspace_dependencies 返回的 Python 解释器运行本脚本。")

FONT_BODY = "宋体"
FONT_HEAD = "黑体"
FONT_LATIN = "Times New Roman"
FONT_MONO = "Consolas"
CN_NUM = ["一", "二", "三", "四", "五", "六", "七", "八", "九", "十"]

LO_NODE = "/usr/lib/deepseek-harness-desktop/resources/runtime/primary-runtime/dependencies/node/bin/node"
LO_CLI = ("/usr/lib/deepseek-harness-desktop/resources/app.asar.unpacked/dsh/"
          "node_modules/@deepseek-ai/libreoffice-kit/lib/cli.js")

# --------------------------------------------------------------------------- 基础排版工具


def set_font(run, size=10.5, bold=False, ea=FONT_BODY, latin=FONT_LATIN, color=None):
    run.font.name = latin
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), latin)
    rFonts.set(qn("w:hAnsi"), latin)
    rFonts.set(qn("w:eastAsia"), ea)
    rFonts.set(qn("w:cs"), latin)
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)


def para(doc, text="", size=10.5, bold=False, ea=FONT_BODY, latin=FONT_LATIN, align="left",
         before=0, after=2, line=1.15, indent=0.0, keep=False, color=None):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line
    pf.left_indent = Cm(indent)
    pf.first_line_indent = Cm(0)
    pf.keep_with_next = keep
    pf.alignment = {"left": WD_ALIGN_PARAGRAPH.LEFT, "center": WD_ALIGN_PARAGRAPH.CENTER,
                    "right": WD_ALIGN_PARAGRAPH.RIGHT, "just": WD_ALIGN_PARAGRAPH.JUSTIFY}[align]
    if text:
        run = p.add_run(text)
        set_font(run, size, bold, ea, latin, color)
    return p


def table_borders(table, color="000000", sz=8, inside=True, val="single"):
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    edges = ["top", "left", "bottom", "right"] + (["insideH", "insideV"] if inside else [])
    for edge in edges:
        el = OxmlElement("w:" + edge)
        el.set(qn("w:val"), val)
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    tblPr.append(borders)


def cell_borders(cell, color="000000", sz=12, val="single"):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        el = OxmlElement("w:" + side)
        el.set(qn("w:val"), val)
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    tcPr.append(borders)


def cell_margins(table, top=20, bottom=20, left=40, right=40):
    tblPr = table._tbl.tblPr
    mar = OxmlElement("w:tblCellMar")
    for side, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        el = OxmlElement("w:" + side)
        el.set(qn("w:w"), str(val))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tblPr.append(mar)


def table_width(table, widths):
    table.autofit = False
    tblPr = table._tbl.tblPr
    tw = tblPr.find(qn("w:tblW"))
    if tw is None:
        tw = OxmlElement("w:tblW")
        tblPr.append(tw)
    tw.set(qn("w:type"), "dxa")
    tw.set(qn("w:w"), str(int(sum(widths) * 567)))
    layout = tblPr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tblPr.append(layout)
    layout.set(qn("w:type"), "fixed")
    for idx, w in enumerate(widths):
        table.columns[idx].width = Cm(w)
    for row in table.rows:
        for idx, w in enumerate(widths):
            if idx < len(row.cells):
                row.cells[idx].width = Cm(w)


def table_indent(table, cm):
    tblPr = table._tbl.tblPr
    ind = OxmlElement("w:tblInd")
    ind.set(qn("w:type"), "dxa")
    ind.set(qn("w:w"), str(int(cm * 567)))
    tblPr.append(ind)


def row_height(row, cm, exact=True):
    trPr = row._tr.get_or_add_trPr()
    h = OxmlElement("w:trHeight")
    h.set(qn("w:val"), str(int(cm * 567)))
    h.set(qn("w:hRule"), "exact" if exact else "atLeast")
    trPr.append(h)


def no_split(row):
    trPr = row._tr.get_or_add_trPr()
    trPr.insert(0, OxmlElement("w:cantSplit"))


def cell_shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


def cell_para(cell, text="", size=10.5, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT,
              color=None, before=0, after=0, line=1.1, ea=FONT_BODY, latin=FONT_LATIN):
    p = cell.paragraphs[0] if not cell.paragraphs[0].text else cell.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line
    pf.first_line_indent = Cm(0)
    p.alignment = align
    if text:
        run = p.add_run(text)
        set_font(run, size, bold, ea, latin, color)
    return p


def add_field(paragraph, instr, placeholder="1", size=9):
    run = paragraph.add_run()
    set_font(run, size)
    r = run._r
    f1 = OxmlElement("w:fldChar"); f1.set(qn("w:fldCharType"), "begin")
    it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = " " + instr + " "
    f2 = OxmlElement("w:fldChar"); f2.set(qn("w:fldCharType"), "separate")
    t = OxmlElement("w:t"); t.text = placeholder
    f3 = OxmlElement("w:fldChar"); f3.set(qn("w:fldCharType"), "end")
    for e in (f1, it, f2, t, f3):
        r.append(e)


def base_document(margins=(2.0, 1.9, 2.4, 2.4), size=10.5):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.top_margin, sec.bottom_margin, sec.left_margin, sec.right_margin = [Cm(v) for v in margins]
    style = doc.styles["Normal"]
    style.font.name = FONT_LATIN
    style.font.size = Pt(size)
    style.element.rPr.rFonts.set(qn("w:eastAsia"), FONT_BODY)
    return doc, sec


# --------------------------------------------------------------------------- 试卷


def _option_label(text, idx):
    text = str(text).strip()
    if re.match(r"^[A-DＡ-Ｄ]\s*[．.、)）]", text):
        return text
    return "ABCD"[idx % 4] + "．" + text


def build_paper(spec, path, page_count):
    meta = spec["meta"]
    doc, sec = base_document(margins=(2.0, 1.9, 2.4, 2.4))
    text_w = 16.2

    hdr = sec.header
    hdr.paragraphs[0].text = ""
    htbl = hdr.add_table(rows=1, cols=2, width=Cm(text_w))
    htbl.alignment = WD_TABLE_ALIGNMENT.LEFT
    tblPr = htbl._tbl.tblPr
    bd = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "right", "insideH", "insideV"):
        el = OxmlElement("w:" + edge); el.set(qn("w:val"), "none"); bd.append(el)
    hb = OxmlElement("w:bottom")
    hb.set(qn("w:val"), "single"); hb.set(qn("w:sz"), "6")
    hb.set(qn("w:space"), "0"); hb.set(qn("w:color"), "000000")
    bd.append(hb)
    tblPr.append(bd)
    cell_margins(htbl, left=0, right=0, top=0, bottom=20)
    table_width(htbl, [text_w / 2, text_w / 2])
    lp = htbl.rows[0].cells[0].paragraphs[0]
    lp.paragraph_format.space_after = Pt(0); lp.paragraph_format.line_spacing = 1.0
    set_font(lp.add_run(meta.get("confidential", "绝密★启用前")), 9, True)
    rp = htbl.rows[0].cells[1].paragraphs[0]
    rp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    rp.paragraph_format.space_after = Pt(0); rp.paragraph_format.line_spacing = 1.0
    set_font(rp.add_run("%s：%s　｜　科目代码：%s" % (
        meta.get("card_type_label", "试卷类型"), meta.get("card_type", "A"),
        meta.get("paper_code", "EXAM-2026"))), 9)
    tail = hdr.add_paragraph()
    tail.paragraph_format.space_before = Pt(0)
    tail.paragraph_format.space_after = Pt(0)
    tail.paragraph_format.line_spacing = Pt(1)

    para(doc, "", size=4, after=0, line=0.6)
    para(doc, meta.get("exam_title", "2026年普通高等学校招生全国统一考试"), size=19, bold=True,
         ea=FONT_HEAD, align="center", before=2, after=2, line=1.1)
    para(doc, "%s（%s）" % (meta.get("subject", ""), meta.get("paper_type", "")), size=15.5, bold=True,
         ea=FONT_HEAD, align="center", before=0, after=4, line=1.1)
    para(doc, "本试卷共 %d 页，%d 小题，满分 %s 分，考试用时 %s 分钟" % (
        page_count, meta.get("question_count", 0), meta.get("total_score", ""),
        meta.get("duration_min", 120)), size=10, bold=True, align="center", after=6)
    para(doc, "姓名：__________　　准考证号：□□□□□□□□□□　　考场号：______　　座位号：______",
         size=10, align="center", after=6)

    notice = spec.get("notice") or []
    if notice:
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        table_borders(tbl, "000000", 6, inside=False)
        cell_margins(tbl, left=110, right=110, top=60, bottom=60)
        table_width(tbl, [text_w])
        cell = tbl.rows[0].cells[0]
        first = True
        for i, line in enumerate(notice):
            p = cell.paragraphs[0] if first else cell.add_paragraph()
            first = False
            pf = p.paragraph_format
            pf.space_before = Pt(0); pf.space_after = Pt(1); pf.line_spacing = 1.1
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            run = p.add_run(line)
            set_font(run, 9, i == 0)

    for si, section in enumerate(spec["sections"]):
        head = "%s、%s：%s" % (CN_NUM[si] if si < len(CN_NUM) else str(si + 1),
                              section.get("heading", ""), section.get("instruction", ""))
        para(doc, head, size=10.5, bold=True, before=10, after=6, line=1.25, align="just", keep=True)

        qtype = section.get("type", "short")
        for qi, q in enumerate(section["questions"]):
            if qtype == "choice":
                para(doc, q["stem"], size=10.5, before=8 if qi == 0 else 6, after=0,
                     line=1.3, align="just", keep=True)
                opts = q["options"]
                tbl = doc.add_table(rows=2, cols=2)
                tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
                table_borders(tbl, "000000", 4, inside=False, val="none")
                cell_margins(tbl)
                table_width(tbl, [text_w / 2, text_w / 2])
                no_split(tbl.rows[0]); no_split(tbl.rows[1])
                for k in range(4):
                    r_i, c_i = divmod(k, 2)
                    cell = tbl.rows[r_i].cells[c_i]
                    txt = _option_label(opts[k], k) if k < len(opts) else ""
                    p = cell_para(cell, txt, 10.5, before=1.5, after=4)
                    if r_i == 0:
                        p.paragraph_format.keep_with_next = True
            else:
                para(doc, q["stem"], size=10.5,
                     before=(9 if qtype == "fill" else 10 if qtype == "short" else 14),
                     after=(9 if qtype == "fill" else 10 if qtype == "short" else 14),
                     line=1.25, align="just")

    ftr = sec.footer.paragraphs[0]
    ftr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_font(ftr.add_run("第 "), 9)
    add_field(ftr, "PAGE", "1")
    set_font(ftr.add_run(" 页（共 "), 9)
    add_field(ftr, "NUMPAGES", str(page_count))
    set_font(ftr.add_run(" 页）"), 9)

    doc.save(path)
    return path


# --------------------------------------------------------------------------- 答题卡


def _q_number(text, fallback):
    m = re.match(r"^\s*(\d+)", str(text))
    return m.group(1) if m else str(fallback)


def _band(container, cells=44, height=0.42, total_w=21.0, indent=-1.6, color="000000"):
    tbl = container.add_table(rows=1, cols=cells, width=Cm(total_w))
    table_borders(tbl, "FFFFFF", 2, inside=False, val="none")
    table_indent(tbl, indent)
    w = total_w / cells
    table_width(tbl, [w] * cells)
    cell_margins(tbl, 0, 0, 0, 0)
    row = tbl.rows[0]
    row_height(row, height)
    no_split(row)
    for i, c in enumerate(row.cells):
        cell_shade(c, color if i % 2 == 0 else "FFFFFF")
        cell_para(c, "", 2)


def _ruled(cell, count, first_text="", line_h=0.92, size=9, bold=False):
    for i in range(count):
        p = cell.paragraphs[0] if (i == 0 and not cell.paragraphs[0].text) else cell.add_paragraph()
        pf = p.paragraph_format
        pf.space_before = Pt(0); pf.space_after = Pt(0)
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Cm(line_h)
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "dotted")
        bottom.set(qn("w:sz"), "4")
        bottom.set(qn("w:space"), "1")
        bottom.set(qn("w:color"), "808080" if i % 2 == 0 else "7F7F7F")
        pBdr.append(bottom)
        pPr.append(pBdr)
        if i == 0 and first_text:
            set_font(p.add_run(first_text), size, bold)


def build_card(spec, path):
    meta = spec["meta"]
    card = spec.get("card", {})
    theme = card.get("theme", "C00000")
    doc, sec = base_document(margins=(1.5, 1.3, 1.6, 1.6), size=9)
    sec.header_distance = Cm(0.5)
    sec.footer_distance = Cm(0.4)
    text_w = 17.8

    hdr = sec.header
    hdr.paragraphs[0].text = ""
    hdr.paragraphs[0].paragraph_format.space_after = Pt(0)
    hdr.paragraphs[0].paragraph_format.line_spacing = Pt(1)
    _band(hdr)
    hp = hdr.add_paragraph()
    hp.paragraph_format.space_before = Pt(0); hp.paragraph_format.space_after = Pt(0)
    hp.paragraph_format.line_spacing = Pt(2)

    ftr = sec.footer
    ftr.paragraphs[0].text = ""
    ftr.paragraphs[0].paragraph_format.space_after = Pt(0)
    ftr.paragraphs[0].paragraph_format.line_spacing = Pt(1)
    tip = ftr.add_table(rows=1, cols=1, width=Cm(text_w))
    table_borders(tip, theme, 8, inside=False)
    table_width(tip, [text_w])
    cell_margins(tip, 30, 30, 40, 40)
    cell_para(tip.rows[0].cells[0], card.get("tip", "请在各题目的答题区域内作答，超出黑色矩形边框限定区域的答案无效"),
              7.5, True, WD_ALIGN_PARAGRAPH.CENTER, theme)
    fp = ftr.add_paragraph()
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fp.paragraph_format.space_before = Pt(1)
    fp.paragraph_format.space_after = Pt(0)
    fp.paragraph_format.line_spacing = 1.0
    set_font(fp.add_run("%s　第 " % meta.get("subject", "")), 8)
    add_field(fp, "PAGE", "1", 8)
    set_font(fp.add_run(" 页　共 "), 8)
    add_field(fp, "NUMPAGES", "1", 8)
    set_font(fp.add_run(" 页"), 8)

    para(doc, meta.get("exam_title", ""), size=10, align="center", color=theme, before=4, line=1.1)
    para(doc, "%s（%s）" % (meta.get("subject", ""), meta.get("paper_type", "")), size=13.5, bold=True,
         ea=FONT_HEAD, align="center", line=1.1)
    para(doc, "答　题　卡", size=18, bold=True, ea=FONT_HEAD, align="center", after=4, line=1.1)
    para(doc, "卡型：%s　　科目代码：%s" % (meta.get("card_type", "A"), meta.get("paper_code", "")),
         size=8.5, align="center", color=theme, after=6)

    info = doc.add_table(rows=1, cols=2)
    info.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_borders(info, theme, 8, inside=False)
    table_width(info, [11.3, 6.5])
    cell_margins(info, 40, 40, 80, 80)
    no_split(info.rows[0])
    left, right = info.rows[0].cells
    p = cell_para(left, "", 9)
    for txt in ("姓　　名：", "_________________", "　　考场号：", "_________"):
        set_font(p.add_run(txt), 9)
    p = cell_para(left, "", 9, before=6)
    for txt in ("座位号：", "__________", "　　准考证号："):
        set_font(p.add_run(txt), 9)
    grid = left.add_table(rows=1, cols=12)
    table_borders(grid, "000000", 6, inside=True)
    table_width(grid, [0.55] * 12)
    cell_margins(grid, 0, 0, 0, 0)
    row_height(grid.rows[0], 0.75)
    for c in grid.rows[0].cells:
        cell_para(c, "", 8)
    p = cell_para(left, "", 6, before=4)
    set_font(p.add_run(card.get("absent_note", "缺考标记（由监考员填涂）：□　　正确填涂：■　　错误填涂：⊠　⊞")), 7.5)
    cell_para(right, "贴　条　形　码　区", 9, True, WD_ALIGN_PARAGRAPH.CENTER, theme)
    cell_para(right, "（正面朝上，切勿贴出虚线框）", 7, False, WD_ALIGN_PARAGRAPH.CENTER, theme, before=2)
    cell_borders(right, theme, 6, val="dashed")
    para(doc, "", size=2, after=2, line=0.6)

    notes = spec.get("card_notice") or spec.get("notice") or []
    if notes:
        note = doc.add_table(rows=1, cols=2)
        note.alignment = WD_TABLE_ALIGNMENT.CENTER
        table_borders(note, theme, 8, inside=False)
        table_width(note, [2.0, text_w - 2.0])
        cell_margins(note, 40, 40, 60, 80)
        no_split(note.rows[0])
        nleft, nright = note.rows[0].cells
        cell_para(nleft, "注\n意\n事\n项", 9, True, WD_ALIGN_PARAGRAPH.CENTER, theme)
        for i, line in enumerate(notes):
            cell_para(nright, line, 7.5, before=0 if i == 0 else 1, line=1.1)
        para(doc, "", size=2, after=4, line=0.6)

    sections = spec["sections"]
    choice = next((s for s in sections if s.get("type") == "choice"), None)
    if choice:
        para(doc, choice.get("card_heading", "一、选择题"), 8.5, True, color=theme, after=2, keep=True)
        per_row = int(card.get("choice_per_row", 5))
        qs = choice["questions"]
        rows = (len(qs) + per_row - 1) // per_row
        sel = doc.add_table(rows=rows + 1, cols=per_row)
        sel.alignment = WD_TABLE_ALIGNMENT.CENTER
        table_borders(sel, "000000", 10, inside=False)
        table_width(sel, [text_w / per_row] * per_row)
        cell_margins(sel, 30, 30, 40, 40)
        head = sel.rows[0].cells[0]
        for c in sel.rows[0].cells[1:]:
            head = head.merge(c)
        cell_para(head, card.get("choice_hint", "选择题填涂区　（请用 2B 铅笔将所选选项涂黑）"),
                   7.5, False, WD_ALIGN_PARAGRAPH.LEFT, theme)
        for i, q in enumerate(qs):
            r_i, c_i = 1 + i // per_row, i % per_row
            cell_para(sel.rows[r_i].cells[c_i], "%s　[A]　[B]　[C]　[D]" % _q_number(q["stem"], i + 1),
                      8, False)
        para(doc, "", size=2, after=2, line=0.6)

    lines_conf = card.get("lines", {})
    for section in sections:
        if section.get("type") == "choice":
            continue
        title = section.get("card_heading") or section.get("heading", "")
        score = section.get("card_score") or ""
        if score:
            title = "%s（%s）" % (title, score)
        para(doc, title, 9, True, color=theme, before=4, after=3, keep=True)
        for q in section["questions"]:
            no = _q_number(q["stem"], "")
            label = "%s．（%s）" % (no, q.get("score", section.get("points_each", "")))
            tbl = doc.add_table(rows=1, cols=1)
            tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
            table_borders(tbl, "000000", 12, inside=False)
            table_width(tbl, [text_w])
            cell_margins(tbl, 40, 40, 70, 70)
            no_split(tbl.rows[0])
            n = int(q.get("lines") or lines_conf.get(section.get("type"), 6))
            _ruled(tbl.rows[0].cells[0], n, first_text=label)
            para(doc, "", size=3, after=2, line=0.6)

    doc.save(path)
    return path


# --------------------------------------------------------------------------- 参考答案


def build_key(spec, path):
    meta = spec["meta"]
    key = spec.get("answer_key", {})
    doc, sec = base_document(margins=(2.0, 1.9, 2.4, 2.4))
    text_w = 16.2

    para(doc, meta.get("key_header", "内部资料　评卷与讲评使用"), size=9, bold=True, after=2)
    para(doc, meta.get("exam_title", ""), size=16, bold=True, ea=FONT_HEAD, align="center", after=2, line=1.1)
    para(doc, "%s（%s）" % (meta.get("subject", ""), meta.get("paper_type", "")), size=13, bold=True,
         ea=FONT_HEAD, align="center", after=2, line=1.1)
    para(doc, "参考答案与评分标准", size=12, bold=True, ea=FONT_HEAD, align="center", after=10, line=1.1)

    choice = key.get("choice")
    if choice:
        para(doc, key.get("choice_heading", "一、选择题"), size=11, bold=True, ea=FONT_HEAD,
             before=10, after=4, keep=True)
        n = len(choice)
        half = (n + 1) // 2
        rows = half + 1
        tbl = doc.add_table(rows=rows, cols=4)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        table_borders(tbl, "000000", 4, inside=True)
        cell_margins(tbl, 60, 60, 30, 30)
        table_width(tbl, [3.2, 4.9, 3.2, 4.9])
        for i, h in enumerate(["题号", "答案", "题号", "答案"]):
            cell_para(tbl.rows[0].cells[i], h, 10, True, WD_ALIGN_PARAGRAPH.CENTER)
        for r in range(1, rows):
            for c in (0, 1):
                idx = (r - 1) + c * half
                if idx < n:
                    cell_para(tbl.rows[r].cells[c * 2], str(idx + 1), 10, False, WD_ALIGN_PARAGRAPH.CENTER)
                    cell_para(tbl.rows[r].cells[c * 2 + 1], str(choice[idx]), 10, False, WD_ALIGN_PARAGRAPH.CENTER)
        if key.get("choice_note"):
            para(doc, key["choice_note"], size=9.5, before=4, after=2, line=1.2)

    for block in key.get("sections", []):
        para(doc, block.get("heading", ""), size=11, bold=True, ea=FONT_HEAD, before=10, after=4, keep=True)
        for item in block.get("items", []):
            p = doc.add_paragraph()
            pf = p.paragraph_format
            pf.space_before = Pt(4); pf.space_after = Pt(2); pf.line_spacing = 1.25
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            if item.get("no"):
                set_font(p.add_run(item["no"]), 10.5, True)
            if item.get("lead"):
                set_font(p.add_run(item["lead"]), 10.5, False)
            if item.get("text"):
                set_font(p.add_run(item["text"]), 10.5, False)
            if item.get("code"):
                tbl = doc.add_table(rows=1, cols=1)
                tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
                table_borders(tbl, "FFFFFF", 2, inside=False, val="none")
                cell_margins(tbl, 100, 100, 40, 40)
                table_width(tbl, [text_w])
                no_split(tbl.rows[0])
                cell = tbl.rows[0].cells[0]
                cell_shade(cell, "F5F5F5")
                for i, line in enumerate(item["code"].strip("\n").split("\n")):
                    cell_para(cell, line if line.strip() else " ", 8.5, ea="Consolas", latin="Consolas",
                              before=0, after=0, line=1.0)
                para(doc, "", size=2, after=4, line=0.6)
            if item.get("points"):
                q = doc.add_paragraph()
                q.paragraph_format.space_before = Pt(0)
                q.paragraph_format.space_after = Pt(5)
                q.paragraph_format.line_spacing = 1.2
                set_font(q.add_run(item["points"]), 9.5)

    doc.save(path)
    return path


# --------------------------------------------------------------------------- PDF / 驱动


def export_pdf(docx_path, pdf_path, node=LO_NODE, cli=LO_CLI):
    if not (os.path.exists(node) and os.path.exists(cli)):
        return False
    if os.path.exists(pdf_path):
        os.remove(pdf_path)
    r = subprocess.run([node, cli, "convert", "--input", str(docx_path), "--output", str(pdf_path)],
                       capture_output=True, text=True)
    return r.returncode == 0 and os.path.exists(pdf_path)


def pdf_pages(pdf_path):
    data = Path(pdf_path).read_bytes()
    return len(re.findall(rb"/Type\s*/Page[^s]", data))


def load_spec(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def main(argv=None):
    ap = argparse.ArgumentParser(description="一键生成试卷 / 答题卡 / 参考答案")
    ap.add_argument("--spec", help="试卷配置 JSON")
    ap.add_argument("--demo", action="store_true", help="使用内置示例配置")
    ap.add_argument("--out", default="exam_out", help="输出目录")
    ap.add_argument("--no-pdf", action="store_true", help="只生成 DOCX")
    ap.add_argument("--only", default="paper,card,key", help="paper,card,key 的子集")
    args = ap.parse_args(argv)

    if args.demo and not args.spec:
        args.spec = str(Path(__file__).resolve().parent.parent / "assets" / "demo_linux_exam.json")
    if not args.spec:
        ap.error("需要 --spec 或 --demo")

    spec = load_spec(args.spec)
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    meta = spec.setdefault("meta", {})
    stem = meta.get("file_stem", "试卷")
    want = {s.strip() for s in args.only.split(",") if s.strip()}

    produced = []
    if "paper" in want:
        docx_path = out / (stem + ".docx")
        build_paper(spec, docx_path, int(meta.get("pages", 4) or 4))
        if not args.no_pdf:
            pdf_path = out / (stem + ".pdf")
            if export_pdf(docx_path, pdf_path):
                pages = pdf_pages(pdf_path)
                for _ in range(3):
                    if pages == int(meta.get("pages", pages)):
                        break
                    meta["pages"] = pages
                    build_paper(spec, docx_path, pages)
                    export_pdf(docx_path, pdf_path)
                    new_pages = pdf_pages(pdf_path)
                    if new_pages == pages:
                        break
                    pages = new_pages
                produced.append(pdf_path)
        produced.append(docx_path)

    if "card" in want:
        docx_path = out / (meta.get("card_stem", "答题卡") + ".docx")
        build_card(spec, docx_path)
        produced.append(docx_path)
        if not args.no_pdf:
            pdf_path = out / (meta.get("card_stem", "答题卡") + ".pdf")
            if export_pdf(docx_path, pdf_path):
                produced.append(pdf_path)

    if "key" in want:
        docx_path = out / (meta.get("key_stem", "参考答案与评分标准") + ".docx")
        build_key(spec, docx_path)
        produced.append(docx_path)
        if not args.no_pdf:
            pdf_path = out / (meta.get("key_stem", "参考答案与评分标准") + ".pdf")
            if export_pdf(docx_path, pdf_path):
                produced.append(pdf_path)

    for p in produced:
        print(p)
    return 0


if __name__ == "__main__":
    sys.exit(main())
