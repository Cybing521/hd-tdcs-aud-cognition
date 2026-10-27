#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""docx 编辑器：段落改写、插入、插表，支持三种输出形态。

形态（mode）：
  clean     —— 直接改写（干净版）
  tracked   —— Word 原生修订痕迹（w:ins / w:del），投稿用
  highlight —— 黄色高亮 + 删除线（给作者/导师看的内部审阅稿）

改写用字符级 diff，只改动真正变化的片段，保留未变部分的原始 run 格式。
表格通过克隆源文档中既有表格的 tblPr / tblGrid / 单元格属性来建，
因此新表与稿件现有三线表样式完全一致。
"""

from __future__ import annotations

import copy
import difflib
import re
import shutil
import zipfile
from pathlib import Path

from lxml import etree

WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XMLNS = "http://www.w3.org/XML/1998/namespace"
W = "{%s}" % WNS


def qn(tag: str) -> str:
    return W + tag


class Docx:
    def __init__(self, path: str | Path):
        self.path = Path(path)
        self.zip_path = self.path
        with zipfile.ZipFile(self.path) as z:
            self.parts = {n: z.read(n) for n in z.namelist()}
        self.root = etree.fromstring(self.parts["word/document.xml"])
        self.body = self.root.find(qn("body"))
        self._id = 1000

    # ---------------------------------------------------------- 基础
    def next_id(self) -> str:
        self._id += 1
        return str(self._id)

    def children(self) -> list:
        return list(self.body)

    def paragraphs(self) -> list:
        return [el for el in self.body if el.tag == qn("p")]

    def tables(self) -> list:
        return [el for el in self.body if el.tag == qn("tbl")]

    @staticmethod
    def text_of(p) -> str:
        parts = []
        for t in p.iter():
            if t.tag == qn("t"):
                parts.append(t.text or "")
            elif t.tag == qn("delText"):
                parts.append(t.text or "")
        return "".join(parts)

    def find_paragraph(self, needle: str, contains: bool = True):
        """按文本定位段落（默认首个命中）。"""
        for p in self.paragraphs():
            txt = self.text_of(p)
            if (needle in txt) if contains else (txt.strip() == needle):
                return p
        raise KeyError(f"未找到段落：{needle!r}")

    def index_of(self, el) -> int:
        return list(self.body).index(el)

    # ---------------------------------------------------------- run 构造
    @staticmethod
    def _rpr_template(p):
        for r in p.findall(qn("r")):
            rpr = r.find(qn("rPr"))
            if rpr is not None:
                return copy.deepcopy(rpr)
        return None

    def _make_run(self, text: str, rpr, mode: str, author: str, date: str):
        r = etree.SubElement(etree.Element(qn("dummy")), qn("r"))
        r = etree.Element(qn("r"))
        if rpr is not None:
            new_rpr = copy.deepcopy(rpr)
            if mode == "highlight":
                for tag in ("highlight", "strike"):
                    for old in new_rpr.findall(qn(tag)):
                        new_rpr.remove(old)
                etree.SubElement(new_rpr, qn("highlight")).set(qn("val"), "yellow")
            r.append(new_rpr)
        t = etree.SubElement(r, qn("t"))
        t.set("{%s}space" % XMLNS, "preserve")
        t.text = text
        if mode == "tracked":
            wrapper = etree.Element(qn("ins"))
            wrapper.set(qn("id"), self.next_id())
            wrapper.set(qn("author"), author)
            wrapper.set(qn("date"), date)
            wrapper.append(r)
            return wrapper
        return r

    def _make_del_run(self, text: str, rpr, author: str, date: str):
        r = etree.Element(qn("r"))
        if rpr is not None:
            r.append(copy.deepcopy(rpr))
        t = etree.SubElement(r, qn("delText"))
        t.set("{%s}space" % XMLNS, "preserve")
        t.text = text
        wrapper = etree.Element(qn("del"))
        wrapper.set(qn("id"), self.next_id())
        wrapper.set(qn("author"), author)
        wrapper.set(qn("date"), date)
        wrapper.append(r)
        return wrapper

    # ---------------------------------------------------------- 段落改写
    def set_paragraph_text(self, p, new_text: str, mode: str, author: str, date: str) -> None:
        old_text = self.text_of(p)
        if old_text == new_text:
            return
        # 逐字符记录所属 run 的 rPr
        rpr_map: list = []
        for r in p.findall(qn("r")):
            rpr = r.find(qn("rPr"))
            txt = "".join((t.text or "") for t in r.findall(qn("t")))
            rpr_map.extend([copy.deepcopy(rpr) if rpr is not None else None] * len(txt))
        default_rpr = self._rpr_template(p)

        def rpr_at(i: int):
            if 0 <= i < len(rpr_map):
                return rpr_map[i]
            return default_rpr

        # 清空原 run（保留 pPr）
        for r in list(p.findall(qn("r"))):
            p.remove(r)
        for wrap in list(p.findall(qn("ins"))) + list(p.findall(qn("del"))):
            p.remove(wrap)

        sm = difflib.SequenceMatcher(None, old_text, new_text, autojunk=False)
        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            if tag == "equal":
                seg = new_text[j1:j2]
                if seg:
                    p.append(self._make_run(seg, rpr_at(i1), "clean", author, date))
            elif tag == "delete":
                seg = old_text[i1:i2]
                if seg:
                    if mode == "tracked":
                        p.append(self._make_del_run(seg, rpr_at(i1), author, date))
                    elif mode == "highlight":
                        r = self._make_run(seg, rpr_at(i1), "highlight", author, date)
                        rpr = r.find(qn("rPr"))
                        if rpr is None:
                            rpr = etree.Element(qn("rPr"))
                            r.insert(0, rpr)
                        etree.SubElement(rpr, qn("strike"))
                        p.append(r)
                    # clean 模式直接丢弃
            elif tag == "insert":
                seg = new_text[j1:j2]
                if seg:
                    p.append(self._make_run(seg, rpr_at(i1 if i1 < len(rpr_map) else 0),
                                            mode, author, date))
            elif tag == "replace":
                old_seg = old_text[i1:i2]
                new_seg = new_text[j1:j2]
                if old_seg:
                    if mode == "tracked":
                        p.append(self._make_del_run(old_seg, rpr_at(i1), author, date))
                    elif mode == "highlight":
                        r = self._make_run(old_seg, rpr_at(i1), "highlight", author, date)
                        rpr = r.find(qn("rPr"))
                        if rpr is None:
                            rpr = etree.Element(qn("rPr"))
                            r.insert(0, rpr)
                        etree.SubElement(rpr, qn("strike"))
                        p.append(r)
                if new_seg:
                    p.append(self._make_run(new_seg, rpr_at(i1 if i1 < len(rpr_map) else 0),
                                            mode, author, date))

    # ---------------------------------------------------------- 插入
    def _new_paragraph(self, text: str, style: str | None, mode: str,
                       author: str, date: str, template=None, inherit_rpr: bool | None = None):
        p = etree.Element(qn("p"))
        if template is not None:
            pPr = template.find(qn("pPr"))
            if pPr is not None:
                p.append(copy.deepcopy(pPr))
        if style:
            pPr = p.find(qn("pPr"))
            if pPr is None:
                pPr = etree.Element(qn("pPr"))
                p.insert(0, pPr)
            ps = pPr.find(qn("pStyle"))
            if ps is None:
                ps = etree.Element(qn("pStyle"))
                pPr.insert(0, ps)
            ps.set(qn("val"), style)
            if inherit_rpr is None:
                inherit_rpr = False        # 标题默认交给样式定义，不继承正文 run 格式
        if inherit_rpr is None:
            inherit_rpr = True
        rpr = self._rpr_template(template) if (template is not None and inherit_rpr) else None
        if text:
            p.append(self._make_run(text, rpr, mode, author, date))
        if mode == "tracked":
            pPr = p.find(qn("pPr"))
            if pPr is None:
                pPr = etree.Element(qn("pPr"))
                p.insert(0, pPr)
            rpr_mark = pPr.find(qn("rPr"))
            if rpr_mark is None:
                rpr_mark = etree.SubElement(pPr, qn("rPr"))
            ins = etree.Element(qn("ins"))
            ins.set(qn("id"), self.next_id())
            ins.set(qn("author"), author)
            ins.set(qn("date"), date)
            rpr_mark.insert(0, ins)
        return p

    def insert_paragraph_after(self, ref, text: str, style: str | None = None,
                               mode: str = "clean", author: str = "Author",
                               date: str = "2026-10-27T00:00:00Z", template=None,
                               inherit_rpr: bool | None = None):
        p = self._new_paragraph(text, style, mode, author, date, template=template,
                                inherit_rpr=inherit_rpr)
        ref.addnext(p)
        return p

    def insert_table_after(self, ref, header: list[str], rows: list[list[str]],
                           mode: str = "clean", author: str = "Author",
                           date: str = "2026-10-27T00:00:00Z", template_table=None,
                           body_template_row=None, header_template_row=None):
        """克隆既有表格的样式，建一张新表。"""
        if template_table is None:
            raise ValueError("需要提供 template_table（源文档中用于克隆样式的表）")
        tbl = copy.deepcopy(template_table)
        for tr in tbl.findall(qn("tr")):
            tbl.remove(tr)
        grid = tbl.find(qn("tblGrid"))
        cols = grid.findall(qn("gridCol"))
        ncol = len(cols)
        if ncol != len(header):
            # 按比例调整列宽
            widths = [int(c.get(qn("w"))) for c in cols]
            total = sum(widths)
            new_w = [int(total * len(header) / ncol)] * len(header)
            for i, c in enumerate(cols):
                if i < len(new_w):
                    c.set(qn("w"), str(new_w[i]))
            extra = cols[len(new_w):]
            for e in extra:
                grid.remove(e)
            while len(grid.findall(qn("gridCol"))) < len(new_w):
                grid.append(copy.deepcopy(cols[0]))
        src_rows = template_table.findall(qn("tr"))
        head_tpl = header_template_row if header_template_row is not None else src_rows[0]
        body_tpl = body_template_row if body_template_row is not None else (src_rows[1] if len(src_rows) > 1 else src_rows[0])

        def fill(tr, values, is_header):
            tcs = tr.findall(qn("tc"))
            for i, tc in enumerate(tcs):
                val = values[i] if i < len(values) else ""
                ps = tc.findall(qn("p"))
                for extra in ps[1:]:
                    tc.remove(extra)
                if not ps:
                    p = etree.SubElement(tc, qn("p"))
                    ps = [p]
                p = ps[0]
                # 保留段落属性，替换文字
                pPr = p.find(qn("pPr"))
                rpr = None
                for r in p.findall(qn("r")):
                    if rpr is None:
                        rpr = r.find(qn("rPr"))
                    p.remove(r)
                for wrap in list(p.findall(qn("ins"))) + list(p.findall(qn("del"))):
                    p.remove(wrap)
                if val:
                    p.append(self._make_run(str(val), rpr, mode if not is_header else mode,
                                            author, date))
            if mode == "tracked":
                trPr = tr.find(qn("trPr"))
                if trPr is None:
                    trPr = etree.Element(qn("trPr"))
                    tr.insert(0, trPr)
                ins = etree.SubElement(trPr, qn("ins"))
                ins.set(qn("id"), self.next_id())
                ins.set(qn("author"), author)
                ins.set(qn("date"), date)
            return tr

        tr_head = fill(copy.deepcopy(head_tpl), header, True)
        tbl.append(tr_head)
        for r in rows:
            tbl.append(fill(copy.deepcopy(body_tpl), r, False))
        ref.addnext(tbl)
        return tbl

    def build_table(self, template_table, header: list[str], rows: list[list[str]],
                    mode: str = "clean", author: str = "Author",
                    date: str = "2026-10-27T00:00:00Z", first_col_left: bool = True,
                    widths: list[float] | None = None):
        """按 template_table 的样式新建一张三线表（顶线 1.5pt / 表头下线 0.75pt / 底线 1.5pt）。"""
        tbl = copy.deepcopy(template_table)
        for tr in tbl.findall(qn("tr")):
            tbl.remove(tr)
        grid = tbl.find(qn("tblGrid"))
        src_cols = grid.findall(qn("gridCol"))
        ncol = len(header)
        total = sum(int(c.get(qn("w"))) for c in src_cols)
        for c in src_cols:
            grid.remove(c)
        if widths and len(widths) == ncol:
            wsum = float(sum(widths))
            widths_px = [int(total * w / wsum) for w in widths]
            # 指定列宽时必须切到固定布局，否则 tblLayout=autofit 会按内容重算列宽
            tblPr = tbl.find(qn("tblPr"))
            if tblPr is not None:
                layout = tblPr.find(qn("tblLayout"))
                if layout is None:
                    layout = etree.SubElement(tblPr, qn("tblLayout"))
                layout.set(qn("type"), "fixed")
                tblW = tblPr.find(qn("tblW"))
                if tblW is not None:
                    tblW.set(qn("type"), "dxa")
                    tblW.set(qn("w"), str(total))
        else:
            widths_px = [int(total / ncol)] * ncol
        widths_px[-1] = total - sum(widths_px[:-1])
        widths = widths_px
        for w in widths:
            gc = etree.SubElement(grid, qn("gridCol"))
            gc.set(qn("w"), str(w))

        src_rows = template_table.findall(qn("tr"))
        head_tpl = src_rows[0]
        body_tpl = src_rows[1] if len(src_rows) > 1 else src_rows[0]

        def make_row(tpl, values, is_header, is_last):
            tr = copy.deepcopy(tpl)
            tcs = tr.findall(qn("tc"))
            # 调整单元格数量
            while len(tcs) > len(values):
                tr.remove(tcs[-1])
                tcs = tr.findall(qn("tc"))
            while len(tcs) < len(values):
                clone = copy.deepcopy(tcs[-1])
                tr.append(clone)
                tcs = tr.findall(qn("tc"))
            for i, (tc, val) in enumerate(zip(tcs, values)):
                tcPr = tc.find(qn("tcPr"))
                if tcPr is not None:
                    tcW = tcPr.find(qn("tcW"))
                    if tcW is not None:
                        tcW.set(qn("w"), str(widths[i]))
                        tcW.set(qn("type"), "dxa")
                    borders = tcPr.find(qn("tcBorders"))
                    if borders is not None:
                        for edge in list(borders):
                            tag = etree.QName(edge).localname
                            edge.set(qn("val"), "nil")
                            edge.set(qn("sz"), "0")
                        if is_header:
                            top = borders.find(qn("top"))
                            if top is not None:
                                top.set(qn("val"), "single")
                                top.set(qn("sz"), "12")
                            bot = borders.find(qn("bottom"))
                            if bot is not None:
                                bot.set(qn("val"), "single")
                                bot.set(qn("sz"), "6")
                        elif is_last:
                            bot = borders.find(qn("bottom"))
                            if bot is not None:
                                bot.set(qn("val"), "single")
                                bot.set(qn("sz"), "12")
                ps = tc.findall(qn("p"))
                for extra in ps[1:]:
                    tc.remove(extra)
                if not ps:
                    p = etree.SubElement(tc, qn("p"))
                    ps = [p]
                p = ps[0]
                pPr = p.find(qn("pPr"))
                if pPr is not None:
                    jc = pPr.find(qn("jc"))
                    if jc is not None:
                        jc.set(qn("val"), "left" if (i == 0 and first_col_left) else "center")
                rpr = None
                for r in p.findall(qn("r")):
                    if rpr is None:
                        rpr = r.find(qn("rPr"))
                    p.remove(r)
                for wrap in list(p.findall(qn("ins"))) + list(p.findall(qn("del"))):
                    p.remove(wrap)
                if val not in (None, ""):
                    p.append(self._make_run(str(val), rpr, mode, author, date))
            if mode == "tracked":
                trPr = tr.find(qn("trPr"))
                if trPr is None:
                    trPr = etree.Element(qn("trPr"))
                    tr.insert(0, trPr)
                ins = etree.SubElement(trPr, qn("ins"))
                ins.set(qn("id"), self.next_id())
                ins.set(qn("author"), author)
                ins.set(qn("date"), date)
            return tr

        tbl.append(make_row(head_tpl, header, True, False))
        for i, r in enumerate(rows):
            tbl.append(make_row(body_tpl, r, False, i == len(rows) - 1))
        return tbl

    # ---------------------------------------------------------- 清空/追加
    def clear_body(self) -> None:
        """清空正文内容，保留 sectPr（页面设置）与样式定义。"""
        sect = self.body.find(qn("sectPr"))
        for el in list(self.body):
            if el is not sect:
                self.body.remove(el)

    def append(self, el) -> None:
        """在 sectPr 之前追加元素。"""
        sect = self.body.find(qn("sectPr"))
        if sect is not None:
            sect.addprevious(el)
        else:
            self.body.append(el)

    # ---------------------------------------------------------- 保存
    def _strip_unused_media(self) -> None:
        """删掉正文已不再引用的图片部件（从原稿模板派生文档时必须做，否则包体虚胖）。"""
        R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
        used = set()
        for el in self.root.iter():
            for attr in ("{%s}embed" % R[1:-1], "{%s}link" % R[1:-1], "{%s}id" % R[1:-1]):
                v = el.get(attr)
                if v:
                    used.add(v)
        rels_name = "word/_rels/document.xml.rels"
        if rels_name not in self.parts:
            return
        rels = etree.fromstring(self.parts[rels_name])
        rel_ns = rels.nsmap.get(None, "http://schemas.openxmlformats.org/package/2006/relationships")
        dropped_targets = []
        for rel in list(rels):
            if rel.get("Id") not in used and (rel.get("TargetMode") or "Internal") != "External":
                target = rel.get("Target", "")
                if target.startswith("media/"):
                    dropped_targets.append("word/" + target)
                rels.remove(rel)
        for name in dropped_targets:
            self.parts.pop(name, None)
        self.parts[rels_name] = etree.tostring(
            rels, xml_declaration=True, encoding="UTF-8", standalone=True)
        # Content_Types 里的 media 声明按实际存在的部件重算
        ct_name = "[Content_Types].xml"
        if ct_name in self.parts:
            ct = etree.fromstring(self.parts[ct_name])
            ct_ns = ct.nsmap.get(None, "http://schemas.openxmlformats.org/package/2006/content-types")
            for override in list(ct):
                if override.get("PartName", "").startswith("/word/media/"):
                    part = override.get("PartName", "").lstrip("/")
                    if part not in self.parts:
                        ct.remove(override)
            self.parts[ct_name] = etree.tostring(
                ct, xml_declaration=True, encoding="UTF-8", standalone=True)

    def save(self, out_path: str | Path) -> None:
        self._strip_unused_media()
        out_path = Path(out_path)
        self.parts["word/document.xml"] = etree.tostring(
            self.root, xml_declaration=True, encoding="UTF-8", standalone=True)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
            for name, data in self.parts.items():
                z.writestr(name, data)


def clone_paragraph_text(p) -> str:
    return Docx.text_of(p)
