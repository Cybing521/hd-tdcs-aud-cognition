#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""列出 docx 正文段落（含样式、表格与图片位置），用于定位改写/插入锚点。"""
import sys, zipfile
from lxml import etree
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'

def para_text(p):
    parts = []
    for t in p.iter():
        if t.tag == W + 't':
            parts.append(t.text or '')
        elif t.tag == W + 'tab':
            parts.append('\t')
    return ''.join(parts)

def main(path):
    z = zipfile.ZipFile(path)
    root = etree.fromstring(z.read('word/document.xml'))
    body = root.find(W + 'body')
    for i, el in enumerate(body):
        if el.tag == W + 'p':
            style = ''
            pPr = el.find(W + 'pPr')
            if pPr is not None:
                ps = pPr.find(W + 'pStyle')
                if ps is not None:
                    style = ps.get(W + 'val')
            nimg = len(el.findall('.//' + A + 'blip'))
            txt = para_text(el)
            print(f"[{i:03d}] P style={style!r} img={nimg} :: {txt[:160]}")
        elif el.tag == W + 'tbl':
            rows = el.findall(W + 'tr')
            first = ' | '.join(para_text(tc).strip() for tc in rows[0].findall(W + 'tc')) if rows else ''
            print(f"[{i:03d}] TABLE rows={len(rows)} :: {first[:150]}")
        elif el.tag == W + 'sectPr':
            print(f"[{i:03d}] SECTPR")

if __name__ == '__main__':
    main(sys.argv[1])
