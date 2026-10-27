#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""交付包对账：磁盘 ↔ zip 逐文件 md5、包内文件白名单、包内无 csv/txt、
包内名全 ASCII、交付目录无解包残留目录。

（取消 manifest.csv 后 preflight 的「清单对账」转 SKIP，这道闸由本脚本接管——
SKIP 不是「没问题」，是「没查」。）
"""

from __future__ import annotations

import hashlib
import json
import re
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PACK = ROOT / "07_交付包"
DELIV = ROOT / "07_交付"
ZIP = DELIV / "交付包_hd_tdcs_aud_cognition_20261027.zip"
REPRO_ZIP = DELIV / "reproducibility_package_hd_tdcs_aud_cognition_20261027.zip"

# 包内允许出现的扩展名（投稿件本体：docx / xlsx / py）
ALLOWED_EXT = {".docx", ".xlsx", ".py"}
# 包内禁止出现的扩展名
FORBIDDEN_EXT = {".csv", ".txt", ".md", ".doc", ".docm", ".xlsm", ".pdf", ".json", ".DS_Store"}
ASCII_RE = re.compile(r"^[\x20-\x7E]+$")

failures: list[str] = []
notes: list[str] = []


def fail(msg: str) -> None:
    failures.append(msg)
    print("FAIL:", msg)


def ok(msg: str) -> None:
    print("PASS:", msg)


def md5(p: Path) -> str:
    return hashlib.md5(p.read_bytes()).hexdigest()


def check_pack(pack_dir: Path, label: str) -> None:
    files = [p for p in sorted(pack_dir.rglob("*")) if p.is_file()]
    if not files:
        fail(f"{label} 目录为空：{pack_dir}")
        return
    bad_ext = [p for p in files if p.suffix.lower() not in ALLOWED_EXT]
    if bad_ext:
        fail(f"{label} 含不允许的扩展名：{[str(p.relative_to(pack_dir)) for p in bad_ext]}")
    else:
        ok(f"{label} 扩展名合规（{sorted(ALLOWED_EXT)}），共 {len(files)} 个文件")
    for p in files:
        rel = str(p.relative_to(pack_dir))
        if not ASCII_RE.match(rel):
            fail(f"{label} 含非 ASCII 路径：{rel}")
    for p in files:
        if p.suffix.lower() in FORBIDDEN_EXT:
            fail(f"{label} 含禁止文件：{rel_of(p, pack_dir)}")
    for p in files:
        if p.stat().st_size == 0:
            fail(f"{label} 含 0 字节文件：{rel_of(p, pack_dir)}")


def rel_of(p: Path, base: Path) -> str:
    return str(p.relative_to(base))


def check_zip(pack_dir: Path, zip_path: Path, label: str) -> None:
    if not zip_path.exists():
        fail(f"{label} zip 不存在：{zip_path}")
        return
    disk = {str(p.relative_to(pack_dir)): md5(p)
            for p in pack_dir.rglob("*") if p.is_file()}
    with zipfile.ZipFile(zip_path) as z:
        bad = z.testzip()
        if bad:
            fail(f"{label} zip 损坏：{bad}")
            return
        names = [n for n in z.namelist() if not n.endswith("/")]
        zipped = {n: hashlib.md5(z.read(n)).hexdigest() for n in names}
    missing = sorted(set(disk) - set(zipped))
    extra = sorted(set(zipped) - set(disk))
    diff = sorted(k for k in set(disk) & set(zipped) if disk[k] != zipped[k])
    if missing:
        fail(f"{label} 磁盘有、zip 无：{missing}")
    if extra:
        fail(f"{label} zip 有、磁盘无：{extra}")
    if diff:
        fail(f"{label} md5 不一致：{diff}")
    if not (missing or extra or diff):
        ok(f"{label} 磁盘 ↔ zip 逐文件 md5 一致（{len(disk)} 个文件）")
    non_ascii = [n for n in zipped if not ASCII_RE.match(n)]
    if non_ascii:
        fail(f"{label} zip 内含非 ASCII 名：{non_ascii}")
    dirty = [n for n in zipped if Path(n).suffix.lower() in FORBIDDEN_EXT]
    if dirty:
        fail(f"{label} zip 内含禁止文件：{dirty}")


def check_no_extraction_residue() -> None:
    for d in (DELIV, PACK):
        for child in d.iterdir():
            if child.is_dir():
                stem = child.name.lower()
                for zp in (ZIP, REPRO_ZIP):
                    if stem == zp.stem.lower():
                        fail(f"交付目录存在与 zip 同名的解包残留目录：{child}")
    ok("交付目录无解包残留目录")


def check_pack_excludes_internal() -> None:
    """包内不得混入内部过程件（原始数据、审稿截图、旧稿）。"""
    forbidden_names = {"Manuscript_Frontiers_投稿版_1016.docx", "Frontiers决策信与审稿意见_1022.docx"}
    for p in PACK.rglob("*"):
        if p.is_file() and p.name in forbidden_names:
            fail(f"交付包混入内部过程件：{p.name}")
    if not any(PACK.rglob("*.png")):
        ok("交付包未混入图片/截图类过程件")


def main() -> int:
    check_pack(PACK, "交付包")
    check_zip(PACK, ZIP, "交付包")
    check_no_extraction_residue()
    check_pack_excludes_internal()
    print()
    if failures:
        print(f"交付包对账未通过：{len(failures)} 项")
        for f in failures:
            print("  -", f)
        (ROOT / "05_验证记录" / "交付包对账.json").write_text(
            json.dumps({"passed": False, "failures": failures}, ensure_ascii=False, indent=2),
            encoding="utf-8")
        return 1
    print("交付包对账全部通过。")
    (ROOT / "05_验证记录" / "交付包对账.json").write_text(
        json.dumps({"passed": True, "failures": []}, ensure_ascii=False, indent=2), encoding="utf-8")
    return 0


if __name__ == "__main__":
    sys.exit(main())
