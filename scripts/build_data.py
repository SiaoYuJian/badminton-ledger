#!/usr/bin/env python3
"""把社費帳本 xlsx 的「場地費明細」轉成每人一個 JSON,給靜態網頁讀取。

用法:
    python3 scripts/build_data.py ref/社費支出.xlsx
    python3 scripts/build_data.py ref/社費支出.xlsx --sheet "「2026場地費明細」" --out web/data

只用標準庫(zipfile + xml.etree),不需要 openpyxl / pandas。

餘額 = 該人員在明細表所有列(盤點 + 收入 + 支出)金額加總。
(不讀分攤表樞紐的「總計」:樞紐表常忘了重新整理,數字會過期。)
"""
import argparse
import datetime
import hashlib
import json
import re
import secrets
import sys
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from pathlib import Path

NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
DEFAULT_SHEET = "「2026場地費明細」"
TYPE_LABEL = {"盤點": "opening", "收入": "topup", "支出": "fee"}
EXCEL_EPOCH = datetime.date(1899, 12, 30)


def normalize_name(name: str) -> str:
    """前後空白、全半形統一。網頁端 JS 的 normalize 必須與此一致 (NFKC + trim)。"""
    return unicodedata.normalize("NFKC", name).strip()


def name_hash(salt: str, name: str) -> str:
    return hashlib.sha256((salt + normalize_name(name)).encode("utf-8")).hexdigest()


def read_sheet(xlsx: Path, sheet_name: str):
    """回傳該工作表每一列的 {欄位字母: 值(字串)}。"""
    with zipfile.ZipFile(xlsx) as z:
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            for si in ET.fromstring(z.read("xl/sharedStrings.xml")).findall("m:si", NS):
                shared.append("".join(t.text or "" for t in si.iter("{%s}t" % NS["m"])))

        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rels = {
            r.get("Id"): r.get("Target")
            for r in ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        }
        target = None
        names = []
        for s in wb.find("m:sheets", NS):
            names.append(s.get("name"))
            if s.get("name") == sheet_name:
                target = rels[s.get("{%s}id" % NS["r"])]
        if target is None:
            sys.exit(f"找不到工作表 {sheet_name!r}。現有工作表:{names}")
        path = target.lstrip("/")
        if not path.startswith("xl/"):
            path = "xl/" + path

        root = ET.fromstring(z.read(path))
        for row in root.find("m:sheetData", NS):
            cells = {}
            for c in row:
                v = c.find("m:v", NS)
                if v is None or v.text is None:
                    continue
                val = v.text
                if c.get("t") == "s":
                    val = shared[int(val)]
                cells[re.match(r"[A-Z]+", c.get("r")).group()] = val
            yield int(row.get("r")), cells


def to_float(s):
    try:
        return float(s)
    except (TypeError, ValueError):
        return None


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("xlsx", type=Path)
    ap.add_argument("--sheet", default=DEFAULT_SHEET)
    ap.add_argument("--out", type=Path, default=Path(__file__).resolve().parent.parent / "web" / "data")
    args = ap.parse_args()

    # 欄位:A日期 B人員 C金額 D類型 E地點 F權重 G備註
    records = defaultdict(list)
    skipped = []
    for rownum, c in read_sheet(args.xlsx, args.sheet):
        if rownum == 1:
            continue  # 標題列
        person, amount = c.get("B"), to_float(c.get("C"))
        if not person:
            continue  # 場次總表頭(無人員)或空白列
        kind = c.get("D")
        date = to_float(c.get("A"))
        if kind not in TYPE_LABEL or amount is None or date is None:
            skipped.append((rownum, person, c.get("A"), c.get("C"), kind))
            continue
        weight = to_float(c.get("F"))
        records[normalize_name(person)].append({
            "date": (EXCEL_EPOCH + datetime.timedelta(days=int(date))).isoformat(),
            "type": TYPE_LABEL[kind],
            "amount": round(amount, 2),
            "location": c.get("E") or "",
            "weight": weight if kind == "支出" and weight is not None else None,
            "note": c.get("G") or "",
            "_row": rownum,
        })

    if not records:
        sys.exit("沒有讀到任何人員資料,請確認工作表與欄位。")

    all_dates = [r["date"] for rs in records.values() for r in rs if r["type"] != "opening"]
    as_of = max(all_dates) if all_dates else ""

    # salt 固定存檔,避免每次重建網址都變
    salt_file = Path(__file__).resolve().parent / "salt.txt"
    if salt_file.exists():
        salt = salt_file.read_text().strip()
    else:
        salt = secrets.token_hex(8)
        salt_file.write_text(salt + "\n")

    args.out.mkdir(parents=True, exist_ok=True)
    for old in args.out.glob("*.json"):  # 清掉舊產物,避免已移除的人還留著
        old.unlink()

    balances = {}
    for name, rs in records.items():
        rs.sort(key=lambda r: (r["date"], r["_row"]))
        for r in rs:
            del r["_row"]
        balance = sum(r["amount"] for r in rs)
        balances[name] = balance
        summary = {
            "opening": round(sum(r["amount"] for r in rs if r["type"] == "opening"), 2),
            "topups": round(sum(r["amount"] for r in rs if r["type"] == "topup"), 2),
            "fees": round(sum(r["amount"] for r in rs if r["type"] == "fee"), 2),
            "sessions": len({r["date"] for r in rs if r["type"] == "fee"}),
        }
        payload = {"name": name, "balance": round(balance, 2), "asOf": as_of,
                   "summary": summary, "records": rs}
        (args.out / f"{name_hash(salt, name)}.json").write_text(
            json.dumps(payload, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")

    # meta 不含任何姓名,只給前端 salt 與更新日期
    (args.out / "meta.json").write_text(
        json.dumps({"salt": salt, "asOf": as_of}, ensure_ascii=False), encoding="utf-8")

    print(f"完成:{len(records)} 位社員,資料至 {as_of},輸出到 {args.out}")
    if skipped:
        print(f"注意:略過 {len(skipped)} 列(有人員但日期/金額/類型不完整):")
        for s in skipped:
            print("  列", s)
    # 供人工核對:餘額由小到大
    print("\n餘額最低 5 位:", sorted(balances.items(), key=lambda kv: kv[1])[:5])


if __name__ == "__main__":
    main()
