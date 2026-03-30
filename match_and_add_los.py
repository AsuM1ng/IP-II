from __future__ import annotations

from collections import Counter, defaultdict
from copy import deepcopy
from datetime import datetime, timedelta
import csv
import re
import xml.etree.ElementTree as ET
import zipfile

RAW_PATH = "原始数据.xlsx"
CLEAN_PATH = "数据_清洗后_切口感染.xlsx"
OUTPUT_XLSX_PATH = "原始数据_匹配切口感染_含住院时长.xlsx"
OUTPUT_CSV_PATH = "原始数据_匹配切口感染_含住院时长.csv"
UNMATCHED_CLEAN_XLSX = "未匹配_清洗数据_待人工核对.xlsx"
UNMATCHED_CLEAN_CSV = "未匹配_清洗数据_待人工核对.csv"

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_REL_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"

ROMAN_TO_INT = {
    "Ⅰ": "1", "Ⅱ": "2", "Ⅲ": "3", "Ⅳ": "4", "Ⅴ": "5",
    "I": "1", "II": "2", "III": "3", "IV": "4", "V": "5",
}


def col_to_idx(col: str) -> int:
    n = 0
    for ch in col:
        if ch.isalpha():
            n = n * 26 + ord(ch.upper()) - 64
    return n - 1


def idx_to_col(idx: int) -> str:
    idx += 1
    out = ""
    while idx:
        idx, rem = divmod(idx - 1, 26)
        out = chr(65 + rem) + out
    return out


def normalize_header(text: str) -> str:
    text = (text or "").replace("\n", "")
    text = text.replace("（", "(").replace("）", ")")
    text = text.replace("：", ":").replace("，", ",").replace("、", ",")
    return re.sub(r"\s+", "", text)


def normalize_value(v):
    if v is None:
        return ""
    if isinstance(v, str):
        t = v.strip()
        t = t.replace("（", "(").replace("）", ")").replace("，", ",").replace("：", ":")
        if t.lower() in {"", "nan", "none", "null", "na", "n/a", "不详", "未知"}:
            return ""
        t = ROMAN_TO_INT.get(t, t)
        if re.fullmatch(r"[-+]?\d+(?:\.0+)?", t):
            return str(int(float(t)))
        return t
    if isinstance(v, float):
        if v.is_integer():
            return str(int(v))
        return ("%.10f" % v).rstrip("0").rstrip(".")
    if isinstance(v, int):
        return str(v)
    return str(v)


def parse_excel_datetime(v):
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v
    if isinstance(v, (int, float)):
        base = datetime(1899, 12, 30)
        return base + timedelta(days=float(v))
    if isinstance(v, str):
        t = v.strip()
        for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S", "%Y.%m.%d"):
            try:
                return datetime.strptime(t, fmt)
            except ValueError:
                continue
    return None


def get_sheet_path(zf: zipfile.ZipFile) -> str:
    wb = ET.fromstring(zf.read("xl/workbook.xml"))
    sheets = wb.find(f"{{{NS_MAIN}}}sheets")
    first = sheets.findall(f"{{{NS_MAIN}}}sheet")[0]
    rid = first.attrib[f"{{{NS_REL_DOC}}}id"]

    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    for rel in rels.findall(f"{{{NS_REL_PKG}}}Relationship"):
        if rel.attrib.get("Id") == rid:
            target = rel.attrib["Target"]
            if target.startswith("/"):
                return target.lstrip("/")
            if target.startswith("xl/"):
                return target
            return "xl/" + target
    raise RuntimeError("Cannot locate first sheet path")


def read_sst(zf: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    return ["".join(t.text or "" for t in si.iter(f"{{{NS_MAIN}}}t")) for si in root.findall(f"{{{NS_MAIN}}}si")]


def read_xlsx(path: str):
    with zipfile.ZipFile(path) as zf:
        sst = read_sst(zf)
        sheet_path = get_sheet_path(zf)
        root = ET.fromstring(zf.read(sheet_path))

    sheet_data = root.find(f"{{{NS_MAIN}}}sheetData")
    rows = []
    for row in sheet_data.findall(f"{{{NS_MAIN}}}row"):
        values = {}
        for c in row.findall(f"{{{NS_MAIN}}}c"):
            ref = c.attrib.get("r", "")
            idx = col_to_idx("".join(ch for ch in ref if ch.isalpha()))
            t = c.attrib.get("t")
            v_elem = c.find(f"{{{NS_MAIN}}}v")
            is_elem = c.find(f"{{{NS_MAIN}}}is")
            value = None
            if is_elem is not None:
                value = "".join(tn.text or "" for tn in is_elem.iter(f"{{{NS_MAIN}}}t"))
            elif v_elem is None:
                value = None
            else:
                raw = v_elem.text
                if t == "s":
                    value = sst[int(raw)] if raw is not None else ""
                elif t == "b":
                    value = 1 if raw == "1" else 0
                else:
                    try:
                        num = float(raw) if raw is not None else None
                        value = int(num) if (num is not None and num.is_integer()) else num
                    except (ValueError, TypeError):
                        value = raw
            values[idx] = value
        rows.append(values)

    header_row = rows[0]
    max_col = max(header_row.keys())
    headers = [header_row.get(i, "") for i in range(max_col + 1)]
    data = [[r.get(i) for i in range(max_col + 1)] for r in rows[1:]]
    return headers, data


def build_header_mapping(raw_headers, clean_headers):
    raw_norm = {normalize_header(h): i for i, h in enumerate(raw_headers)}
    explicit = {
        "性别": "性别(1=男,2=女)",
        "术前前白蛋白PALB": "术前前白蛋白PALB(0=未测)",
        "术前白蛋白ALB": "术前白蛋白ALB(0=未测)",
        "术前血红蛋白HGB": "术前血红蛋白HGB(0=未测)",
        "术前口咽拭子": "术前口咽拭子(未生长=0,阳性=1,未测=2)",
        "病变部位": "病变部位(喉=1,鼻=2,扁桃体=3,腮腺=4,咽部=5,唇=6,甲状腺=7,食管=8,舌=9,耳=10,颌,面部=11,口底=12,气管=13,口腔,牙龈=14,梨状窝=15,上颌窦=16,皮肤肿物=17,腭=18,胸部=19,颈部=20)",
        "术前抗菌药物": "术前抗菌药物:未用=0,头孢曲松钠=1,头孢呋辛钠=2,左奥硝唑氯化钠=3,吗啉硝唑氯化钠=4,甲磺酸左氧氟沙星氯化钠=5,克林霉素磷酸酯=6,盐酸莫西沙星=7,头孢噻肟钠=8,甲硝唑氯化钠=9,头孢米诺钠=10,奥硝唑=11,美洛西林钠舒巴坦钠=12,注射用哌拉西林钠舒巴坦钠=13,头孢哌酮钠舒巴坦钠=14,阿奇霉素=15,美罗培南=16,万古霉素=17",
        "手术时长(min)": "手术时长(min)(不详=0,其他具体写出)",
        "吻合方式": "吻合方式(如有吻合口:手工=1,器械=2,不详=0)",
        "术前同步放化疗": "术前同步放化疗(0=无,1:<=60Gy,2:>60Gy,有,但具体剂量不详=3)",
        "皮瓣": "皮瓣(无=0,颏下皮瓣=1,股前外侧皮瓣=2,鼻唇沟=3,颈阔肌=4,锁骨=5,其他=6,前臂桡侧皮瓣=7,游离空肠瓣=8,游离腓骨瓣=9,胸大肌皮瓣=10,局部转移皮瓣=11)待整理",
        "术后病理": "术后病理无=0,warthin瘤=1,鳞癌=2,乳头状癌=3,多形性腺瘤=4,基底细胞癌=5,腺癌=6,黑色素瘤=7,未见癌=8,肉瘤=9,良性=10,甲状腺髓样癌=11,粘液表皮样癌=12,分化差的癌=13,腺样囊性癌=14,淋巴细胞瘤=15,梭形细胞瘤=16,囊肿=17",
        "最新版pTNM": "最新版pTNM(5=5期,1=I期,2=II期,3=III期,4=IV期,5=无)",
        "术后0-3天白蛋白ALB": "术后0-3天白蛋白ALB(选最低的)(0=未测)",
        "术后0-3天前白蛋白PALB": "术后0-3天前白蛋白PALB(0=未测)",
        "非计划二次手术": "非计划二次手术(0=否,1=是)",
        "肺部感染": "肺部感染(0=无,1=有)",
        "吻合口瘘": "吻合口瘘(0=无,1=有)",
        "吻合口瘘确认距术后天数": "吻合口瘘确认距术后天数(0=未发生,具体已写出)",
        "脂肪液化": "脂肪液化(0=无,1=有)",
        "切口感染": "切口感染(0=无,1=有)",
        "是否多重耐药": "是否多重耐药(0=否,1=是)",
    }

    mapping = {}
    for ci, ch in enumerate(clean_headers):
        key = normalize_header(ch)
        if key in raw_norm:
            mapping[ci] = raw_norm[key]
            continue
        if key in explicit and explicit[key] in raw_norm:
            mapping[ci] = raw_norm[explicit[key]]
            continue
        for rk, ri in raw_norm.items():
            if rk.startswith(key):
                mapping[ci] = ri
                break
        if ci not in mapping:
            raise KeyError(f"无法为清洗列找到原始列映射: {ch}")
    return mapping


def make_key(row, indexes):
    return tuple(normalize_value(row[i] if i < len(row) else None) for i in indexes)


def match_rows(raw_rows, clean_rows, mapping):
    clean_len = len(clean_rows[0]) if clean_rows else 0
    map_order = [mapping[i] for i in range(clean_len)]

    clean_counter = Counter(make_key(r, range(clean_len)) for r in clean_rows)
    raw_exact_key = [make_key(r, map_order) for r in raw_rows]

    matched_raw = set()
    matched_clean = set()

    key_to_clean_idxs = defaultdict(list)
    for ci, crow in enumerate(clean_rows):
        key_to_clean_idxs[make_key(crow, range(clean_len))].append(ci)

    for ri, rkey in enumerate(raw_exact_key):
        if clean_counter[rkey] > 0:
            clean_counter[rkey] -= 1
            matched_raw.add(ri)
            matched_clean.add(key_to_clean_idxs[rkey].pop())

    core_cols = [0, 1, 37, 41, 20, 27, 29]
    raw_core_map = defaultdict(list)
    for ri, row in enumerate(raw_rows):
        if ri in matched_raw:
            continue
        core = tuple(normalize_value(row[map_order[c]]) for c in core_cols)
        raw_core_map[core].append(ri)

    def score_pair(crow, rrow):
        eq = 0
        ne = 0
        for ci in range(clean_len):
            cv = normalize_value(crow[ci])
            rv = normalize_value(rrow[map_order[ci]])
            if cv == "":
                continue
            ne += 1
            if cv == rv:
                eq += 1
        return eq, ne

    for ci, crow in enumerate(clean_rows):
        if ci in matched_clean:
            continue
        core = tuple(normalize_value(crow[c]) for c in core_cols)
        candidates = raw_core_map.get(core, [])
        best = None
        best_score = (-1, 1)
        for ri in candidates:
            if ri in matched_raw:
                continue
            eq, ne = score_pair(crow, raw_rows[ri])
            if eq > best_score[0] or (eq == best_score[0] and ne > best_score[1]):
                best = ri
                best_score = (eq, ne)
        if best is None:
            continue
        eq, ne = best_score
        if ne >= 16 and eq / ne >= 0.70:
            matched_raw.add(best)
            matched_clean.add(ci)

    return sorted(matched_raw), sorted(matched_clean)


def write_csv(path: str, headers: list[str], rows: list[list]):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        writer.writerows(rows)


def write_simple_xlsx(path: str, headers: list[str], rows: list[list]):
    content_types = """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">
  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>
  <Default Extension=\"xml\" ContentType=\"application/xml\"/>
  <Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>
  <Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>
</Types>"""
    rels = """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>
</Relationships>"""
    wb = """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">
  <sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets>
</workbook>"""
    wb_rels = """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>
</Relationships>"""

    def cell_xml(r, c, v):
        ref = f"{idx_to_col(c)}{r}"
        if v is None or v == "":
            return f'<c r="{ref}"/>'
        if isinstance(v, (int, float)):
            return f'<c r="{ref}"><v>{v}</v></c>'
        text = str(v).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        return f'<c r="{ref}" t="inlineStr"><is><t>{text}</t></is></c>'

    all_rows = [headers] + rows
    row_xml = []
    for r_idx, row in enumerate(all_rows, start=1):
        cells = "".join(cell_xml(r_idx, c_idx, row[c_idx] if c_idx < len(row) else None) for c_idx in range(len(headers)))
        row_xml.append(f'<row r="{r_idx}">{cells}</row>')

    dim = f"A1:{idx_to_col(len(headers)-1)}{len(all_rows)}"
    sheet = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        f"<worksheet xmlns=\"{NS_MAIN}\"><dimension ref=\"{dim}\"/>"
        "<sheetViews><sheetView workbookViewId=\"0\"/></sheetViews>"
        f"<sheetData>{''.join(row_xml)}</sheetData></worksheet>"
    )

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", rels)
        z.writestr("xl/workbook.xml", wb)
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        z.writestr("xl/worksheets/sheet1.xml", sheet)


def update_dimension(root: ET.Element, end_col_idx: int, end_row_idx: int):
    dim = root.find(f"{{{NS_MAIN}}}dimension")
    if dim is None:
        dim = ET.SubElement(root, f"{{{NS_MAIN}}}dimension")
    dim.attrib["ref"] = f"A1:{idx_to_col(end_col_idx)}{end_row_idx}"


def rewrite_raw_xlsx_preserve_dates(raw_path: str, output_path: str, matched_raw_idx: list[int], los_values: dict[int, int | str]):
    with zipfile.ZipFile(raw_path, "r") as zin:
        file_map = {name: zin.read(name) for name in zin.namelist()}
        sheet_path = get_sheet_path(zin)

    root = ET.fromstring(file_map[sheet_path])
    sheet_data = root.find(f"{{{NS_MAIN}}}sheetData")
    rows = sheet_data.findall(f"{{{NS_MAIN}}}row")
    if not rows:
        raise RuntimeError("原始表为空")

    header_row = deepcopy(rows[0])
    data_row_map = {int(r.attrib.get("r", "0")) - 2: r for r in rows[1:]}

    # 计算原始最后一列
    header_cells = header_row.findall(f"{{{NS_MAIN}}}c")
    max_col_idx = max(col_to_idx("".join(ch for ch in c.attrib.get("r", "") if ch.isalpha())) for c in header_cells)
    los_col_idx = max_col_idx + 1

    # 头部追加“住院时长”
    header_cell = ET.Element(f"{{{NS_MAIN}}}c", {"r": f"{idx_to_col(los_col_idx)}1", "t": "inlineStr"})
    is_elem = ET.SubElement(header_cell, f"{{{NS_MAIN}}}is")
    ET.SubElement(is_elem, f"{{{NS_MAIN}}}t").text = "住院时长"
    header_row.append(header_cell)

    new_rows = [header_row]
    for out_r, raw_idx in enumerate(matched_raw_idx, start=2):
        if raw_idx not in data_row_map:
            continue
        row = deepcopy(data_row_map[raw_idx])
        row.attrib["r"] = str(out_r)

        for c in row.findall(f"{{{NS_MAIN}}}c"):
            old_ref = c.attrib.get("r", "")
            col = "".join(ch for ch in old_ref if ch.isalpha())
            c.attrib["r"] = f"{col}{out_r}"

        los = los_values.get(raw_idx, "")
        los_cell = ET.Element(f"{{{NS_MAIN}}}c", {"r": f"{idx_to_col(los_col_idx)}{out_r}"})
        if los != "":
            ET.SubElement(los_cell, f"{{{NS_MAIN}}}v").text = str(los)
        row.append(los_cell)
        new_rows.append(row)

    for old in list(sheet_data):
        sheet_data.remove(old)
    for nr in new_rows:
        sheet_data.append(nr)

    update_dimension(root, los_col_idx, len(new_rows))
    file_map[sheet_path] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, content in file_map.items():
            zout.writestr(name, content)


def main():
    raw_headers, raw_rows = read_xlsx(RAW_PATH)
    clean_headers, clean_rows = read_xlsx(CLEAN_PATH)
    mapping = build_header_mapping(raw_headers, clean_headers)

    matched_raw_idx, matched_clean_idx = match_rows(raw_rows, clean_rows, mapping)

    admission_idx = raw_headers.index("入院日期")
    discharge_idx = raw_headers.index("出院日期")
    out_headers = deepcopy(raw_headers) + ["住院时长"]

    out_rows = []
    los_values: dict[int, int | str] = {}
    for ri in matched_raw_idx:
        row = raw_rows[ri]
        admit = parse_excel_datetime(row[admission_idx] if admission_idx < len(row) else None)
        dis = parse_excel_datetime(row[discharge_idx] if discharge_idx < len(row) else None)
        los = ""
        if admit and dis:
            d = (dis.date() - admit.date()).days
            if d >= 0:
                los = d
        los_values[ri] = los
        out_rows.append(deepcopy(row) + [los])

    write_csv(OUTPUT_CSV_PATH, out_headers, out_rows)
    # 该输出通过保留原始单元格样式来保留日期格式
    rewrite_raw_xlsx_preserve_dates(RAW_PATH, OUTPUT_XLSX_PATH, matched_raw_idx, los_values)

    unmatched_clean_set = set(range(len(clean_rows))) - set(matched_clean_idx)
    unmatched_clean_rows = [clean_rows[i] for i in sorted(unmatched_clean_set)]
    write_csv(UNMATCHED_CLEAN_CSV, clean_headers, unmatched_clean_rows)
    write_simple_xlsx(UNMATCHED_CLEAN_XLSX, clean_headers, unmatched_clean_rows)

    print(f"raw rows: {len(raw_rows)}")
    print(f"clean rows: {len(clean_rows)}")
    print(f"matched raw rows written: {len(matched_raw_idx)}")
    print(f"matched clean rows: {len(matched_clean_idx)}")
    print(f"unmatched clean rows: {len(unmatched_clean_rows)}")
    print(f"output xlsx (preserved date format): {OUTPUT_XLSX_PATH}")
    print(f"output csv: {OUTPUT_CSV_PATH}")
    print(f"unmatched clean xlsx: {UNMATCHED_CLEAN_XLSX}")
    print(f"unmatched clean csv: {UNMATCHED_CLEAN_CSV}")


if __name__ == "__main__":
    main()
