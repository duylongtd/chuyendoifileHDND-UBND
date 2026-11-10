"""Kiểm chứng độc lập (openpyxl + expat): 2 file kết quả so với file gốc, từng ô.

Dùng thư viện KHÁC với bộ tách (openpyxl) để chắc chắn kết quả đúng:
  - zip toàn vẹn, thành phần không tách giữ nguyên từng byte, đã bật tính lại công thức
  - XML mọi sheet hợp lệ, đủ số dòng / số công thức, công thức dùng chung không mất ô gốc
  - từng dòng giữ lại khớp từng ô với dòng gốc (công thức dịch đúng theo dòng), đúng thứ tự
  - không mất dòng, không thừa dòng

Cách chạy (cần: pip install openpyxl):
  python scripts/verify_split.py <file gốc> <file HĐND> <file UBND> <các mã VTVL HĐND, cách nhau dấu phẩy>
Ví dụ:
  python scripts/verify_split.py goc.xlsx hdnd.xlsx ubnd.xlsx H27.120_VT2,H27.120_VT4,H27.120_VT5,H27.120_VT6
"""
import sys, re, zipfile, zlib, collections, unicodedata, xml.parsers.expat
import openpyxl
orig_p, hd_p, ub_p = sys.argv[1:4]
HD = set(sys.argv[4].split(","))
SPECS = [("1. A1_NhiemVu","A","NV"),("2. A2_CongViec","C","VT"),("3. B1_KhoiLuong","A","CV"),("4. B2_CT_TacDong","A","CV"),("5. B2_ThayDoi","A","CV"),
         ("6. C1_VTVL","A","VT"),("7. C2_NangLuc","A","VT"),("8. C3_NhuCau","A","VT"),("9. D1_NhanLuc","B","VT"),("10. D2_PhuongAn","A","VT"),("DU_BAO_5_NAM","A","VT")]
ok = True
def check(cond, msg):
    global ok
    print(("  OK   " if cond else "  SAI  ") + msg)
    ok = ok and cond

# 1) Cấu trúc zip: thành phần không tách phải giống hệt (CRC), calcChain bị bỏ
zo, zh, zu = (zipfile.ZipFile(p) for p in (orig_p, hd_p, ub_p))
for z, lab in ((zh, "HĐND"), (zu, "UBND")):
    check(z.testzip() is None, f"{lab}: zip toàn vẹn (CRC mọi thành phần)")
    names_o = {i.filename: i for i in zo.infolist()}; names_n = {i.filename: i for i in z.infolist()}
    # file gốc có thể không có calcChain (tùy phần mềm lưu) -> chỉ cho phép thiếu đúng thành phần đó
    check(set(names_o) - set(names_n) <= {"xl/calcChain.xml"} and not (set(names_n) - set(names_o)), f"{lab}: đủ thành phần (chỉ được bỏ calcChain nếu có)")
    changed = sorted(n for n in names_n if names_n[n].CRC != names_o[n].CRC)
    print(f"       {lab}: {len(names_n) - len(changed)} thành phần giữ nguyên, {len(changed)} thành phần viết lại: {[c.split('/')[-1] for c in changed]}")
    wbx = z.read("xl/workbook.xml").decode(); check('fullCalcOnLoad="1"' in wbx, f"{lab}: bật tính lại công thức khi mở")
    check("calcChain" not in z.read("[Content_Types].xml").decode() and "calcChain" not in z.read("xl/_rels/workbook.xml.rels").decode(), f"{lab}: gỡ tham chiếu calcChain")

# 2) XML hợp lệ + công thức dùng chung đủ ô gốc (quét toàn bộ, kể cả 30.000 dòng mẫu)
def scan_sheet(z, path):
    st = {"rows": 0, "f": 0, "cells": 0, "masters": {}, "orphans": 0, "cur": None, "prev": 0, "order": True}
    def start(name, a):
        if name == "row":
            st["rows"] += 1; r = int(a["r"]); st["order"] &= r > st["prev"]; st["prev"] = r
        elif name == "c": st["cells"] += 1; st["cur"] = a["r"]
        elif name == "f":
            st["f"] += 1
            if a.get("t") == "shared":
                if "ref" in a: st["masters"][a["si"]] = a["ref"]
                elif a["si"] not in st["masters"]: st["orphans"] += 1
                else:
                    m = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", st["masters"][a["si"]] if ":" in st["masters"][a["si"]] else st["masters"][a["si"]] + ":" + st["masters"][a["si"]])
                    c = re.match(r"([A-Z]+)(\d+)", st["cur"])
                    if not (int(m.group(2)) <= int(c.group(2)) <= int(m.group(4))): st["orphans"] += 1
    p = xml.parsers.expat.ParserCreate(); p.StartElementHandler = start; p.buffer_text = True
    with z.open(path) as f:
        while True:
            chunk = f.read(1 << 22)
            if not chunk: break
            p.Parse(chunk, False)
    p.Parse(b"", True)
    return st
def sheet_paths(z):
    wb = z.read("xl/workbook.xml").decode(); rels = z.read("xl/_rels/workbook.xml.rels").decode()
    t = {re.search(r'Id="([^"]+)"', x).group(1): re.search(r'Target="([^"]+)"', x).group(1) for x in re.findall(r"<Relationship\b[^>]*>", rels)}
    return {n: "xl/" + t[r].lstrip("/").replace("xl/", "") for n, r in re.findall(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wb)}
po = sheet_paths(zo)
for name, _, _ in SPECS:
    so = scan_sheet(zo, po[name])
    for z, lab in ((zh, "HĐND"), (zu, "UBND")):
        s = scan_sheet(z, po[name])
        check(s["orphans"] == 0 and s["order"] and s["rows"] == so["rows"] and s["f"] == so["f"],
              f"{lab} {name}: XML hợp lệ, {s['rows']} dòng, {s['f']} công thức (gốc {so['f']}), công thức dùng chung thiếu ô gốc: {s['orphans']}")

# 3) Nội dung từng ô: dòng giữ lại = dòng gốc của khối đó, đúng thứ tự, công thức dịch đúng theo dòng
def norm(v, row):
    if hasattr(v, "text"):  # ArrayFormula
        v = "{" + str(v.text) + "}"
    if isinstance(v, str) and v.startswith(("=", "{=")):
        # chỉ chuẩn hóa tham chiếu NGOÀI chuỗi văn bản "..." và tên sheet '...'
        parts = re.split(r'("(?:[^"]|"")*"|\'(?:[^\']|\'\')*\')', v)
        for k in range(0, len(parts), 2):
            parts[k] = re.sub(r"(?<![A-Za-z0-9_$.])(\$?[A-Z]{1,3})(\d+)(?![\w(!])", lambda m: m.group(1) + ("{r}" if int(m.group(2)) == row else m.group(2)), parts[k])
        return "".join(parts)
    return v
def load(path, maxrow):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=False)
    out = {}
    for name, _, _ in SPECS:
        rows = {}
        for i, r in enumerate(wb[name].iter_rows(min_row=1, max_row=maxrow, values_only=True), 1):
            rows[i] = r
        out[name] = rows
    return out
def cached(path, maxrow):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    return {name: {i: r for i, r in enumerate(wb[name].iter_rows(min_row=1, max_row=maxrow, values_only=True), 1)} for name, _, _ in SPECS}
MAXROW = 760
O, H, U = load(orig_p, MAXROW), load(hd_p, MAXROW), load(ub_p, MAXROW)
OC, HC, UC = cached(orig_p, MAXROW), cached(hd_p, MAXROW), cached(ub_p, MAXROW)
def isconst(v): return v is not None and v != "" and not (hasattr(v, "text") or (isinstance(v, str) and v.startswith("=")))
cv2org, nv2org = {}, {}
for i, r in O["2. A2_CongViec"].items():
    if i >= 3 and r[2]:
        o = "H" if r[2] in HD else "U"; cv2org[r[0]] = o; nv2org.setdefault(r[1], set()).add(o)
def org_of(name, kind, keycol, i):
    k = OC[name][i][ord(keycol) - 65]
    if k in (None, ""): return None
    if kind == "VT": return "H" if k in HD else "U"
    if kind == "CV": return cv2org.get(k, "B")
    s = nv2org.get(k, set())
    if s: return "H" if s == {"H"} else "U" if s == {"U"} else "B"
    # nhiệm vụ chưa gắn công việc nào ở A2: xếp theo cột "Cơ quan" (C); chỉ nhắc HĐND -> HĐND, còn lại -> UBND
    cq = str(OC[name][i][2] or "").lower()
    h = "hđnd" in cq or "hội đồng nhân dân" in cq; u = "ubnd" in cq or "ủy ban nhân dân" in cq or "uỷ ban nhân dân" in cq
    return "H" if h and not u else "U"
total_cells = 0
for name, keycol, kind in SPECS:
    rows_o = O[name]; data = [i for i in rows_o if i >= 3 and any(isconst(v) for v in rows_o[i]) or (i >= 3 and OC[name][i][ord(keycol) - 65] not in (None, ""))]
    if not data: 
        N = lambda rows, i: [norm(v, i) for v in rows[i]]
        check(all(N(H[name], i) == N(rows_o, i) and N(U[name], i) == N(rows_o, i) for i in rows_o) and zh.getinfo(po[name]).CRC == zo.getinfo(po[name]).CRC == zu.getinfo(po[name]).CRC, f"{name}: không có dữ liệu → giữ nguyên từng byte ở cả hai file"); continue
    last = max(data)
    for new, newc, lab, tag in ((H, HC, "HĐND", "H"), (U, UC, "UBND", "U")):
        exp = [i for i in data if org_of(name, kind, keycol, i) in (tag, "B")]
        bad = 0; ncell = 0
        check(new[name][1] == rows_o[1] and new[name][2] == rows_o[2], f"{lab} {name}: 2 dòng tiêu đề giữ nguyên")
        for d, s in enumerate(exp, 3):
            a = [norm(v, s) for v in rows_o[s]]; b = [norm(v, d) for v in new[name][d]]
            ncell += sum(1 for v in a if v not in (None, ""))
            if a != b or OC[name][s] != newc[name][d]:
                bad += 1
                if bad <= 2: print("        lệch dòng gốc", s, "→", d, [(x, y) for x, y in zip(a, b) if x != y][:2])
        # các vị trí còn lại trong vùng dữ liệu: không còn ô nhập tay, không còn giá trị cũ; công thức mẫu vẫn còn
        leftover = 0; cached_left = 0; fml = 0
        for d in range(3 + len(exp), last + 1):
            leftover += sum(1 for v in new[name][d] if isconst(v)); cached_left += sum(1 for v in newc[name][d] if v not in (None, ""))
            fml += sum(1 for v in new[name][d] if v is not None and not isconst(v))
        # sau vùng dữ liệu: dòng mẫu giữ nguyên
        tail_same = all([norm(v, i) for v in new[name][i]] == [norm(v, i) for v in rows_o[i]] for i in range(last + 1, MAXROW + 1))
        total_cells += ncell
        check(bad == 0 and leftover == 0 and cached_left == 0 and tail_same,
              f"{lab} {name}: {len(exp)}/{len(data)} dòng giữ đúng thứ tự, {ncell} ô khớp; dòng thừa còn ô nhập tay={leftover}, còn giá trị cũ={cached_left}, công thức mẫu còn={fml}; dòng mẫu phía sau giữ nguyên={tail_same}")
    # không mất dòng: mỗi dòng gốc phải có mặt ở ít nhất một file
    nh = sum(1 for i in data if org_of(name, kind, keycol, i) in ("H", "B")); nu = sum(1 for i in data if org_of(name, kind, keycol, i) in ("U", "B"))
    nb = sum(1 for i in data if org_of(name, kind, keycol, i) == "B")
    check(nh + nu - nb == len(data), f"{name}: HĐND {nh} + UBND {nu} − dùng chung {nb} = {len(data)} dòng gốc (không mất, không thừa)")
print(f"\nTổng số ô đã so khớp: {total_cells}")
print("KẾT QUẢ:", "ĐẠT TẤT CẢ" if ok else "CÓ LỖI")
sys.exit(0 if ok else 1)
