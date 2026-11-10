"""Bộ test đầu-cuối cho web tách file HĐND-UBND (chạy với server đang mở, ví dụ npm run dev).

Cách chạy:
  ACCESS_CODE=<mã truy cập> python scripts/e2e_test.py [--base http://localhost:3600] [--sample file_that.xlsx]
  (hoặc --cookie-jar <file cookie của curl> nếu đã đăng nhập bằng curl)

Cần: pip install openpyxl. Không in mã truy cập hay cookie ra màn hình.
Mỗi ca test in PASS/FAIL; mã thoát khác 0 nếu có ca FAIL.
"""
import argparse, hashlib, io, json, math, os, struct, sys, threading, time, urllib.error, urllib.request, zipfile, zlib
from concurrent.futures import ThreadPoolExecutor

ap = argparse.ArgumentParser()
ap.add_argument("--base", default="http://localhost:3600")
ap.add_argument("--cookie-jar")
ap.add_argument("--sample", help="file bộ biểu thật để test tải/tách (tùy chọn)")
ap.add_argument("--only", default="", help="chỉ chạy nhóm có tên chứa chuỗi này")
args = ap.parse_args()
U = args.base.rstrip("/")
COOKIE = ""
CLIENT_IP = ""  # mỗi nhóm test dùng một IP giả riêng để giới hạn theo IP của nhóm này không ảnh hưởng nhóm khác
RESULTS = []
PART = 4 * 1024 * 1024


def http(method, path, data=None, ctype="application/json", cookie=None, headers=None, timeout=300):
    h = dict(headers or {})
    if CLIENT_IP and "X-Forwarded-For" not in h:
        h["X-Forwarded-For"] = CLIENT_IP
    ck = COOKIE if cookie is None else cookie
    if ck:
        h["Cookie"] = ck
    if data is not None:
        h["Content-Type"] = ctype
    req = urllib.request.Request(U + path, data=data, method=method, headers=h)
    try:
        r = urllib.request.urlopen(req, timeout=timeout)
        return r.status, r.read(), dict(r.headers)
    except urllib.error.HTTPError as e:
        return e.code, e.read(), dict(e.headers)
    except (urllib.error.URLError, ConnectionError, TimeoutError) as e:
        return 0, str(e).encode(), {}


def J(b):
    try:
        return json.loads(b)
    except Exception:
        return {}


def msg(b):
    d = b if isinstance(b, dict) else J(b)
    return (d.get("error") or {}).get("message") or (d.get("job", {}).get("error") or {}).get("message") or ""


def check(group, name, ok, detail=""):
    RESULTS.append((group, name, bool(ok), detail))
    print(f"  {'PASS' if ok else 'FAIL'}  [{group}] {name}" + (f" — {detail}" if detail else ""), flush=True)


def login():
    global COOKIE
    if args.cookie_jar:
        rows = [l.rstrip("\n").split("\t") for l in open(args.cookie_jar) if l.count("\t") == 6]
        COOKIE = "; ".join(f"{r[5]}={r[6]}" for r in rows)
    else:
        code = os.environ.get("ACCESS_CODE")
        if not code:
            sys.exit("Thiếu ACCESS_CODE (biến môi trường) hoặc --cookie-jar")
        st, _, hd = http("POST", "/api/session", json.dumps({"code": code}).encode(), cookie="")
        if st != 200:
            sys.exit(f"Đăng nhập thất bại ({st})")
        COOKIE = (hd.get("Set-Cookie") or hd.get("set-cookie") or "").split(";")[0]
    st, b, _ = http("GET", "/api/session")
    if not J(b).get("authed"):
        sys.exit("Phiên đăng nhập không hợp lệ")


# ---------- dựng file .xlsx tối giản (ô dạng inlineStr) ----------
def cell(ref, v):
    if v is None:
        return ""
    if isinstance(v, str) and v.startswith("="):
        return f'<c r="{ref}"><f>{xml(v[1:])}</f></c>'
    if isinstance(v, (int, float)):
        return f'<c r="{ref}"><v>{v}</v></c>'
    return f'<c r="{ref}" t="inlineStr"><is><t>{xml(str(v))}</t></is></c>'


def xml(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def col(i):
    s = ""
    i += 1
    while i:
        i, m = divmod(i - 1, 26)
        s = chr(65 + m) + s
    return s


def sheet_xml(rows, extra=""):
    """rows: list các dòng (list giá trị) bắt đầu từ dòng 1; None = bỏ qua dòng đó."""
    out = []
    for i, r in enumerate(rows, 1):
        if r is None:
            continue
        out.append(f'<row r="{i}">' + "".join(cell(f"{col(j)}{i}", v) for j, v in enumerate(r)) + "</row>")
    return ('<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            f'<sheetData>{"".join(out)}</sheetData>{extra}</worksheet>')


def workbook(sheets, raw_entries=None):
    """sheets: dict tên -> XML sheet. Trả về bytes của file .xlsx."""
    bio = io.BytesIO()
    with zipfile.ZipFile(bio, "w", zipfile.ZIP_DEFLATED) as z:
        names = list(sheets)
        z.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
                   '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                   + "".join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' for i in range(1, len(names) + 1)) + "</Types>")
        z.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        z.writestr("xl/workbook.xml", '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
                   + "".join(f'<sheet name="{xml(n)}" sheetId="{i}" r:id="rId{i}"/>' for i, n in enumerate(names, 1)) + "</sheets></workbook>")
        z.writestr("xl/_rels/workbook.xml.rels", '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   + "".join(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>' for i, _ in enumerate(names, 1)) + "</Relationships>")
        for i, n in enumerate(names, 1):
            z.writestr(f"xl/worksheets/sheet{i}.xml", sheets[n])
        for name, data in (raw_entries or {}).items():
            z.writestr(name, data)
    return bio.getvalue()


HDR = ["tiêu đề"], ["cột"]
POS = [("X_VT1", "Chủ tịch UBND xã"), ("X_VT2", "Chủ tịch HĐND xã"), ("X_VT3", "Phó Chủ tịch UBND xã"), ("X_VT4", "Phó Chủ tịch HĐND xã"),
       ("X_VT5", "Trưởng ban Kinh tế - Ngân sách"), ("X_VT6", "Phó Trưởng ban Văn hóa - Xã hội"), ("X_VT7", "Chánh Văn phòng HĐND và UBND xã"), ("X_VT8", "Chuyên viên")]
HD = {"X_VT2", "X_VT4", "X_VT5", "X_VT6"}


def standard_sheets(pos=POS, d1_extra=(), c1_extra=""):
    a2 = [[f"CV{i}", f"NV{(i + 1) // 2}", p[0], f"việc {i}"] for i, p in enumerate(pos, 1)]
    return {
        "1. A1_NhiemVu": sheet_xml([*HDR] + [[f"NV{i}", f"nhiệm vụ {i}", "UBND xã X"] for i in range(1, (len(pos) + 1) // 2 + 1)]),
        "2. A2_CongViec": sheet_xml([*HDR] + a2),
        "3. B1_KhoiLuong": sheet_xml([*HDR] + [[r[0], None, 10, 30, None, "=C{0}*D{0}/60".format(i)] for i, r in enumerate(a2, 3)]),
        "6. C1_VTVL": sheet_xml([*HDR] + [[c, n, "Công chức"] for c, n in pos], c1_extra),
        "7. C2_NangLuc": sheet_xml([*HDR] + [[c, f"trách nhiệm của {n}"] for c, n in pos]),
        "8. C3_NhuCau": sheet_xml([*HDR] + [[c, "PP1", 100] for c, _ in pos]),
        "9. D1_NhanLuc": sheet_xml([*HDR] + [[f"NL{i}", c, "Cán bộ"] for i, (c, _) in enumerate(pos, 1)] + [list(x) for x in d1_extra]),
        "10. D2_PhuongAn": sheet_xml([*HDR] + [[c, y, 1] for c, _ in pos for y in (2027, 2028)]),
        "Ghi chu": sheet_xml([["không tách"]]),
    }


# ---------- thao tác với API ----------
def new_job(name, size):
    st, b, _ = http("POST", "/api/jobs", json.dumps({"files": [{"name": name, "size": size}]}).encode())
    return J(b)["jobs"][0]["id"] if st == 200 else None


def upload(jid, data, workers=4):
    total = max(1, math.ceil(len(data) / PART))
    def one(i):
        return http("PUT", f"/api/jobs/{jid}/upload?part={i}&total={total}", data[i * PART:(i + 1) * PART], "application/octet-stream")[0]
    with ThreadPoolExecutor(workers) as ex:
        sts = list(ex.map(one, range(total)))
    return total, sts


def submit(name, data, sha=True):
    """Tạo job + tải lên + đọc vị trí. Trả về (status http của inspect, body json, job id)."""
    jid = new_job(name, len(data))
    total, _ = upload(jid, data)
    body = {"parts": total}
    if sha:
        body["sha256"] = hashlib.sha256(data).hexdigest()
    st, b, _ = http("POST", f"/api/jobs/{jid}/inspect", json.dumps(body).encode())
    return st, J(b), jid


def wait_done(jid, timeout=180):
    if not jid:
        return {"status": "không tạo được tác vụ"}
    t = time.time()
    while time.time() - t < timeout:
        j = J(http("GET", f"/api/jobs/{jid}")[1]).get("job") or {}
        if j.get("status") in ("done", "error"):
            return j
        time.sleep(0.5)
    return {"status": "timeout"}


def split_and_fetch(jid, mapping=None):
    http("POST", f"/api/jobs/{jid}/split", json.dumps({"mapping": mapping or {}}).encode())
    j = wait_done(jid)
    files = {}
    if j.get("status") == "done":
        for org in ("HDND", "UBND"):
            files[org] = http("GET", f"/api/jobs/{jid}/download?org={org}")[1]
    return j, files


def keys_of(xlsx_bytes, sheet, colidx=0):
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), read_only=True)
    return [r[colidx] for r in wb[sheet].iter_rows(min_row=3, values_only=True) if r and len(r) > colidx and r[colidx] not in (None, "")]


def cleanup(*jids):
    for j in jids:
        if j:
            http("DELETE", f"/api/jobs/{j}")


def run(group):
    def deco(fn):
        if args.only and args.only.lower() not in group.lower():
            return fn
        print(f"\n== {group} ==", flush=True)
        global CLIENT_IP
        CLIENT_IP = f"198.51.100.{sum(map(ord, group)) % 250 + 1}"
        try:
            fn(group)
        except Exception as e:  # một nhóm lỗi không dừng cả bộ test
            check(group, "nhóm test chạy hết không văng lỗi", False, f"{type(e).__name__}: {e}")
        return fn
    return deco


login()
print(f"Máy chủ: {U}")

# =====================================================================
@run("1. Xác thực")
def _(g):
    check(g, "không cookie → 401", http("POST", "/api/jobs", b'{"files":[]}', cookie="")[0] == 401)
    check(g, "cookie giả → 401", http("GET", "/api/jobs?ids=x", cookie="vtvl_split_session=9999999999999.abc.fake")[0] == 401)
    name, val = COOKIE.split("=", 1)
    payload, sig = val.rsplit(".", 1)
    forged = f"{name}=9{payload}.{sig}"  # sửa hạn dùng, giữ chữ ký
    check(g, "cookie bị sửa hạn dùng → 401", http("GET", "/api/jobs?ids=x", cookie=forged)[0] == 401)
    check(g, "tải kết quả không cookie → 401", http("GET", "/api/jobs/00000000-0000-4000-8000-000000000000/download?org=HDND", cookie="")[0] == 401)
    ip = {"X-Forwarded-For": "203.0.113.77"}
    sts = [http("POST", "/api/session", json.dumps({"code": f"sai-{i}"}).encode(), cookie="", headers=ip)[0] for i in range(12)]
    check(g, "sai mã 12 lần → bị khóa tạm (429)", sts[:10] == [401] * 10 and sts[-1] == 429, str(sts))
    check(g, "body đăng nhập không phải JSON → 401, không 500", http("POST", "/api/session", b"\x00\xff{", cookie="", headers={"X-Forwarded-For": "203.0.113.78"})[0] == 401)
    check(g, "trang chủ có header chống nhúng iframe", http("GET", "/", cookie="")[2].get("x-frame-options", http("GET", "/", cookie="")[2].get("X-Frame-Options", "")).upper() == "DENY")


@run("2. Tạo tác vụ")
def _(g):
    ip = {"X-Forwarded-For": "203.0.113.80"}
    def post(body, raw=False):
        return http("POST", "/api/jobs", body if raw else json.dumps(body).encode(), headers=ip)
    f = lambda n="a.xlsx", s=10: {"name": n, "size": s}
    cases = [
        ("body rỗng", b"", True, 400), ("JSON hỏng", b"{{{", True, 400), ("files không phải mảng", {"files": "x"}, False, 400), ("0 file", {"files": []}, False, 400),
        ("4 file", {"files": [f(f"{i}.xlsx") for i in range(4)]}, False, 400), ("file .xls", {"files": [f("a.xls")]}, False, 400), ("file .xlsm", {"files": [f("a.xlsm")]}, False, 400),
        ("file tạm ~$", {"files": [f("~$a.xlsx")]}, False, 400), ("size 0", {"files": [f(s=0)]}, False, 400), ("size âm", {"files": [f(s=-5)]}, False, 400),
        ("size là chuỗi", {"files": [{"name": "a.xlsx", "size": "10"}]}, False, 400), ("size 61 MB", {"files": [f(s=61 * 1024 * 1024)]}, False, 400),
        ("tên dài 300 ký tự", {"files": [f("a" * 300 + ".xlsx")]}, False, 400), ("phần tử null", {"files": [None]}, False, 400), ("tên là số", {"files": [{"name": 5, "size": 10}]}, False, 400),
    ]
    for name, body, raw, exp in cases:
        st = post(body, raw)[0]
        check(g, f"{name} → {exp}", st == exp, f"nhận {st}")
    st, b, _ = post({"files": [f("../../etc/passwd.xlsx"), f("a\r\nSet-Cookie: x=1.xlsx"), f("<img src=x onerror=alert(1)>.xlsx")]})
    jobs = J(b).get("jobs", [])
    check(g, "tên file nguy hiểm vẫn tạo được job với id ngẫu nhiên", st == 200 and len(jobs) == 3 and all(len(j["id"]) == 36 for j in jobs))
    cleanup(*[j["id"] for j in jobs])
    sts = [post({"files": [f("1.xlsx"), f("2.xlsx"), f("3.xlsx")]}) for _ in range(34)]
    created = [j["id"] for st, b, _ in sts for j in J(b).get("jobs", [])]
    check(g, "tạo dồn dập 102 file trong 1 phút từ một IP → bị chặn (429)", sts[-1][0] == 429 and len(created) <= 90, f"tạo được {len(created)}")
    cleanup(*created)


@run("3. Tải lên từng phần")
def _(g):
    bad = ["abc", "../../x", "00000000-0000-4000-8000-000000000000"]
    check(g, "id sai/không tồn tại → 404", all(http("PUT", f"/api/jobs/{i}/upload?part=0&total=1", b"x", "application/octet-stream")[0] == 404 for i in bad))
    jid = new_job("a.xlsx", 1)
    q = lambda p, t, d=b"x": http("PUT", f"/api/jobs/{jid}/upload?part={p}&total={t}", d, "application/octet-stream")[0]
    for name, p, t in [("part âm", -1, 2), ("part = total", 2, 2), ("part chữ", "a", 2), ("total 0", 0, 0), ("total 1.5", 0, 1.5), ("total 9999", 0, 9999), ("thiếu tham số", "", "")]:
        check(g, f"{name} → 400", q(p, t) == 400)
    check(g, "phần rỗng → 400", q(0, 1, b"") == 400)
    check(g, "phần 5 MB → bị từ chối", q(0, 2, b"\0" * (5 * 1024 * 1024)) in (413, 0))
    check(g, "total vượt số phần tối đa của file 60 MB → 400", q(0, 40) == 400, "chặn dồn dung lượng vượt giới hạn")
    check(g, "ghi đè cùng một phần 2 lần → 200", q(0, 1, b"a") == 200 and q(0, 1, b"b") == 200)
    st, b, _ = http("POST", f"/api/jobs/{jid}/inspect", json.dumps({"parts": 1}).encode())
    check(g, "nội dung rác → lỗi rõ, không 500", st == 422 and "Excel" in msg(b), msg(b))
    check(g, "tải thêm sau khi đã kiểm tra → 409", q(0, 1) == 409)
    cleanup(jid)
    jid = new_job("b.xlsx", 100)
    st, b, _ = http("POST", f"/api/jobs/{jid}/inspect", b"{}")
    check(g, "kiểm tra khi chưa tải gì → lỗi rõ", st == 422 and "Chưa nhận được file" in msg(b), msg(b))
    cleanup(jid)


@run("4. Toàn vẹn file tải lên")
def _(g):
    data = workbook(standard_sheets()) + b""
    big = data + os.urandom(0)
    jid = new_job("a.xlsx", len(data)); upload(jid, data)
    st, b, _ = http("POST", f"/api/jobs/{jid}/inspect", json.dumps({"parts": 1, "sha256": "0" * 64}).encode())
    check(g, "SHA-256 không khớp → từ chối", st == 422 and "không khớp" in msg(b), msg(b)); cleanup(jid)
    jid = new_job("a.xlsx", len(data) + 7); upload(jid, data)
    st, b, _ = http("POST", f"/api/jobs/{jid}/inspect", json.dumps({"parts": 1}).encode())
    check(g, "kích thước khai báo khác thực tế → từ chối", st == 422 and "dung lượng" in msg(b), msg(b)); cleanup(jid)
    jid = new_job("a.xlsx", len(data)); upload(jid, data)
    st, b, _ = http("POST", f"/api/jobs/{jid}/inspect", json.dumps({"parts": 3}).encode())
    check(g, "khai 3 phần, chỉ có 1 → từ chối", st == 422 and "chưa đủ" in msg(b), msg(b)); cleanup(jid)
    st, b, jid = submit("ok.xlsx", data)
    check(g, "file đúng + SHA-256 đúng → sẵn sàng", st == 200 and b["job"]["status"] == "ready", msg(b)); cleanup(jid)


@run("5. File lạ / hỏng / độc hại")
def _(g):
    std = standard_sheets()
    def expect_error(name, data, needle="", allow_ready=False):
        t = time.time()
        st, b, jid = submit(name + ".xlsx", data)
        dt = time.time() - t
        ok = st in (200, 422) and (b.get("job", {}).get("status") == "error" or (allow_ready and b.get("job", {}).get("status") == "ready")) and (needle in msg(b) if needle else True) and dt < 10
        check(g, name, ok, f"http {st}, {dt:.1f}s, “{msg(b)[:90]}”")
        cleanup(jid)
        return b
    expect_error("file .xls cũ (OLE) đổi đuôi", b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\0" * 600, "không phải file Excel")
    expect_error("file zip rỗng", b"PK\x05\x06" + b"\0" * 18, "")
    expect_error("file zip bị cắt cụt", workbook(std)[:900], "")
    z = io.BytesIO()
    with zipfile.ZipFile(z, "w") as zz:
        zz.writestr("a.txt", "x")
    expect_error("zip thường (không phải xlsx)", z.getvalue(), "Excel")
    expect_error("xlsx không có biểu C1", workbook({"Sheet1": sheet_xml([["a"]])}), "C1_VTVL")
    s = dict(std); s["6. C1_VTVL"] = sheet_xml([*HDR])
    expect_error("C1 không có vị trí nào", workbook(s), "chưa có vị trí")
    # entry sheet bị hỏng dữ liệu nén
    raw = bytearray(workbook(std)); i = raw.find(b"xl/worksheets/sheet4.xml") + 40; raw[i:i + 30] = os.urandom(30)
    expect_error("dữ liệu nén của biểu C1 bị hỏng", bytes(raw), "")
    # bom nén: một sheet 1,2 GB số 0 nén còn ~1 MB
    bomb = io.BytesIO(); src = zipfile.ZipFile(io.BytesIO(workbook(std)))
    with zipfile.ZipFile(bomb, "w", zipfile.ZIP_DEFLATED) as zz:
        for info in src.infolist():
            if info.filename == "xl/worksheets/sheet4.xml":
                with zz.open(info.filename, "w") as w:
                    blk = b"\0" * (4 << 20)
                    for _ in range(300):
                        w.write(blk)
            else:
                zz.writestr(info.filename, src.read(info.filename))
    expect_error("bom nén: biểu C1 phình 1,2 GB", bomb.getvalue(), "quá lớn")
    # thứ tự thuộc tính ô khác Excel (s trước r) – phải báo lỗi rõ, KHÔNG được âm thầm bỏ qua
    s = dict(std); s["6. C1_VTVL"] = s["6. C1_VTVL"].replace('<c r="', '<c s="0" r="')
    b = expect_error("ô có thứ tự thuộc tính lạ (phần mềm khác lưu)", workbook(s), "")
    # dòng thiếu số thứ tự / sai thứ tự
    s = dict(std); s["6. C1_VTVL"] = s["6. C1_VTVL"].replace('<row r="4">', '<row r="2">')
    expect_error("số thứ tự dòng bị lùi", workbook(s), "")
    # XML dị dạng: rất nhiều thẻ mở không đóng (thử làm treo máy chủ) – phải trả lời nhanh, không treo
    s = dict(std); s["6. C1_VTVL"] = s["6. C1_VTVL"].replace("</sheetData>", '<row r="50">' + '<c r="A50">' * 300000 + "</sheetData>")
    expect_error("XML dị dạng 300.000 thẻ ô không đóng (không được treo)", workbook(s), "", allow_ready=True)
    s = dict(std); s["6. C1_VTVL"] = s["6. C1_VTVL"].replace("</sheetData>", '<row r="50">' * 300000 + '<row r="51"><c r="A51"><v>1</v></c></row></sheetData>')
    expect_error("XML dị dạng 300.000 thẻ dòng không đóng (không được treo)", workbook(s), "", allow_ready=True)
    s = dict(std); s["9. D1_NhanLuc"] = s["9. D1_NhanLuc"].replace("</row></sheetData>", '<c r="Z9">' * 300000 + "</row></sheetData>")
    t = time.time(); st, b, jid = submit("d1-di-dang.xlsx", workbook(s)); j, _f = split_and_fetch(jid)
    check(g, "biểu D1 dị dạng 300.000 thẻ ô → tác vụ báo lỗi nhanh, không treo", j.get("status") == "error" and time.time() - t < 20, f"{j.get('status')} {time.time() - t:.1f}s “{(j.get('error') or {}).get('message', '')[:70]}”")
    cleanup(jid)
    # workbook trỏ tới sheet không tồn tại / đường dẫn lạ
    raw = workbook(std, {"xl/_rels/extra.txt": "x"})
    wbz = zipfile.ZipFile(io.BytesIO(raw)); out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zz:
        for i in wbz.infolist():
            d = wbz.read(i.filename)
            if i.filename == "xl/_rels/workbook.xml.rels":
                d = d.replace(b"worksheets/sheet4.xml", b"../../../../etc/passwd")
            zz.writestr(i.filename, d)
    expect_error("đường dẫn sheet trỏ ra ngoài (../../etc/passwd)", out.getvalue(), "C1_VTVL")
    # mã/tên chứa HTML, tên rất dài, 3.000 vị trí
    pos = [("<script>alert(1)</script>", "Chủ tịch HĐND <b>xã</b>"), ("X2", "Chủ tịch UBND " + "x" * 20000)] + [(f"P{i}", "Chuyên viên") for i in range(3000)]
    t = time.time(); st, b, jid = submit("nhieu-vi-tri.xlsx", workbook(standard_sheets(pos)))
    n = len(b.get("job", {}).get("positions", []))
    check(g, "3.002 vị trí, mã chứa HTML, tên 20.000 ký tự → đọc được", st == 200 and n == 3002, f"{n} vị trí, {time.time() - t:.1f}s")
    j, files = split_and_fetch(jid)
    check(g, "…và tách được", j.get("status") == "done" and len(keys_of(files.get("HDND", b""), "6. C1_VTVL")) == 1 if files else False, str(j.get("error") or j.get("status")))
    cleanup(jid)


@run("6. Tách: quy tắc và dữ liệu biên")
def _(g):
    # file do openpyxl tạo (phần mềm khác Excel) – shared strings, kiểu XML khác
    import openpyxl
    wb = openpyxl.Workbook(); wb.remove(wb.active)
    def add(name, rows):
        ws = wb.create_sheet(name); ws.append(["tiêu đề"]); ws.append(["cột"])
        for r in rows:
            ws.append(r)
        return ws
    a2 = [[f"CV{i}", f"NV{(i + 1) // 2}", p[0], f"việc {i}"] for i, p in enumerate(POS, 1)]
    add("1. A1_NhiemVu", [[f"NV{i}", f"nhiệm vụ {i}"] for i in range(1, 5)] + [["NV9", "chưa gắn công việc"]])
    add("2. A2_CongViec", a2)
    b1 = add("3. B1_KhoiLuong", [[r[0], None, 10, 30] for r in a2])
    for i in range(3, 3 + len(a2)):
        b1[f"B{i}"] = f"=IFERROR(INDEX('2. A2_CongViec'!$C$3:$C$100,MATCH(A{i},'2. A2_CongViec'!$A$3:$A$100,0)),\"\")"
        b1[f"F{i}"] = f'=IF(A{i}="","",C{i}*D{i}/60)'
    add("6. C1_VTVL", [[c, n] for c, n in POS])
    add("7. C2_NangLuc", [[c, "mô tả"] for c, _ in POS])
    add("8. C3_NhuCau", [[c, "PP1"] for c, _ in POS])
    d1 = add("9. D1_NhanLuc", [[f"NL{i}", c] for i, (c, _) in enumerate(POS, 1)] + [["NL99", "MA_LA"], ["ghi chú không có mã", None]])
    add("10. D2_PhuongAn", [[c, y] for c, _ in POS for y in (2027, 2028)])
    wb.create_sheet("Khac").append(["giữ nguyên"])
    bio = io.BytesIO(); wb.save(bio); data = bio.getvalue()
    st, b, jid = submit("02.HDND_UBND Xã Thử.xlsx", data)
    job = b.get("job", {})
    sug = {p["code"]: p["suggested"] for p in job.get("positions", [])}
    check(g, "file do openpyxl tạo → đọc được 8 vị trí", st == 200 and len(sug) == 8, msg(b))
    check(g, "gợi ý đúng 4 vị trí HĐND", {c for c, o in sug.items() if o == "HDND"} == HD, str(sug))
    j, files = split_and_fetch(jid)
    check(g, "tách xong", j.get("status") == "done", str(j.get("error")))
    if files:
        H, Ub = files["HDND"], files["UBND"]
        check(g, "tên file kết quả", j["outputs"]["HDND"]["name"] == "02.HĐND Xã Thử.xlsx" and j["outputs"]["UBND"]["name"] == "02.UBND Xã Thử.xlsx", str(j["outputs"]))
        check(g, "C1: HĐND đúng 4 mã, UBND 4 mã, đúng thứ tự", keys_of(H, "6. C1_VTVL") == ["X_VT2", "X_VT4", "X_VT5", "X_VT6"] and keys_of(Ub, "6. C1_VTVL") == ["X_VT1", "X_VT3", "X_VT7", "X_VT8"])
        check(g, "A2 chia theo mã VTVL", keys_of(H, "2. A2_CongViec") == ["CV2", "CV4", "CV5", "CV6"] and keys_of(Ub, "2. A2_CongViec") == ["CV1", "CV3", "CV7", "CV8"])
        check(g, "B1 chia theo mã công việc (cột VTVL là công thức)", keys_of(H, "3. B1_KhoiLuong") == ["CV2", "CV4", "CV5", "CV6"] and keys_of(Ub, "3. B1_KhoiLuong") == ["CV1", "CV3", "CV7", "CV8"])
        import openpyxl as ox
        hb = ox.load_workbook(io.BytesIO(H)); ws = hb["3. B1_KhoiLuong"]
        check(g, "công thức B1 dịch đúng dòng sau khi dồn", ws["F3"].value == '=IF(A3="","",C3*D3/60)' and "MATCH(A4," in ws["B4"].value, f"{ws['F3'].value}")
        check(g, "dòng thừa đã trống, công thức mẫu còn", ws["A7"].value is None and ws["C7"].value is None and str(ws["F7"].value).startswith("=IF(A7"), f"A7={ws['A7'].value} F7={ws['F7'].value}")
        check(g, "A1: NV1 dùng chung (VT1+VT2) có ở cả hai; NV9 chưa gắn có ở cả hai", "NV1" in keys_of(H, "1. A1_NhiemVu") and "NV1" in keys_of(Ub, "1. A1_NhiemVu") and "NV9" in keys_of(H, "1. A1_NhiemVu") and "NV9" in keys_of(Ub, "1. A1_NhiemVu"), f"{keys_of(H, '1. A1_NhiemVu')}")
        check(g, "A1: NV3 (VT5+VT6, chỉ HĐND) chỉ ở file HĐND", "NV3" in keys_of(H, "1. A1_NhiemVu") and "NV3" not in keys_of(Ub, "1. A1_NhiemVu"))
        check(g, "D1: mã lạ + dòng không mã giữ ở cả hai, có cảnh báo", "NL99" in keys_of(H, "9. D1_NhanLuc") and "NL99" in keys_of(Ub, "9. D1_NhanLuc") and len(j["report"]["warnings"]) >= 2, str(j["report"]["warnings"]))
        check(g, "D2: 2 dòng/vị trí chia đúng", keys_of(H, "10. D2_PhuongAn") == ["X_VT2", "X_VT2", "X_VT4", "X_VT4", "X_VT5", "X_VT5", "X_VT6", "X_VT6"])
        check(g, "sheet khác giữ nguyên", hb["Khac"]["A1"].value == "giữ nguyên" and ox.load_workbook(io.BytesIO(Ub))["Khac"]["A1"].value == "giữ nguyên")
        check(g, "bật tính lại công thức khi mở", b'fullCalcOnLoad="1"' in zipfile.ZipFile(io.BytesIO(H)).read("xl/workbook.xml"))
    # mapping do người dùng sửa
    j2, f2 = split_and_fetch(jid, {"X_VT6": "UBND", "X_VT8": "HDND"})
    check(g, "tách lại với bảng đã sửa (VT6→UBND, VT8→HĐND)", j2.get("status") == "done" and keys_of(f2["HDND"], "6. C1_VTVL") == ["X_VT2", "X_VT4", "X_VT5", "X_VT8"], str(j2.get("error")))
    full = lambda v: {c: v for c, _ in POS}
    for name, m, exp in [("toàn UBND", full("UBND"), 400), ("toàn HĐND", full("HDND"), 400), ("giá trị lạ", {**full("UBND"), "X_VT1": "ABC"}, 400)]:
        st = http("POST", f"/api/jobs/{jid}/split", json.dumps({"mapping": m}).encode())[0]
        check(g, f"mapping {name} → {exp}", st == exp, f"nhận {st}")
    for name, m in [("là mảng (dùng lại bảng đã lưu)", [1, 2]), ("có mã không tồn tại (bỏ qua mã thừa)", {"KHONG_CO": "HDND"})]:
        st = http("POST", f"/api/jobs/{jid}/split", json.dumps({"mapping": m}).encode())[0]; jx = wait_done(jid)
        check(g, f"mapping {name} → vẫn tách được", st == 200 and jx.get("status") == "done", f"{st} {jx.get('status')}")
    cleanup(jid)
    # ô gộp trong vùng dữ liệu, dòng trống xen giữa, mã trùng, thiếu biểu
    s = standard_sheets(c1_extra='<mergeCells count="1"><mergeCell ref="B4:C5"/></mergeCells>')
    st, b, jid = submit("o-gop.xlsx", workbook(s)); j, _f = split_and_fetch(jid)
    check(g, "ô gộp trong vùng dữ liệu → có cảnh báo", j.get("status") == "done" and any("gộp" in w for w in j["report"]["warnings"]), str(j.get("report", {}).get("warnings")))
    cleanup(jid)
    rows = [*HDR, ["X_VT1", "Chủ tịch UBND xã"], None, None, ["X_VT2", "Chủ tịch HĐND xã"], ["X_VT2", "Chủ tịch HĐND xã (trùng)"], None, ["X_VT3", "Chuyên viên"]]
    s = standard_sheets(POS[:3]); s["6. C1_VTVL"] = sheet_xml(rows); del s["3. B1_KhoiLuong"]; del s["2. A2_CongViec"]
    st, b, jid = submit("khuyet.xlsx", workbook(s)); job = b.get("job", {})
    check(g, "thiếu biểu A2, B1 → vẫn đọc được và có cảnh báo thiếu biểu", st == 200 and set(job.get("missingSheets", [])) >= {"A2", "B1"}, str(job.get("missingSheets")))
    j, files = split_and_fetch(jid)
    check(g, "dòng trống xen giữa + mã trùng → dồn liền, giữ cả 2 dòng trùng", j.get("status") == "done" and keys_of(files["HDND"], "6. C1_VTVL") == ["X_VT2", "X_VT2"] and keys_of(files["UBND"], "6. C1_VTVL") == ["X_VT1", "X_VT3"], str(j.get("error")))
    cleanup(jid)


@run("7. Trạng thái tác vụ và tải về")
def _(g):
    data = workbook(standard_sheets())
    jid = new_job("a.xlsx", len(data))
    check(g, "tách khi chưa đọc vị trí → 409", http("POST", f"/api/jobs/{jid}/split", b"{}")[0] == 409)
    check(g, "tải về khi chưa xong → 404", http("GET", f"/api/jobs/{jid}/download?org=HDND")[0] == 404)
    cleanup(jid)
    st, b, jid = submit('ten "la" \r\n X-Hack: 1 HĐND-UBND.xlsx', data)
    # bấm Tách 5 lần cùng lúc
    with ThreadPoolExecutor(5) as ex:
        sts = list(ex.map(lambda _: http("POST", f"/api/jobs/{jid}/split", b"{}")[0], range(5)))
    j = wait_done(jid)
    check(g, "bấm Tách 5 lần cùng lúc → không lỗi, xong 1 kết quả", set(sts) == {200} and j.get("status") == "done", f"{sts} {j.get('status')} lần thử={j.get('attempts')}")
    st, body, hd = http("GET", f"/api/jobs/{jid}/download?org=HDND")
    cd = hd.get("content-disposition") or hd.get("Content-Disposition") or ""
    check(g, "tên file lạ (nháy kép, xuống dòng) → tải được, header an toàn", st == 200 and body[:2] == b"PK" and "\n" not in cd and "X-Hack" not in hd, cd[:80])
    check(g, "org sai → 400", http("GET", f"/api/jobs/{jid}/download?org=ABC")[0] == 400 and http("GET", f"/api/jobs/{jid}/download")[0] == 400)
    check(g, "xem nhiều job: id rác bị bỏ qua", http("GET", f"/api/jobs?ids={jid},abc,../x," + ",".join(["00000000-0000-4000-8000-00000000000%d" % i for i in range(9)]))[0] == 200)
    cleanup(jid)
    check(g, "sau khi xóa → 404 và tải về 404", http("GET", f"/api/jobs/{jid}")[0] == 404 and http("GET", f"/api/jobs/{jid}/download?org=HDND")[0] == 404)
    check(g, "xóa lần 2 / xóa id rác → không lỗi", http("DELETE", f"/api/jobs/{jid}")[0] == 200 and http("DELETE", "/api/jobs/abc")[0] == 200)
    # xóa trong lúc đang xử lý: không được để lại file mồ côi, máy chủ không lỗi
    st, b, jid = submit("xoa-giua-chung.xlsx", data)
    http("POST", f"/api/jobs/{jid}/split", b"{}"); cleanup(jid); time.sleep(3)
    check(g, "xóa ngay sau khi bấm Tách → tác vụ biến mất hẳn", http("GET", f"/api/jobs/{jid}")[0] == 404 and http("GET", f"/api/jobs/{jid}/download?org=HDND")[0] == 404)
    check(g, "đường dẫn không tồn tại → 404 (không 500)", http("GET", "/api/khong-co")[0] == 404 and http("POST", "/api/jobs/abc/inspect", b"{}")[0] == 404)


@run("8. Tải nặng và đồng thời")
def _(g):
    if not args.sample:
        print("  (bỏ qua: không có --sample)"); return
    data = open(args.sample, "rb").read(); sha = hashlib.sha256(data).hexdigest()
    lat = []; stop = threading.Event()
    def ping():
        while not stop.is_set():
            t = time.time(); http("GET", "/api/session"); lat.append(time.time() - t); time.sleep(0.2)
    th = threading.Thread(target=ping); th.start()
    t0 = time.time(); jids = []
    def one(i):
        jid = new_job(f"{i:02d}.HĐND-UBND file {i}.xlsx", len(data)); total, sts = upload(jid, data)
        st, b, _ = http("POST", f"/api/jobs/{jid}/inspect", json.dumps({"parts": total, "sha256": sha}).encode())
        http("POST", f"/api/jobs/{jid}/split", b"{}"); return jid
    with ThreadPoolExecutor(6) as ex:
        jids = list(ex.map(one, range(6)))
    js = [wait_done(j, 300) for j in jids]; total_t = time.time() - t0
    stop.set(); th.join()
    check(g, f"6 file {len(data) / 1048576:.0f} MB gửi cùng lúc (2 lượt × 3) → xong cả 6", all(j.get("status") == "done" for j in js), f"{total_t:.0f}s, mỗi file {[round(j.get('durationMs', 0) / 1000, 1) for j in js]}s")
    hs = {hashlib.md5(http("GET", f"/api/jobs/{j}/download?org=HDND")[1]).hexdigest() for j in jids}
    check(g, "6 kết quả HĐND giống hệt nhau (không lẫn dữ liệu giữa các tác vụ)", len(hs) == 1)
    lat.sort()
    check(g, "máy chủ vẫn phản hồi trong lúc tách (trễ lớn nhất < 5s)", lat and lat[-1] < 5, f"trung vị {lat[len(lat) // 2] * 1000:.0f} ms, lớn nhất {lat[-1] * 1000:.0f} ms, {len(lat)} lần hỏi")
    cleanup(*jids)


@run("9. Dọn file quá hạn")
def _(g):
    st, b, _ = http("GET", "/api/cron/cleanup", cookie="")
    check(g, "gọi dọn file → 200 hoặc 401 (nếu có CRON_SECRET)", st in (200, 401), f"{st} {b[:80]!r}")
    data = workbook(standard_sheets()); st2, b2, jid = submit("moi.xlsx", data)
    if st == 200:
        http("GET", "/api/cron/cleanup", cookie="")
        check(g, "tác vụ mới tạo KHÔNG bị dọn", http("GET", f"/api/jobs/{jid}")[0] == 200)
    cleanup(jid)


# =====================================================================
fails = [r for r in RESULTS if not r[2]]
print(f"\n===== TỔNG KẾT: {len(RESULTS) - len(fails)}/{len(RESULTS)} PASS, {len(fails)} FAIL =====")
for g, n, _, d in fails:
    print(f"  FAIL [{g}] {n} — {d}")
sys.exit(1 if fails else 0)
