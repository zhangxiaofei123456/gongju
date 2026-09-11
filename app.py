# -*- coding: utf-8 -*-
"""
考勤核对 Web 系统 v3
- 解析总表 xlsx（表头第4行、每员工2行 上午/下午、按"部门"列精确分组）
- 管理员：上传总表 / 按部门建主管账号 / 看凭证 / 改单元格 / 导出回写总表
- 主管：看本区一览表 / 详情传照片凭证+说明(变黄) / 看回写总表确认
- 状态色：黄=待审(主管已提交) 红=有问题 绿=无问题
仅依赖 Python 标准库 + openpyxl（首次自动安装）
"""
import sqlite3, json, os, base64, io, hashlib, secrets, datetime, re, traceback, sys
import urllib.parse
import http.server

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "att.db")
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

# 上线部署：可用环境变量 ADMIN_PASS 覆盖默认密码（强烈建议部署时设置）
ADMIN_PASS = os.environ.get("ADMIN_PASS", "admin123")
# 云托管平台（Render/Railway 等）会注入 PORT 环境变量，必须读取
PORT = int(os.environ.get("PORT", 8000))

# ----------------------------------------------------------------------------
# openpyxl 懒加载（首次自动 pip 安装）
# ----------------------------------------------------------------------------
openpyxl = None
def ensure_openpyxl():
    global openpyxl
    if openpyxl: return True
    try:
        import openpyxl as _m
        openpyxl = _m
        return True
    except Exception:
        try:
            import subprocess, sys
            subprocess.run([sys.executable, "-m", "pip", "install", "openpyxl",
                           "-i", "https://pypi.tuna.tsinghua.edu.cn/simple"],
                          check=False, capture_output=True)
            import openpyxl as _m
            openpyxl = _m
            return True
        except Exception:
            return False

# ----------------------------------------------------------------------------
# 工具
# ----------------------------------------------------------------------------
def send_json(self, obj, code=200):
    body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
    self.send_response(code)
    self.send_header("Content-Type", "application/json; charset=utf-8")
    self.send_header("Content-Length", str(len(body)))
    self.send_header("Cache-Control", "no-store")
    self.end_headers()
    self.wfile.write(body)

def send_html(self, html):
    body = html.encode("utf-8")
    self.send_response(200)
    self.send_header("Content-Type", "text/html; charset=utf-8")
    self.send_header("Content-Length", str(len(body)))
    self.send_header("Cache-Control", "no-store")
    self.end_headers()
    self.wfile.write(body)

def read_json_body(self):
    try:
        ln = int(self.headers.get("Content-Length", 0))
        raw = self.rfile.read(ln) if ln else b""
        return json.loads(raw.decode("utf-8"))
    except Exception:
        return {}

def b64_to_bytes(s):
    s = s.strip()
    if "," in s and s.startswith("data:"):
        s = s.split(",", 1)[1]
    return base64.b64decode(s)

def pw_hash(pw):
    return hashlib.sha256(("att_salt_" + pw).encode("utf-8")).hexdigest()

TOKENS = {}  # token -> username

def current_user(self):
    auth = self.headers.get("Authorization", "")
    if auth.startswith("Bearer "):
        tok = auth[7:].strip()
        u = TOKENS.get(tok)
        if u:
            return u
    return None

# ----------------------------------------------------------------------------
# 修改申请 helper
# cell_key 形如 "am_5"；同 cell_key 下可能有多个申请，用递增序号 request_id "am_5_2"
# change_requests 结构：{
#   "am_5_1": {cell_key:"am_5", old:"*", new:"√", reason:"...", evidence_id:12,
#               status:"pending"|"approved"|"rejected",
#               reject_reason:"", created_at, reviewed_at, reviewer:""}
# }
def new_request_id(cell_key, existing):
    n = 1
    while True:
        rid = "%s_%d" % (cell_key, n)
        if rid not in existing: return rid
        n += 1

# ----------------------------------------------------------------------------
# 数据库
# ----------------------------------------------------------------------------
def get_db():
    # timeout=30 + WAL：多主管并发提交时避免 "database is locked"
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA busy_timeout=30000")
    except Exception:
        pass
    return conn

def init_db():
    conn = get_db()
    c = conn.cursor()
    c.executescript("""
    CREATE TABLE IF NOT EXISTS meta (k TEXT PRIMARY KEY, v TEXT);
    CREATE TABLE IF NOT EXISTS employees (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        seq TEXT, dept TEXT, name TEXT,
        am TEXT, pm TEXT, summary TEXT, company TEXT
    );
    CREATE TABLE IF NOT EXISTS users (
        username TEXT PRIMARY KEY, pw TEXT, role TEXT, region TEXT
    );
    CREATE TABLE IF NOT EXISTS submissions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        emp_id INTEGER, supervisor TEXT, note TEXT,
        status TEXT DEFAULT 'pending',
        cells TEXT,
        reviewed_cells TEXT,
        admin_note TEXT,
        created_at TEXT, updated_at TEXT
    );
    CREATE TABLE IF NOT EXISTS evidence (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        submission_id INTEGER, filename TEXT, cell_key TEXT, data BLOB, created_at TEXT
    );
    CREATE TABLE IF NOT EXISTS users_meta (
        id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT, pw TEXT, role TEXT, region TEXT
    );
    """)
    # 迁移：老库升级（submissions 加 cells / reviewed_cells 字段）
    cols = [r[1] for r in c.execute("PRAGMA table_info(submissions)").fetchall()]
    if "cells" not in cols:
        c.execute("ALTER TABLE submissions ADD COLUMN cells TEXT")
    if "reviewed_cells" not in cols:
        c.execute("ALTER TABLE submissions ADD COLUMN reviewed_cells TEXT")
    # 迁移：submissions 增加 change_requests（修改申请 JSON 字典）
    if "change_requests" not in cols:
        c.execute("ALTER TABLE submissions ADD COLUMN change_requests TEXT")
    # evidence 加 cell_key 字段（凭证关联到具体格子）
    ev_cols = [r[1] for r in c.execute("PRAGMA table_info(evidence)").fetchall()]
    if "cell_key" not in ev_cols:
        c.execute("ALTER TABLE evidence ADD COLUMN cell_key TEXT")
    # 迁移：evidence 增加 request_id（绑定到具体"修改申请" cell_key_序号，便于管理员按申请查看凭证）
    if "request_id" not in ev_cols:
        c.execute("ALTER TABLE evidence ADD COLUMN request_id TEXT")
    # 管理员账号（若未存在）
    c.execute("SELECT 1 FROM users WHERE username='admin'")
    if not c.fetchone():
        c.execute("INSERT INTO users (username,pw,role,region) VALUES (?,?,?,?)",
                  ("admin", pw_hash(ADMIN_PASS), "admin", None))
    conn.commit()
    conn.close()

# ----------------------------------------------------------------------------
# 总表解析
# ----------------------------------------------------------------------------
def parse_xlsx(data: bytes):
    if not ensure_openpyxl():
        raise RuntimeError("后端未安装 openpyxl，请先 pip install openpyxl")
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    ws = wb["总表"]
    # 表头第4行
    header = [c.value for c in ws[4]]
    # 日期列 = 表头为整数(1..31)的列
    day_idx = [i for i, h in enumerate(header) if isinstance(h, int)]
    day_numbers = [header[i] for i in day_idx]
    # 汇总列 = 最后一个日期列之后的非空列
    summary_start = (day_idx[-1] + 1) if day_idx else 4
    summary_names = [str(h).replace("\n", "") for h in header[summary_start:]
                      if h not in (None, "")]
    ndays = len(day_idx)
    # 图例（第2行，通常第5列）
    legend_raw = ws.cell(row=2, column=5).value or ""
    legend = parse_legend(legend_raw)
    # 标题/年月
    title = ws.cell(row=1, column=1).value or ""
    m = re.search(r"(\d+)年(\d+)月", str(title))
    year = int(m.group(1)) if m else 2026
    month = int(m.group(2)) if m else 8
    # 数据行
    employees = []
    r = 5
    maxr = ws.max_row
    while r <= maxr:
        dept = ws.cell(row=r, column=2).value
        name = ws.cell(row=r, column=3).value
        if name in (None, "") and dept in (None, ""):
            r += 1
            continue
        t = ws.cell(row=r, column=4).value
        seq = ws.cell(row=r, column=1).value
        am = [ws.cell(row=r, column=i + 1).value for i in day_idx]
        summary = {}
        for sidx, sn in enumerate(summary_names):
            summary[sn] = ws.cell(row=r, column=summary_start + sidx + 1).value
        company = ws.cell(row=r, column=len(header)).value
        # 下午行
        pm = [None] * ndays
        if r + 1 <= maxr:
            t2 = ws.cell(row=r + 1, column=4).value
            if t2 == "下午" or t2 not in (None, ""):
                pm = [ws.cell(row=r + 1, column=i + 1).value for i in day_idx]
                r += 2
            else:
                r += 1
        else:
            r += 1
        employees.append({
            "seq": seq, "dept": dept, "name": name,
            "am": am, "pm": pm, "summary": summary, "company": company
        })
    return {
        "employees": employees,
        "day_numbers": day_numbers,
        "summary_names": summary_names,
        "legend_raw": legend_raw,
        "legend": legend,
        "year": year, "month": month,
        "ndays": ndays,
    }

def parse_legend(s):
    out = []
    for tok in str(s).split(","):
        tok = tok.strip()
        if not tok:
            continue
        if " " in tok:
            sym, mean = tok.split(" ", 1)
            sym, mean = sym.strip(), mean.strip()
        else:
            m = re.search(r"[一-鿿]", tok)
            if m:
                sym = tok[:m.start()].strip()
                mean = tok[m.start():].strip()
            else:
                sym, mean = tok, ""
        if not sym and mean:
            sym = mean[0]
            mean = mean[1:]
        out.append({"symbol": sym, "meaning": mean})
    return out

# ----------------------------------------------------------------------------
# 导出回写总表
# ----------------------------------------------------------------------------
WEEK = ["一", "二", "三", "四", "五", "六", "日"]
def weekday_letter(y, mo, d):
    wd = datetime.date(y, mo, d).weekday()  # Mon=0
    return WEEK[wd]

# ----------------------------------------------------------------------------
# 考勤统计规则（v2026-09-04-stats）
#   用户规则：
#     出差天数 = 出差符号(cc) 出现次数 / 2
#     出勤天数 = 月应出勤天数 - 事假/2 - 缺卡/2
#     其他(缺卡/迟到/早退/调休/旷工/年假/生日假/事假/病假/婚假/丧假/产假(陪)) = 符号次数 / 2
#     "公司" 是文本列，不计入数字统计
# ----------------------------------------------------------------------------
# 符号 -> 统计类别（精确匹配）。类别标签与用户截图一致：14 列简写 + 公司列
SYMBOL_TO_STATS = {
    "cc": "出差天",
    "/":  "缺",
    "！": "迟",
    "!":  "迟",          # 英文感叹号也兼容
    "※": "调",
    "年": "年",
    "●": "事",
    "◇": "产假",
    "×": "旷",          # × 视为旷工
    "△": "丧",
    "▲": "病",
    "★": "婚",
    "▽": "早",
}
# 子串匹配（处理 Excel OCR 脏数据如 "cc出差，生"）
SUBSTR_RULES = (
    ("cc", "出差天"),
    ("生", "生"),
)
STATS_CATEGORIES = [
    "出差天", "出勤天", "缺", "迟", "早", "调", "旷",
    "年", "生", "病", "事", "婚", "丧", "产假",
]

def _stat_symbol_to_category(sym):
    """把单个符号映射成统计类别名；未匹配返回 None。"""
    if sym is None:
        return None
    s = str(sym).strip()
    if not s:
        return None
    if s in SYMBOL_TO_STATS:
        return SYMBOL_TO_STATS[s]
    # 子串兜底（脏数据 "cc出差，生" → 出差天数）
    for sub, cat in SUBSTR_RULES:
        if sub in s:
            return cat
    return None

def compute_employee_stats(am, pm, month_workdays=26):
    """
    输入：员工 am(31)、pm(31)、月应出勤天数(26)
    输出：dict 14 项（与 summary_names 前 14 项对应, 除"公司"外）
      出差天 出勤天 缺 迟 早 调 旷 年 生 病 事 婚 丧 产假
    公式：
      出勤天  = 月应出勤 - 事假天数 - 缺卡天数（不能为负）
      出差天  = 出差符号数 / 2
      其他    = 出现次数 / 2
    """
    cnt = {c: 0 for c in STATS_CATEGORIES if c != "出勤天"}
    for arr in (am or [], pm or []):
        for v in arr:
            cat = _stat_symbol_to_category(v)
            if cat and cat in cnt:
                cnt[cat] += 1
    # 出差天 = 出差符号数/2
    chuchai = cnt.get("出差天", 0) / 2.0
    shijia  = cnt.get("事", 0) / 2.0
    queka   = cnt.get("缺", 0) / 2.0
    out = {}
    for c in STATS_CATEGORIES:
        if c == "出勤天":
            # 月应出勤 - 事假 - 缺卡（不能为负）
            v = max(0, month_workdays - shijia - queka)
            out[c] = round(v, 2)
        elif c == "出差天":
            out[c] = round(chuchai, 2)
        else:
            out[c] = round(cnt.get(c, 0) / 2.0, 2)
    return out

def _empty_stats():
    return {c: 0 for c in STATS_CATEGORIES}

def build_export(conn, region_filter=None):
    if not ensure_openpyxl():
        raise RuntimeError("后端未安装 openpyxl")
    meta = dict((r["k"], r["v"]) for r in conn.execute("SELECT k,v FROM meta"))
    day_numbers = json.loads(meta.get("day_numbers", "[]"))
    summary_names = json.loads(meta.get("summary_names", "[]"))
    year = int(meta.get("year", 2026))
    month = int(meta.get("month", 8))
    legend_raw = meta.get("legend_raw", "")
    ndays = len(day_numbers)
    c = conn.cursor()
    if region_filter:
        emps = c.execute("SELECT * FROM employees WHERE dept=? ORDER BY id",
                         (region_filter,)).fetchall()
    else:
        emps = c.execute("SELECT * FROM employees ORDER BY id").fetchall()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "总表"
    title = f"{year}年{month}月份考勤表"
    ws.cell(row=1, column=1, value=title)
    # 星期行
    for di, d in enumerate(day_numbers):
        ws.cell(row=1, column=5 + di, value=weekday_letter(year, month, d))
    # 图例
    ws.cell(row=2, column=5, value=legend_raw)
    # 表头
    header = ["序号", "部门", "姓名", "时间"] + day_numbers + summary_names
    for ci, h in enumerate(header, start=1):
        ws.cell(row=4, column=ci, value=h)
    # 数据
    rownum = 5
    for e in emps:
        am = json.loads(e["am"] or "[]")
        pm = json.loads(e["pm"] or "[]")
        summary = json.loads(e["summary"] or "{}")
        # 上午行
        ws.cell(row=rownum, column=1, value=e["seq"])
        ws.cell(row=rownum, column=2, value=e["dept"])
        ws.cell(row=rownum, column=3, value=e["name"])
        ws.cell(row=rownum, column=4, value="上午")
        for di in range(ndays):
            ws.cell(row=rownum, column=5 + di, value=am[di] if di < len(am) else None)
        for si, sn in enumerate(summary_names):
            ws.cell(row=rownum, column=5 + ndays + si, value=summary.get(sn))
        ws.cell(row=rownum, column=5 + ndays + len(summary_names), value=e["company"])
        rownum += 1
        # 下午行
        ws.cell(row=rownum, column=3, value=e["name"])
        ws.cell(row=rownum, column=4, value="下午")
        for di in range(ndays):
            ws.cell(row=rownum, column=5 + di, value=pm[di] if di < len(pm) else None)
        rownum += 1
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ----------------------------------------------------------------------------
# HTTP 处理器
# ----------------------------------------------------------------------------
class Handler(http.server.BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _auth_user(self):
        return current_user(self)

    # ---- 统一异常兜底：任何接口出错都返回 JSON 500，而不是直接断连 ----
    def do_GET(self):
        try:
            self._do_GET()
        except Exception as e:
            traceback.print_exc()
            try:
                send_json(self, {"error": "服务器内部错误: " + str(e)}, 500)
            except Exception:
                pass

    def do_POST(self):
        try:
            self._do_POST()
        except Exception as e:
            traceback.print_exc()
            try:
                send_json(self, {"error": "服务器内部错误: " + str(e)}, 500)
            except Exception:
                pass

    def _do_GET(self):
        p = urllib.parse.urlparse(self.path).path
        if p == "/" or p == "/index.html":
            self.serve_index()
            return
        if p.startswith("/uploads/"):
            self.serve_evidence(p[len("/uploads/"):])
            return
        if p == "/api/me":
            u = self._auth_user()
            if not u:
                return send_json(self, {"error": "未登录"}, 401)
            conn = get_db(); cur = conn.cursor()
            row = cur.execute("SELECT role,region FROM users WHERE username=?",
                              (u,)).fetchone()
            conn.close()
            if not row:
                return send_json(self, {"error": "无此用户"}, 401)
            send_json(self, {"username": u, "role": row["role"], "region": row["region"]})
            return
        if p == "/api/meta":
            conn = get_db(); cur = conn.cursor()
            rows = cur.execute("SELECT k,v FROM meta").fetchall(); conn.close()
            meta = {r["k"]: r["v"] for r in rows}
            if meta:
                meta = {**meta,
                        "day_numbers": json.loads(meta.get("day_numbers", "[]")),
                        "summary_names": json.loads(meta.get("summary_names", "[]")),
                        "legend": json.loads(meta.get("legend", "[]"))}
            # 月应出勤天数（v2026-09-04-stats）；未设置时默认 26
            try:
                meta["month_workdays"] = int(meta.get("month_workdays") or 26)
            except Exception:
                meta["month_workdays"] = 26
            meta["stats_categories"] = STATS_CATEGORIES
            send_json(self, meta)
            return
        if p == "/api/departments":
            u = self._auth_user()
            if not u or get_role(u) != "admin":
                return send_json(self, {"error": "需要管理员权限"}, 403)
            conn = get_db(); cur = conn.cursor()
            emps = cur.execute("SELECT dept, COUNT(*) c FROM employees GROUP BY dept").fetchall()
            users = cur.execute("SELECT username,region FROM users WHERE role='supervisor'").fetchall()
            acc = {r["region"]: r["username"] for r in users}
            depts = [{"dept": e["dept"], "account": acc.get(e["dept"]),
                      "emp_count": e["c"]} for e in emps]
            conn.close()
            send_json(self, {"departments": depts})
            return
        if p == "/api/admin/employees":
            u = self._auth_user()
            if not u or get_role(u) != "admin":
                return send_json(self, {"error": "需要管理员权限"}, 403)
            conn = get_db(); cur = conn.cursor()
            emps = cur.execute("SELECT * FROM employees ORDER BY dept,id").fetchall()
            out = []
            for e in emps:
                out.append({"id": e["id"], "dept": e["dept"], "name": e["name"],
                            "am": json.loads(e["am"] or "[]"),
                            "pm": json.loads(e["pm"] or "[]"),
                            "summary": json.loads(e["summary"] or "{}")})
            conn.close()
            send_json(self, {"employees": out})
            return
        if p == "/api/submissions":
            u = self._auth_user()
            if not u or get_role(u) != "admin":
                return send_json(self, {"error": "需要管理员权限"}, 403)
            conn = get_db(); cur = conn.cursor()
            rows = cur.execute("""
                SELECT s.*, e.name, e.dept FROM submissions s
                JOIN employees e ON e.id=s.emp_id
                ORDER BY s.updated_at DESC
            """).fetchall()
            out = []
            for s in rows:
                evs = cur.execute("SELECT id,filename,cell_key FROM evidence WHERE submission_id=?",
                                  (s["id"],)).fetchall()
                out.append({
                    "id": s["id"], "emp_id": s["emp_id"], "name": s["name"],
                    "dept": s["dept"], "supervisor": s["supervisor"],
                    "note": s["note"], "status": s["status"],
                    "cells": json.loads(s["cells"] or "{}"),
                    "reviewed_cells": json.loads(s["reviewed_cells"] or "{}"),
                    "change_requests": json.loads(s["change_requests"] or "{}"),
                    "admin_note": s["admin_note"],
                    "evidence": [{"id": ev["id"], "filename": ev["filename"], "cell_key": ev["cell_key"]} for ev in evs],
                })
            conn.close()
            send_json(self, {"submissions": out})
            return
        if p == "/api/admin/change_requests":
            u = self._auth_user()
            if not u or get_role(u) != "admin":
                return send_json(self, {"error": "需要管理员权限"}, 403)
            q = urllib.parse.urlparse(self.path).query
            qs = urllib.parse.parse_qs(q)
            status_filter = (qs.get("status") or ["all"])[0]
            conn = get_db(); cur = conn.cursor()
            rows = cur.execute("""
                SELECT s.*, e.name, e.dept FROM submissions s
                JOIN employees e ON e.id=s.emp_id
                WHERE s.change_requests IS NOT NULL AND s.change_requests != ''
                ORDER BY s.updated_at DESC
            """).fetchall()
            out = []
            for s in rows:
                crs = json.loads(s["change_requests"] or "{}")
                if not crs: continue
                # 收集每个 request 对应的凭证
                rid_to_evs = {}
                evs = cur.execute(
                    "SELECT id,filename,cell_key,request_id FROM evidence "
                    "WHERE submission_id=? AND request_id IS NOT NULL AND request_id != ''",
                    (s["id"],)).fetchall()
                for ev in evs:
                    rid_to_evs.setdefault(ev["request_id"], []).append(
                        {"id": ev["id"], "filename": ev["filename"], "cell_key": ev["cell_key"]})
                # 把每个 request 展开成数组元素
                for rid in sorted(crs.keys()):
                    cr = crs[rid]
                    if status_filter != "all" and cr.get("status") != status_filter:
                        continue
                    out.append({
                        "submission_id": s["id"], "emp_id": s["emp_id"],
                        "name": s["name"], "dept": s["dept"], "supervisor": s["supervisor"],
                        "request_id": rid,
                        "cell_key": cr["cell_key"],
                        "old": cr["old"], "new": cr["new"],
                        "reason": cr.get("reason", ""),
                        "status": cr.get("status", "pending"),
                        "reject_reason": cr.get("reject_reason", ""),
                        "created_at": cr.get("created_at", ""),
                        "reviewed_at": cr.get("reviewed_at", ""),
                        "reviewer": cr.get("reviewer", ""),
                        "evidence": rid_to_evs.get(rid, []),
                    })
            conn.close()
            send_json(self, {"requests": out})
            return
        if p == "/api/admin/stats" or p.startswith("/api/admin/stats?"):
            # 考勤统计：按部门汇总 + 全公司汇总 + 每员工明细
            u = self._auth_user()
            if not u or get_role(u) != "admin":
                return send_json(self, {"error": "需要管理员权限"}, 403)
            qs = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
            try:
                month_workdays = int((qs.get("month_workdays") or [26])[0])
            except Exception:
                month_workdays = 26
            dept_filter = (qs.get("dept") or [None])[0]
            conn = get_db(); cur = conn.cursor()
            rows = cur.execute("SELECT id,dept,name,am,pm,company FROM employees ORDER BY dept,id").fetchall()
            conn.close()
            grand = _empty_stats()
            dept_map = {}   # dept -> {"employees":[...], "totals":{...}}
            for r in rows:
                if dept_filter and r["dept"] != dept_filter:
                    continue
                am = json.loads(r["am"] or "[]")
                pm = json.loads(r["pm"] or "[]")
                stats = compute_employee_stats(am, pm, month_workdays)
                bucket = dept_map.setdefault(r["dept"], {"employees": [], "totals": _empty_stats()})
                bucket["employees"].append({
                    "id": r["id"], "name": r["name"], "company": r["company"] or "",
                    "stats": stats,
                })
                for k, v in stats.items():
                    bucket["totals"][k] += v
                    grand[k] += v
            # 排序部门
            out_by_dept = []
            for dept in sorted(dept_map.keys()):
                bucket = dept_map[dept]
                # 排序员工按 id
                bucket["employees"].sort(key=lambda e: e["id"])
                # 部门 totals 保留 2 位小数
                bucket["totals"] = {k: round(v, 2) for k, v in bucket["totals"].items()}
                out_by_dept.append({"dept": dept, **bucket})
            send_json(self, {
                "month_workdays": month_workdays,
                "categories": STATS_CATEGORIES,
                "by_dept": out_by_dept,
                "grand_totals": {k: round(v, 2) for k, v in grand.items()},
                "total_employees": sum(len(b["employees"]) for b in out_by_dept),
                "total_departments": len(out_by_dept),
            })
            return
        if p == "/api/admin/stats/export" or p.startswith("/api/admin/stats/export?"):
            # 导出统计 Excel：15 列 = 14 stats + 公司
            u = self._auth_user()
            if not u or get_role(u) != "admin":
                return send_json(self, {"error": "需要管理员权限"}, 403)
            if not ensure_openpyxl():
                return send_json(self, {"error": "未安装 openpyxl"}, 500)
            qs = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
            try:
                month_workdays = int((qs.get("month_workdays") or [26])[0])
            except Exception:
                month_workdays = 26
            conn = get_db(); cur = conn.cursor()
            rows = cur.execute("SELECT id,dept,name,am,pm,company FROM employees ORDER BY dept,id").fetchall()
            meta = dict((r["k"], r["v"]) for r in cur.execute("SELECT k,v FROM meta"))
            conn.close()
            wb = openpyxl.Workbook(); ws = wb.active; ws.title = "考勤统计"
            # 标题：14 stats + 公司（与前端截图一致）
            for i, c in enumerate(STATS_CATEGORIES):
                ws.cell(1, 1 + i, c)
            ws.cell(1, 1 + len(STATS_CATEGORIES), "公司")
            r = 2
            for row in rows:
                am = json.loads(row["am"] or "[]"); pm = json.loads(row["pm"] or "[]")
                stats = compute_employee_stats(am, pm, month_workdays)
                for i, c in enumerate(STATS_CATEGORIES):
                    ws.cell(r, 1 + i, stats[c])
                ws.cell(r, 1 + len(STATS_CATEGORIES), row["company"] or "")
                r += 1
            # 列宽
            for i in range(len(STATS_CATEGORIES)):
                col = chr(ord("A") + i) if i < 26 else None
                if col: ws.column_dimensions[col].width = 11
            col = chr(ord("A") + len(STATS_CATEGORIES)) if len(STATS_CATEGORIES) < 26 else None
            if col: ws.column_dimensions[col].width = 14
            from io import BytesIO
            bio = BytesIO(); wb.save(bio); body = bio.getvalue()
            self.send_response(200)
            self.send_header("Content-Type",
                             "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Length", str(len(body)))
            fname = "考勤统计.xlsx"
            self.send_header("Content-Disposition",
                "attachment; filename=\"%s\"; filename*=UTF-8''%s"
                % (urllib.parse.quote(fname), urllib.parse.quote(fname)))
            self.end_headers(); self.wfile.write(body)
            return
        if p == "/api/my/employees":
            u = self._auth_user()
            if not u or get_role(u) != "supervisor":
                return send_json(self, {"error": "需要主管权限"}, 403)
            region = get_role_region(u)
            conn = get_db(); cur = conn.cursor()
            emps = cur.execute("SELECT * FROM employees WHERE dept=? ORDER BY id",
                               (region,)).fetchall()
            subs = cur.execute("SELECT emp_id,status,id,cells,reviewed_cells,change_requests FROM submissions").fetchall()
            st = {s["emp_id"]: {"status": s["status"], "sub_id": s["id"],
                                "cells": json.loads(s["cells"] or "{}"),
                                "reviewed_cells": json.loads(s["reviewed_cells"] or "{}"),
                                "change_requests": json.loads(s["change_requests"] or "{}")} for s in subs}
            out = []
            for e in emps:
                info = st.get(e["id"], {})
                # 统计修改申请摘要
                crs = info.get("change_requests", {})
                req_summary = {"pending": 0, "approved": 0, "rejected": 0}
                for rid, cr in crs.items():
                    stt = cr.get("status", "pending")
                    if stt in req_summary: req_summary[stt] += 1
                out.append({"id": e["id"], "name": e["name"],
                            "am": json.loads(e["am"] or "[]"),
                            "pm": json.loads(e["pm"] or "[]"),
                            "status": info.get("status", "none"),
                            "sub_id": info.get("sub_id"),
                            "cells": info.get("cells", {}),
                            "reviewed_cells": info.get("reviewed_cells", {}),
                            "change_requests": crs,
                            "req_summary": req_summary})
            conn.close()
            send_json(self, {"employees": out, "region": region})
            return
        if p == "/api/my/submission":
            u = self._auth_user()
            if not u or get_role(u) != "supervisor":
                return send_json(self, {"error": "需要主管权限"}, 403)
            q = urllib.parse.urlparse(self.path).query
            emp_id = urllib.parse.parse_qs(q).get("emp_id", [None])[0]
            conn = get_db(); cur = conn.cursor()
            s = cur.execute("SELECT * FROM submissions WHERE emp_id=? ORDER BY id DESC LIMIT 1",
                            (emp_id,)).fetchone()
            if not s:
                conn.close()
                return send_json(self, {"submission": None})
            evs = cur.execute("SELECT id,filename,cell_key,request_id FROM evidence WHERE submission_id=?",
                              (s["id"],)).fetchall()
            conn.close()
            send_json(self, {"submission": {
                "id": s["id"], "note": s["note"], "status": s["status"],
                "admin_note": s["admin_note"],
                "cells": json.loads(s["cells"] or "{}"),
                "reviewed_cells": json.loads(s["reviewed_cells"] or "{}"),
                "change_requests": json.loads(s["change_requests"] or "{}"),
                "evidence": [{"id": ev["id"], "filename": ev["filename"], "cell_key": ev["cell_key"], "request_id": ev["request_id"]} for ev in evs],
            }})
            return
        if p == "/api/my/change_requests":
            u = self._auth_user()
            if not u or get_role(u) != "supervisor":
                return send_json(self, {"error": "需要主管权限"}, 403)
            q = urllib.parse.urlparse(self.path).query
            qs = urllib.parse.parse_qs(q)
            emp_id = (qs.get("emp_id") or [None])[0]
            region = get_role_region(u)
            conn = get_db(); cur = conn.cursor()
            if emp_id:
                emp = cur.execute("SELECT id,dept FROM employees WHERE id=?", (emp_id,)).fetchone()
                if not emp or emp["dept"] != region:
                    conn.close()
                    return send_json(self, {"error": "无权操作该员工"}, 403)
                s = cur.execute("SELECT * FROM submissions WHERE emp_id=? ORDER BY id DESC LIMIT 1",
                                (emp_id,)).fetchone()
            else:
                s = None
            out = []
            if s:
                crs = json.loads(s["change_requests"] or "{}")
                evs = cur.execute(
                    "SELECT id,filename,cell_key,request_id FROM evidence "
                    "WHERE submission_id=? AND request_id IS NOT NULL AND request_id != ''",
                    (s["id"],)).fetchall()
                rid_to_evs = {}
                for ev in evs:
                    rid_to_evs.setdefault(ev["request_id"], []).append(
                        {"id": ev["id"], "filename": ev["filename"], "cell_key": ev["cell_key"]})
                for rid in sorted(crs.keys()):
                    cr = crs[rid]
                    out.append({
                        "submission_id": s["id"], "emp_id": emp_id,
                        "request_id": rid,
                        "cell_key": cr["cell_key"],
                        "old": cr["old"], "new": cr["new"],
                        "reason": cr.get("reason", ""),
                        "status": cr.get("status", "pending"),
                        "reject_reason": cr.get("reject_reason", ""),
                        "created_at": cr.get("created_at", ""),
                        "reviewed_at": cr.get("reviewed_at", ""),
                        "reviewer": cr.get("reviewer", ""),
                        "evidence": rid_to_evs.get(rid, []),
                    })
            conn.close()
            send_json(self, {"requests": out})
            return
        if p == "/api/export":
            u = self._auth_user()
            if not u:
                return send_json(self, {"error": "未登录"}, 401)
            role = get_role(u)
            conn = get_db()
            if role == "admin":
                data = build_export(conn)
                fname = "考勤总表_更新.xlsx"
            elif role == "supervisor":
                region = get_role_region(u)
                data = build_export(conn, region_filter=region)
                fname = f"考勤总表_{region}_更新.xlsx"
            else:
                return send_json(self, {"error": "无权限"}, 403)
            conn.close()
            body = data
            self.send_response(200)
            self.send_header("Content-Type",
                             "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Length", str(len(body)))
            from email.utils import encode_rfc2231
            self.send_header("Content-Disposition",
                             "attachment; filename=\"%s\"; filename*=UTF-8''%s" %
                             (urllib.parse.quote(fname), urllib.parse.quote(fname)))
            self.end_headers()
            self.wfile.write(body)
            return
        return send_json(self, {"error": "not found"}, 404)

    def _do_POST(self):
        p = urllib.parse.urlparse(self.path).path
        if p == "/api/login":
            d = read_json_body(self)
            conn = get_db(); cur = conn.cursor()
            row = cur.execute("SELECT pw,role,region FROM users WHERE username=?",
                              (d.get("username"),)).fetchone()
            conn.close()
            if not row or row["pw"] != pw_hash(d.get("password", "")):
                return send_json(self, {"error": "账号或密码错误"}, 401)
            tok = secrets.token_hex(16)
            TOKENS[tok] = d.get("username")
            return send_json(self, {"token": tok, "role": row["role"], "region": row["region"]})
        if p == "/api/upload":
            u = self._auth_user()
            if not u or get_role(u) != "admin":
                return send_json(self, {"error": "需要管理员权限"}, 403)
            d = read_json_body(self)
            try:
                raw = b64_to_bytes(d.get("data", ""))
                parsed = parse_xlsx(raw)
            except Exception as ex:
                return send_json(self, {"error": "解析失败：" + str(ex)}, 400)
            conn = get_db(); cur = conn.cursor()
            cur.execute("DELETE FROM employees")
            cur.execute("DELETE FROM submissions")
            cur.execute("DELETE FROM evidence")
            cur.execute("DELETE FROM meta")
            for e in parsed["employees"]:
                cur.execute(
                    "INSERT INTO employees (seq,dept,name,am,pm,summary,company) VALUES (?,?,?,?,?,?,?)",
                    (e["seq"], e["dept"], e["name"],
                     json.dumps(e["am"], ensure_ascii=False),
                     json.dumps(e["pm"], ensure_ascii=False),
                     json.dumps(e["summary"], ensure_ascii=False),
                     e["company"]))
            cur.execute("INSERT INTO meta (k,v) VALUES (?,?)",
                        ("day_numbers", json.dumps(parsed["day_numbers"], ensure_ascii=False)))
            cur.execute("INSERT INTO meta (k,v) VALUES (?,?)",
                        ("summary_names", json.dumps(parsed["summary_names"], ensure_ascii=False)))
            cur.execute("INSERT INTO meta (k,v) VALUES (?,?)",
                        ("legend_raw", parsed["legend_raw"]))
            cur.execute("INSERT INTO meta (k,v) VALUES (?,?)",
                        ("legend", json.dumps(parsed["legend"], ensure_ascii=False)))
            cur.execute("INSERT INTO meta (k,v) VALUES (?,?)",
                        ("year", str(parsed["year"])))
            cur.execute("INSERT INTO meta (k,v) VALUES (?,?)",
                        ("month", str(parsed["month"])))
            conn.commit()
            depts = [e["dept"] for e in parsed["employees"]]
            conn.close()
            return send_json(self, {"ok": True, "records": len(parsed["employees"]),
                                    "departments": len(set(depts))})
        if p == "/api/users":
            u = self._auth_user()
            if not u or get_role(u) != "admin":
                return send_json(self, {"error": "需要管理员权限"}, 403)
            d = read_json_body(self)
            region = d.get("region"); uname = d.get("username"); pw = d.get("password", "123456")
            if not region or not uname:
                return send_json(self, {"error": "地区和用户名必填"}, 400)
            conn = get_db(); cur = conn.cursor()
            cur.execute("DELETE FROM users WHERE username=?", (uname,))
            cur.execute("INSERT INTO users (username,pw,role,region) VALUES (?,?,?,?)",
                        (uname, pw_hash(pw), "supervisor", region))
            conn.commit(); conn.close()
            return send_json(self, {"ok": True})
        if p == "/api/admin/month_workdays":
            # 设置"月应出勤天数"（用于出勤天数 = 月应出勤 - 事假/2 - 缺卡/2）
            u = self._auth_user()
            if not u or get_role(u) != "admin":
                return send_json(self, {"error": "需要管理员权限"}, 403)
            d = read_json_body(self)
            try:
                mw = int(d.get("month_workdays"))
            except Exception:
                return send_json(self, {"error": "month_workdays 必须是整数"}, 400)
            if mw < 0 or mw > 31:
                return send_json(self, {"error": "month_workdays 必须在 0~31 之间"}, 400)
            conn = get_db(); cur = conn.cursor()
            cur.execute("INSERT INTO meta(k,v) VALUES('month_workdays',?) "
                        "ON CONFLICT(k) DO UPDATE SET v=excluded.v", (str(mw),))
            conn.commit(); conn.close()
            return send_json(self, {"ok": True, "month_workdays": mw})
        if p == "/api/submission/review":
            u = self._auth_user()
            if not u or get_role(u) != "admin":
                return send_json(self, {"error": "需要管理员权限"}, 403)
            d = read_json_body(self)
            conn = get_db(); cur = conn.cursor()
            ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            action = d.get("action") or "status_only"
            sub_id = d.get("id")
            # 取出 submission 与 employees
            row = cur.execute("SELECT * FROM submissions WHERE id=?",
                              (sub_id,)).fetchone()
            if not row:
                conn.close()
                return send_json(self, {"error": "无此提交"}, 404)
            emp = cur.execute("SELECT am,pm FROM employees WHERE id=?",
                              (row["emp_id"],)).fetchone()
            am_arr = json.loads(emp["am"] or "[]")
            pm_arr = json.loads(emp["pm"] or "[]")
            cells = json.loads(row["cells"] or "{}")
            reviewed = json.loads(row["reviewed_cells"] or "{}")
            if action == "cell_review":
                # 管理员审批单个格子：修改考勤数据 + 标记该格子 reviewed=green
                cell_key = d.get("cell_key")    # "am_1" / "pm_5"
                symbol = d.get("symbol", "")
                if not cell_key or "_" not in cell_key:
                    conn.close()
                    return send_json(self, {"error": "cell_key 必填 (格式: am_1)"}, 400)
                period, day_s = cell_key.split("_", 1)
                try:
                    day = int(day_s) - 1
                except Exception:
                    conn.close()
                    return send_json(self, {"error": "day 格式错误"}, 400)
                arr = am_arr if period == "am" else pm_arr
                while len(arr) <= day:
                    arr.append(None)
                arr[day] = symbol if symbol != "" else None
                # 写回 employees
                field = "am" if period == "am" else "pm"
                cur.execute("UPDATE employees SET {}=? WHERE id=?".format(field),
                            (json.dumps(arr, ensure_ascii=False), row["emp_id"]))
                # 标记 reviewed_cells
                reviewed[cell_key] = "green"
                # 检查是否还有未审批的 red
                pending = [k for k, v in cells.items() if v == "red" and reviewed.get(k) != "green"]
                new_status = "approved" if not pending else "submitted"
                cur.execute("UPDATE submissions SET reviewed_cells=?, status=?, updated_at=? WHERE id=?",
                            (json.dumps(reviewed, ensure_ascii=False),
                             new_status, ts, sub_id))
            elif action == "cell_revert":
                # 撤销某个格子的审批（标记为未审批）
                cell_key = d.get("cell_key")
                if reviewed.pop(cell_key, None) is None:
                    conn.close()
                    return send_json(self, {"error": "该格子未被审批"}, 400)
                pending = [k for k, v in cells.items() if v == "red" and reviewed.get(k) != "green"]
                new_status = "submitted" if pending else "submitted"
                cur.execute("UPDATE submissions SET reviewed_cells=?, status=?, updated_at=? WHERE id=?",
                            (json.dumps(reviewed, ensure_ascii=False),
                             new_status, ts, sub_id))
            else:
                # 整单备注/状态（兼容老接口）
                cur.execute("UPDATE submissions SET status=?, admin_note=?, updated_at=? WHERE id=?",
                            (d.get("status"), d.get("admin_note", ""), ts, sub_id))
            conn.commit(); conn.close()
            return send_json(self, {"ok": True, "reviewed_cells": reviewed, "cells": cells})
        if p == "/api/employee/edit":
            u = self._auth_user()
            if not u or get_role(u) != "admin":
                return send_json(self, {"error": "需要管理员权限"}, 403)
            d = read_json_body(self)
            conn = get_db(); cur = conn.cursor()
            e = cur.execute("SELECT am,pm FROM employees WHERE id=?", (d.get("emp_id"),)).fetchone()
            if not e:
                conn.close()
                return send_json(self, {"error": "无此员工"}, 404)
            field = "am" if d.get("period") == "am" else "pm"
            arr = json.loads(e[field] or "[]")
            day = int(d.get("day", 1)) - 1
            while len(arr) <= day:
                arr.append(None)
            arr[day] = d.get("symbol")
            cur.execute("UPDATE employees SET %s=? WHERE id=?" % field,
                        (json.dumps(arr, ensure_ascii=False), d.get("emp_id")))
            # 同步 reviewed_cells
            sub = cur.execute("SELECT id,reviewed_cells FROM submissions WHERE emp_id=?",
                              (d.get("emp_id"),)).fetchone()
            if sub:
                reviewed = json.loads(sub["reviewed_cells"] or "{}")
                k = "{}_{}".format(field, day+1)
                if reviewed.get(k) != "green":
                    reviewed[k] = "green"
                    cur.execute("UPDATE submissions SET reviewed_cells=?, updated_at=? WHERE id=?",
                                (json.dumps(reviewed, ensure_ascii=False),
                                 datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                 sub["id"]))
            conn.commit(); conn.close()
            return send_json(self, {"ok": True})
        if p == "/api/employee/batch":
            u = self._auth_user()
            if not u or get_role(u) != "admin":
                return send_json(self, {"error": "需要管理员权限"}, 403)
            d = read_json_body(self)
            emp_id = d.get("emp_id")
            period = d.get("period")
            updates = d.get("updates") or []
            if period not in ("am","pm"):
                return send_json(self, {"error": "period 必须是 am/pm"}, 400)
            if not isinstance(updates, list):
                return send_json(self, {"error": "updates 必须是数组"}, 400)
            conn = get_db(); cur = conn.cursor()
            e = cur.execute("SELECT am,pm FROM employees WHERE id=?", (emp_id,)).fetchone()
            if not e:
                conn.close()
                return send_json(self, {"error": "无此员工"}, 404)
            field = "am" if period == "am" else "pm"
            arr = json.loads(e[field] or "[]")
            saved = 0
            changed_days = []
            for u_item in updates:
                try:
                    day = int(u_item.get("day", 1)) - 1
                except Exception:
                    continue
                sym = u_item.get("symbol")
                if sym is not None and sym != "":
                    sym = str(sym)[:4]  # 防止超长
                while len(arr) <= day:
                    arr.append(None)
                arr[day] = sym
                saved += 1
                changed_days.append(day)
            cur.execute("UPDATE employees SET %s=? WHERE id=?" % field,
                        (json.dumps(arr, ensure_ascii=False), emp_id))
            # 同步标记 reviewed_cells（如该员工有 submission）
            sub = cur.execute("SELECT id,cells,reviewed_cells FROM submissions WHERE emp_id=?",
                              (emp_id,)).fetchone()
            if sub and changed_days:
                reviewed = json.loads(sub["reviewed_cells"] or "{}")
                ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                for d in changed_days:
                    k = "{}_{}".format(period, d+1)
                    if reviewed.get(k) == "green":
                        continue
                    reviewed[k] = "green"
                cur.execute("UPDATE submissions SET reviewed_cells=?, updated_at=? WHERE id=?",
                            (json.dumps(reviewed, ensure_ascii=False), ts, sub["id"]))
            conn.commit(); conn.close()
            return send_json(self, {"ok": True, "saved": saved})
        if p == "/api/my/submission":
            u = self._auth_user()
            if not u or get_role(u) != "supervisor":
                return send_json(self, {"error": "需要主管权限"}, 403)
            d = read_json_body(self)
            region = get_role_region(u)
            conn = get_db(); cur = conn.cursor()
            emp = cur.execute("SELECT id,dept,am,pm FROM employees WHERE id=?",
                              (d.get("emp_id"),)).fetchone()
            if not emp or emp["dept"] != region:
                return send_json(self, {"error": "无权操作该员工"}, 403)
            ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            mode = d.get("mode", "all_ok")   # all_ok=整单无误；partial=部分有问题
            # 计算所有有数据的格子位置 {am_1, am_5, pm_3...}
            def all_positions(am_arr, pm_arr):
                pos = []
                for i, v in enumerate(am_arr):
                    if v not in (None, ""): pos.append(("am", i+1))
                for i, v in enumerate(pm_arr):
                    if v not in (None, ""): pos.append(("pm", i+1))
                return pos
            am_arr = json.loads(emp["am"] or "[]")
            pm_arr = json.loads(emp["pm"] or "[]")
            all_pos = all_positions(am_arr, pm_arr)
            cells = {}
            problem_cells = d.get("problem_cells", []) or []
            problem_set = set()
            for pc in problem_cells:
                p = pc.get("period")
                day = pc.get("day")
                if p not in ("am","pm") or not isinstance(day,int):
                    continue
                problem_set.add((p, day))
                cells["{}_{}".format(p, day)] = "red"
            # 其他有数据的格子标 green
            for p, day in all_pos:
                k = "{}_{}".format(p, day)
                if k not in cells:
                    cells[k] = "green"
            # 处理"撤销问题"的格子（remove_cells）：从 cells 中移除 red，强制设 green
            remove_cells = d.get("remove_cells", []) or []
            for rc in remove_cells:
                p = rc.get("period")
                day = rc.get("day")
                if p in ("am","pm") and isinstance(day,int):
                    k = "{}_{}".format(p, day)
                    cells[k] = "green"  # 撤销后强制 green
            # 找/创建 submission
            existing = cur.execute("SELECT id,cells FROM submissions WHERE emp_id=?",
                                   (d.get("emp_id"),)).fetchone()
            if existing:
                # 合并：旧的 red 格子中，本次未明确移除的才保留
                old_cells = json.loads(existing["cells"] or "{}")
                for k, v in old_cells.items():
                    if v == "red" and k in problem_set:
                        cells[k] = "red"
                # 删除 remove_cells 中的 red 凭证
                if remove_cells:
                    for rc in remove_cells:
                        p = rc.get("period")
                        day = rc.get("day")
                        if p in ("am","pm") and isinstance(day,int):
                            k = "{}_{}".format(p, day)
                            cur.execute("DELETE FROM evidence WHERE submission_id=? AND cell_key=?",
                                        (existing["id"], k))
                cur.execute("UPDATE submissions SET note=?, status='submitted', cells=?, updated_at=? WHERE id=?",
                            (d.get("note", ""), json.dumps(cells, ensure_ascii=False), ts, existing["id"]))
                sid = existing["id"]
            else:
                cur.execute("INSERT INTO submissions (emp_id,supervisor,note,status,cells,created_at,updated_at) "
                            "VALUES (?,?,?,?,?,?,?)",
                            (d.get("emp_id"), u, d.get("note", ""), "submitted",
                             json.dumps(cells, ensure_ascii=False), ts, ts))
                sid = cur.lastrowid
            # 写入 problem_cells 的凭证
            for pc in problem_cells:
                if pc.get("period") not in ("am","pm") or not isinstance(pc.get("day"), int):
                    continue
                raw = b64_to_bytes(pc.get("data", ""))
                if not raw:
                    continue
                fname = pc.get("filename", "evidence.bin")
                ext = os.path.splitext(fname)[1] or ".bin"
                store_name = "%d_%d_%s_%d%s" % (sid, int(datetime.datetime.now().timestamp()*1000),
                                                pc["period"], pc["day"], ext)
                cur.execute("INSERT INTO evidence (submission_id,filename,cell_key,data,created_at) "
                            "VALUES (?,?,?,?,?)",
                            (sid, fname, "{}_{}".format(pc["period"], pc["day"]),
                             raw, ts))
                with open(os.path.join(UPLOAD_DIR, store_name), "wb") as f:
                    f.write(raw)
            conn.commit(); conn.close()
            return send_json(self, {"ok": True, "submission_id": sid,
                                    "status": "submitted", "cells": cells,
                                    "red_count": len(problem_cells)})
        if p == "/api/my/change_request":
            u = self._auth_user()
            if not u or get_role(u) != "supervisor":
                return send_json(self, {"error": "需要主管权限"}, 403)
            d = read_json_body(self)
            emp_id = d.get("emp_id")
            period = d.get("period")
            day = d.get("day")
            new_symbol = (d.get("new_symbol") or "").strip()
            reason = (d.get("reason") or "").strip()
            file_data = d.get("data") or ""
            filename = d.get("filename") or "evidence.bin"
            if period not in ("am", "pm") or not isinstance(day, int) or not (1 <= day <= 31):
                return send_json(self, {"error": "参数错误：period/day 必填，day 范围 1-31"}, 400)
            if not file_data:
                return send_json(self, {"error": "必须上传凭证照片"}, 400)
            if not new_symbol:
                return send_json(self, {"error": "新符号必填"}, 400)
            region = get_role_region(u)
            conn = get_db(); cur = conn.cursor()
            emp = cur.execute("SELECT dept,am,pm FROM employees WHERE id=?", (emp_id,)).fetchone()
            if not emp or emp["dept"] != region:
                conn.close()
                return send_json(self, {"error": "无权操作该员工"}, 403)
            old_arr = json.loads(emp["am"] if period == "am" else emp["pm"] or "[]")
            old_symbol = old_arr[day - 1] if 0 <= day - 1 < len(old_arr) else None
            if old_symbol == new_symbol:
                conn.close()
                return send_json(self, {"error": "新符号与原符号相同，无需申请"}, 400)
            # 找/创建 submission
            s = cur.execute("SELECT id,change_requests FROM submissions WHERE emp_id=?",
                            (emp_id,)).fetchone()
            ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            if not s:
                cur.execute("INSERT INTO submissions (emp_id,supervisor,note,status,created_at,updated_at) "
                            "VALUES (?,?,?,?,?,?)",
                            (emp_id, u, "（暂无说明）", "submitted", ts, ts))
                s_id = cur.lastrowid
                crs = {}
            else:
                s_id = s["id"]
                crs = json.loads(s["change_requests"] or "{}")
            cell_key = "%s_%d" % (period, day)
            rid = new_request_id(cell_key, crs)
            crs[rid] = {
                "cell_key": cell_key,
                "old": old_symbol,
                "new": new_symbol,
                "reason": reason,
                "status": "pending",
                "reject_reason": "",
                "created_at": ts,
                "reviewed_at": "",
                "reviewer": "",
            }
            cur.execute("UPDATE submissions SET change_requests=?, updated_at=? WHERE id=?",
                        (json.dumps(crs, ensure_ascii=False), ts, s_id))
            # 写入凭证
            try:
                raw = b64_to_bytes(file_data)
            except Exception as ex:
                conn.close()
                return send_json(self, {"error": "凭证数据解析失败：" + str(ex)}, 400)
            if not raw:
                conn.close()
                return send_json(self, {"error": "凭证为空"}, 400)
            ext = os.path.splitext(filename)[1] or ".png"
            store_name = "%d_%d_%s%s" % (s_id, int(datetime.datetime.now().timestamp() * 1000), rid, ext)
            cur.execute(
                "INSERT INTO evidence (submission_id,filename,cell_key,request_id,data,created_at) "
                "VALUES (?,?,?,?,?,?)",
                (s_id, filename, cell_key, rid, raw, ts))
            with open(os.path.join(UPLOAD_DIR, store_name), "wb") as f:
                f.write(raw)
            conn.commit(); conn.close()
            return send_json(self, {"ok": True, "request_id": rid,
                                    "submission_id": s_id, "status": "pending"})
        if p == "/api/my/change_request/cancel":
            u = self._auth_user()
            if not u or get_role(u) != "supervisor":
                return send_json(self, {"error": "需要主管权限"}, 403)
            d = read_json_body(self)
            emp_id = d.get("emp_id")
            req_id = d.get("request_id")
            region = get_role_region(u)
            conn = get_db(); cur = conn.cursor()
            s = cur.execute("SELECT * FROM submissions WHERE emp_id=?", (emp_id,)).fetchone()
            if not s:
                conn.close()
                return send_json(self, {"error": "无此提交"}, 404)
            emp = cur.execute("SELECT dept FROM employees WHERE id=?", (emp_id,)).fetchone()
            if not emp or emp["dept"] != region:
                conn.close()
                return send_json(self, {"error": "无权操作该员工"}, 403)
            crs = json.loads(s["change_requests"] or "{}")
            if req_id not in crs:
                conn.close()
                return send_json(self, {"error": "无此申请"}, 404)
            if crs[req_id].get("status") != "pending":
                conn.close()
                return send_json(self, {"error": "只能撤回待审批的申请"}, 400)
            crs.pop(req_id)
            ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            cur.execute("UPDATE submissions SET change_requests=?, updated_at=? WHERE id=?",
                        (json.dumps(crs, ensure_ascii=False), ts, s["id"]))
            # 删除该申请对应的凭证
            cur.execute("DELETE FROM evidence WHERE submission_id=? AND request_id=?",
                        (s["id"], req_id))
            conn.commit(); conn.close()
            return send_json(self, {"ok": True})
        if p == "/api/admin/change_request/review":
            u = self._auth_user()
            if not u or get_role(u) != "admin":
                return send_json(self, {"error": "需要管理员权限"}, 403)
            d = read_json_body(self)
            sub_id = d.get("submission_id")
            req_id = d.get("request_id")
            decision = d.get("decision")  # "approve" / "reject"
            reject_reason = (d.get("reject_reason") or "").strip()
            if not sub_id or not req_id or decision not in ("approve", "reject"):
                return send_json(self, {"error": "参数错误"}, 400)
            if decision == "reject" and not reject_reason:
                return send_json(self, {"error": "拒绝必须填写理由"}, 400)
            conn = get_db(); cur = conn.cursor()
            s = cur.execute("SELECT * FROM submissions WHERE id=?", (sub_id,)).fetchone()
            if not s:
                conn.close()
                return send_json(self, {"error": "无此提交"}, 404)
            crs = json.loads(s["change_requests"] or "{}")
            if req_id not in crs:
                conn.close()
                return send_json(self, {"error": "无此申请"}, 404)
            cr = crs[req_id]
            if cr.get("status") not in (None, "", "pending"):
                conn.close()
                return send_json(self, {"error": "该申请已被处理（" + cr.get("status") + "）"}, 400)
            ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            # 解析 cell_key
            try:
                period, day_s = cr["cell_key"].split("_", 1)
                day = int(day_s) - 1
            except Exception:
                conn.close()
                return send_json(self, {"error": "cell_key 格式错误"}, 400)
            if decision == "approve":
                # 改 employees 表对应格子符号
                emp = cur.execute("SELECT am,pm FROM employees WHERE id=?",
                                  (s["emp_id"],)).fetchone()
                am_arr = json.loads(emp["am"] or "[]")
                pm_arr = json.loads(emp["pm"] or "[]")
                arr = am_arr if period == "am" else pm_arr
                while len(arr) <= day:
                    arr.append(None)
                arr[day] = cr["new"] if cr["new"] != "" else None
                field = "am" if period == "am" else "pm"
                cur.execute("UPDATE employees SET {}=? WHERE id=?".format(field),
                            (json.dumps(arr, ensure_ascii=False), s["emp_id"]))
                # 也把 reviewed_cells 该格子标绿以同步主网格色
                reviewed = json.loads(s["reviewed_cells"] or "{}")
                reviewed[cr["cell_key"]] = "green"
                # 也清掉 cells 中的 red（如果存在）
                cells = json.loads(s["cells"] or "{}")
                if cells.get(cr["cell_key"]) == "red":
                    cells[cr["cell_key"]] = "green"
                cr["status"] = "approved"
                cr["reject_reason"] = ""
                cr["reviewed_at"] = ts
                cr["reviewer"] = u
                crs[req_id] = cr
                # 更新整个 submission：reviewed_cells + cells + change_requests
                pending_changes = [k for k, v in crs.items() if v.get("status") == "pending"]
                # 若还有未审批修改申请 + 未审批红格 → status 保持 submitted；否则 approved
                pending_red = [k for k, v in cells.items() if v == "red" and reviewed.get(k) != "green"]
                new_status = "submitted" if (pending_changes or pending_red) else "approved"
                cur.execute(
                    "UPDATE submissions SET change_requests=?, cells=?, reviewed_cells=?, status=?, updated_at=? WHERE id=?",
                    (json.dumps(crs, ensure_ascii=False),
                     json.dumps(cells, ensure_ascii=False),
                     json.dumps(reviewed, ensure_ascii=False),
                     new_status, ts, sub_id))
            else:
                # 拒绝：仅修改申请状态
                cr["status"] = "rejected"
                cr["reject_reason"] = reject_reason
                cr["reviewed_at"] = ts
                cr["reviewer"] = u
                crs[req_id] = cr
                # 若没有其他 pending 申请 + 没有未审批红格，submission 状态退回 approved
                pending_changes = [k for k, v in crs.items() if v.get("status") == "pending"]
                cells = json.loads(s["cells"] or "{}")
                reviewed = json.loads(s["reviewed_cells"] or "{}")
                pending_red = [k for k, v in cells.items() if v == "red" and reviewed.get(k) != "green"]
                new_status = "submitted" if (pending_changes or pending_red) else "approved"
                cur.execute(
                    "UPDATE submissions SET change_requests=?, status=?, updated_at=? WHERE id=?",
                    (json.dumps(crs, ensure_ascii=False), new_status, ts, sub_id))
            conn.commit(); conn.close()
            return send_json(self, {"ok": True, "status": cr["status"],
                                    "reject_reason": cr.get("reject_reason", "")})
        if p == "/api/my/evidence":
            u = self._auth_user()
            if not u or get_role(u) != "supervisor":
                return send_json(self, {"error": "需要主管权限"}, 403)
            d = read_json_body(self)
            sid = d.get("submission_id")
            conn = get_db(); cur = conn.cursor()
            sub = cur.execute("SELECT emp_id FROM submissions WHERE id=?", (sid,)).fetchone()
            if not sub:
                return send_json(self, {"error": "无此提交"}, 404)
            emp = cur.execute("SELECT dept FROM employees WHERE id=?", (sub["emp_id"],)).fetchone()
            if not emp or emp["dept"] != get_role_region(u):
                return send_json(self, {"error": "无权操作"}, 403)
            raw = b64_to_bytes(d.get("data", ""))
            fname = d.get("filename", "evidence.bin")
            ext = os.path.splitext(fname)[1] or ".bin"
            store_name = "%d_%d%s" % (sid, int(datetime.datetime.now().timestamp()*1000), ext)
            cur.execute("INSERT INTO evidence (submission_id,filename,cell_key,data,created_at) VALUES (?,?,?,?,?)",
                        (sid, fname, d.get("cell_key") or "", raw,
                         datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
            # 同时存磁盘便于查看
            with open(os.path.join(UPLOAD_DIR, store_name), "wb") as f:
                f.write(raw)
            conn.commit(); conn.close()
            return send_json(self, {"ok": True})
        return send_json(self, {"error": "not found"}, 404)

    def serve_index(self):
        try:
            with open(os.path.join(BASE_DIR, "index.html"), "r", encoding="utf-8") as f:
                send_html(self, f.read())
        except Exception:
            send_html(self, "<h1>index.html not found</h1>")

    def serve_evidence(self, fname):
        # 鉴权：支持 header 或 query string ?t=<token>（img 标签无法加 header）
        tok = None
        auth = self.headers.get("Authorization", "")
        if auth.startswith("Bearer "):
            tok = auth[7:].strip()
        if not tok:
            qs = urllib.parse.urlparse(self.path).query
            q = urllib.parse.parse_qs(qs)
            tok = (q.get("t") or [None])[0]
        if not tok or tok not in TOKENS:
            return send_json(self, {"error": "未登录"}, 401)
        u = TOKENS[tok]
        conn = get_db(); cur = conn.cursor()
        ev = cur.execute("SELECT submission_id,filename,data FROM evidence WHERE filename=?",
                        (fname,)).fetchone()
        if not ev:
            conn.close()
            return send_json(self, {"error": "无此凭证"}, 404)
        # 权限：管理员 或 该提交所属主管
        role = get_role(u)
        sub = cur.execute("SELECT emp_id FROM submissions WHERE id=?",
                          (ev["submission_id"],)).fetchone()
        allow = False
        if role == "admin":
            allow = True
        elif role == "supervisor":
            if sub:
                emp = cur.execute("SELECT dept FROM employees WHERE id=?", (sub["emp_id"],)).fetchone()
                if emp and emp["dept"] == get_role_region(u):
                    allow = True
        conn.close()
        if not allow:
            return send_json(self, {"error": "无权限"}, 403)
        data = ev["data"]
        ext = os.path.splitext(ev["filename"])[1].lower()
        ctype = "image/jpeg" if ext in (".jpg", ".jpeg") else \
                "image/png" if ext == ".png" else \
                "image/gif" if ext == ".gif" else "application/octet-stream"
        self.send_response(200)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)

# 角色辅助
def get_role(username):
    conn = get_db(); cur = conn.cursor()
    r = cur.execute("SELECT role FROM users WHERE username=?", (username,)).fetchone()
    conn.close()
    return r["role"] if r else None

def get_role_region(username):
    conn = get_db(); cur = conn.cursor()
    r = cur.execute("SELECT region FROM users WHERE username=?", (username,)).fetchone()
    conn.close()
    return r["region"] if r else None

def main():
    init_db()
    import socket
    port = PORT
    while True:
        try:
            server = http.server.ThreadingHTTPServer(("0.0.0.0", port), Handler)
            break
        except OSError:
            port += 1
    print(f"考勤核对 Web 已启动： http://localhost:{port}")
    print(f"管理员账号： admin / 密码： {ADMIN_PASS}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        server.shutdown()

if __name__ == "__main__":
    main()
