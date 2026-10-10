# -*- coding: utf-8 -*-
"""
DNA Pareto 报表自动采集

完整链路（IE11 抓包逐字段证实，三步必须共用同一个会话段）：

    ① GET  /DNAreport/Viewer.aspx?report=<报表路径>
           → 302 跳到 /DNAreport/(S(xxx))/Viewer.aspx?report=<报表路径>
           → 从最终 URL 取会话段，从 HTML 取 __VIEWSTATE / __VIEWSTATEGENERATOR
    ② POST /DNAreport/(S(xxx))/Viewer.aspx?report=<报表路径>
           带 67 个表单字段（含新的时间范围）→ 返回 HTML，只是重新渲染，丢弃
    ③ GET  /DNAreport/(S(xxx))/ExcelPreview.aspx
           不带任何参数 → 服务端按会话里记着的时间范围吐出 .xls

三个必须守住的点：
    * 会话段 (S(xxx)) 必须带在 URL 上。否则服务器会 302 派发新会话，而
      requests 遇到 POST 的 302 会把方法降级成 GET，表单直接丢掉。
    * 时间字段传字面空格，不要手工 replace(" ", "+")：requests/urlencode
      会把空格编成 +、把字面 + 编成 %2B，手工替换反而送出错误的加号。
    * ExcelPreview.aspx 返回的是老式 .xls（OLE2），不是 .xlsx。

运行环境（Windows 那台机器）：
    pip install requests requests-ntlm xlrd

两种跑法：
    python timer_get_data.py           常驻，每 RUN_EVERY_MINUTES 分钟一轮
    python timer_get_data.py --once    只跑一轮就退出（配合任务计划程序）

也可以让看板 app.py 起个线程来跑（见 app.py 里的 start_background 调用），
单实例锁保证这两种方式不会同时采集。
"""

import os
import re
import sys
import time
import glob
import logging
import datetime
import threading
from html.parser import HTMLParser
from urllib.parse import quote
from logging.handlers import TimedRotatingFileHandler

import requests
from requests_ntlm import HttpNtlmAuth

class _NullLog:
    """吞掉 xlrd 刷的 OLE2/扇区告警。

    DCS 的 .xls 是第三方库生成的，xlrd 每次打开都会打两行
    "file size ... not 512 + multiple of sector size" /
    "OLE2 inconsistency: SSCS size is 0"，非致命但会把日志刷满。
    注意它是直接 print(file=logfile)，不是 warnings.warn，所以只能从参数压。
    """

    def write(self, *_args):
        return 0

    def flush(self):
        pass


NULL_LOG = _NullLog()

# ===================== 配置区 =====================
BASE_URL = "http://fh01/dna/DNAreport"
DOMAIN_USER = os.environ.get("DNA_USER") or "Administrator"   # 401 的话改成 "fh01\\Administrator"
DOMAIN_PASS = os.environ.get("DNA_PASS") or "#F.h071."

# 采集产物落到「看板项目的 data/」下这两个子目录里，和看板读的是同一处：
#     data/daily/   Alarm_YYYYMMDD.xls        Operation_YYYYMMDD.xls
#     data/hourly/  Alarm_YYYYMMDD_HH00.xls   Operation_YYYYMMDD_HH00.xls
# 部署位置固定，所以路径直接写死 —— 不让它跟着 exe 走。
# 这样即使 exe 被挪动、或者从别的机器拷过来直接跑，数据也只会落在这一处，
# 不会出现「采集写到一个地方、看板读另一个地方」这种最难查的问题。
# （打包成 exe 后 __file__ 指向的是临时解包目录 _MEIxxxxxx，本来也不能用。）
# 要换位置就设环境变量 DASHBOARD_DATA_DIR —— app.py 认的是同一个变量。
DEFAULT_DATA_ROOT = r"C:\Users\Administrator\Desktop\data"

if os.name == "nt":
    DATA_ROOT = os.environ.get("DASHBOARD_DATA_DIR") or DEFAULT_DATA_ROOT
else:
    # 非 Windows（本地开发自测）落到脚本旁边的 data/
    DATA_ROOT = os.environ.get("DASHBOARD_DATA_DIR") or os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "data")
DIR_DAILY = os.path.join(DATA_ROOT, "daily")
DIR_HOURLY = os.path.join(DATA_ROOT, "hourly")
LOG_DIR = os.path.join(DATA_ROOT, "logs")

DAILY_CHECK_DAYS = 7         # 日报回溯天数
HOURLY_CHECK_HOURS = 24      # 时报回溯小时数
RETENTION_DAYS = 30          # 文件保留天数
RUN_EVERY_MINUTES = 30       # 调度间隔

# 当前这个「还没走完」的小时/当天怎么处理：
#   "once"    = 缺就采一次，等它走完之后自动补采一次补齐（默认，稳态每轮 0 请求）
#   "refresh" = 每轮都重采（看板能看到当天/当小时的实时进展，代价是每轮 4 个请求）
#   "skip"    = 干脆不采（看板那一格显示为「缺失」，数据永远只到昨天/上一个整点）
CURRENT_BUCKET_MODE = "once"

# 采集状态 —— 每轮采集的进度和结果。日志汇总用它；采集线程不往外抛异常，
# 出错时这些字段就是事后排查的线索。看板不读它。
STATUS = {
    "enabled": False,       # 这个进程里有没有真的起采集线程
    "running": False,       # 当前是否正在跑一轮
    "rounds": 0,            # 已经跑完的轮数
    "last_start": None,     # 上一轮开始时间
    "last_finish": None,    # 上一轮结束时间
    "last_ok": None,        # 上一次「跑完且没有失败项」的时间
    "last_saved": 0,        # 上一轮保存了几个文件
    "last_failed": 0,       # 上一轮失败了几个
    "last_error": "",       # 最近一次的错误摘要
}

REPORT_MAP = {
    "alarm_day": "Users/administrator/Alarm-day",
    "alarm_hour": "Users/administrator/Alarm-hour",
    "oper_day": "Users/administrator/Operation-day",
    "oper_hour": "Users/administrator/Operation-hour",
}
# 每个报表里明细的 Type 应该是什么 —— 用来校验「拿回来的确实是这个报表」
EXPECTED_TYPE = {
    "alarm_day": "Alarm",
    "alarm_hour": "Alarm",
    "oper_day": "Operation",
    "oper_hour": "Operation",
}

# 抓包逐字段核对过的表单底稿（Alarm 页面）。时间字段在下面按次覆盖。
BASE_FORM = {
    "__EVENTTARGET": "refreshButton",
    "__EVENTARGUMENT": "",
    "__LASTFOCUS": "",
    "ShowPanelsByDefaultHidden": "True",
    "ScaleContentByWidthHidden": "",
    "ctl22$TimeSpanHidden": "Unlimited",
    "ctl22$TimeLineModeHidden": "Normal",
    "ctl22$UiDisabledItemsHidden": "",
    "ctl22$SpanLockHidden": "false",
    "ctl22$ShowParametersHidden": "false",
    "ctl22$AutoUpdateIntervalHidden": "",
    "ctl22$SessionTimeoutHidden": "20",
    "ctl22$MailDefaultRecipientHidden": "",
    "ctl22$MailDefaultSubjectHidden": "",
    "ctl22$MailDefaultMessageHidden": "",
    "ctl22$MailDefaultAttachmentHidden": "",
    "ctl22$SynchronizationEnabledHidden": "false",
    "ctl22$SynchronizationOffsetHidden": "00:00",
    "ctl22$SpanEditMode": "",
    "ctl22$SpanConfigHidden": '{"DAY": "01:00:00","MONTH": "1.00:00:00","YEAR": "MONTH"}',
    "ctl23$CPBodyState": "none",
    "ctl23$EventTypeList$0": "on",
    "ctl23$Priority": "PriorityLevelsRB",
    "ctl23$PriorityLevels$0": "on",
    "ctl23$PriorityLevels$1": "on",
    "ctl23$PriorityLevels$2": "on",
    "ctl23$PriorityLevels$3": "on",
    "ctl23$PriorityLevels$4": "on",
    "ctl23$FilterType": "FilterTypeIn",
    "ctl23$TagName": "",
    "ctl23$TagDescription": "",
    "ctl23$EventDescription": "",
    "ctl23$RowLimit": "5000",
    "ctl23$CurrentTabField": "OAF",
    "ctl23$TabOAFOriginList_table$ctl03$ctl00": "on",
    "ctl23$TabOAFAreaList_table$ctl03$ctl00": "on",
    "ctl23$TabOAFAreaList_table$ctl04$ctl00": "on",
    "ctl23$TabOAFAreaList_table$ctl05$ctl00": "on",
    "ctl23$TabOAFAreaList_table$ctl07$ctl00": "on",
    "ctl23$TabOAFAreaList_table$ctl08$ctl00": "on",
    "ctl23$TabOAFAreaList_table$ctl11$ctl00": "on",
    "ctl23$TabOAFAreaList_table$ctl12$ctl00": "on",
    "ctl23$TabOAFAreaList_table$ctl13$ctl00": "on",
    "ctl23$TabOAFAreaList_table$ctl14$ctl00": "on",
    "ctl23$TabOAFAreaList_table$ctl15$ctl00": "on",
    "ctl23$TabOAFAreaList_table$ctl16$ctl00": "on",
    "ctl23$TabOAFAreaList_table$ctl17$ctl00": "on",
    "ctl23$TabOAFAreaList_table$ctl19$ctl00": "on",
    "ctl23$TabOAFAreaList_table$ctl21$ctl00": "on",
    "ctl23$TabOAFAreaList_table$ctl22$ctl00": "on",
    "ctl23$TabOAFAreaList_table$ctl24$ctl00": "on",
    "ctl23$TabOAFAreaList_table$ctl25$ctl00": "on",
    "ctl23$TabOAFAreaList_table$ctl26$ctl00": "on",
    "ctl23$TabOAFAreaList_table$ctl27$ctl00": "on",
    "ctl23$cbSourceNoFilter": "on",
    "ctl23$cbGroupNoFilter": "on",
    "ctl23$TabSShowGraph": "on",
    "ctl23$TabSBarLimitList": "10",
    "ctl24$CPBodyState": "none",
    "ClientLocalization_Delete": "Really delete report?",
    "ClientLocalization_Save": "The old report will be overwritten and the changes will be visible to all users. Continue with save?",
}
# =================================================

XLS_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
RE_SESSION = re.compile(r"\(S\([^)]+\)\)")
# 报表头的时间有两种写法，都要认：
#   * 手工导出的样张（报警/9.20.xls）：2026/9/20 8:00:00        年在前、24 小时制
#   * ExcelPreview.aspx 导出的：      10/10/2026 12:00:00 AM    月在前、12 小时制
# 后者和页面上报表头显示的格式一致（Start: 10/10/2026 10:41:25 A）。
RE_DATE_YMD = re.compile(r"(\d{4})[/-](\d{1,2})[/-](\d{1,2})[\sT]+(\d{1,2}):(\d{2})(?::(\d{2}))?")
RE_DATE_MDY = re.compile(
    r"(\d{1,2})/(\d{1,2})/(\d{4})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AP])\.?\s*M\.?", re.I)


def parse_dt(text: str):
    """把报表头里的一段文本解析成 datetime，两种版式都认。"""
    m = RE_DATE_YMD.search(text)
    if m:
        y, mo, d, h, mi, s = m.groups()
        return datetime.datetime(int(y), int(mo), int(d), int(h), int(mi), int(s or 0))
    m = RE_DATE_MDY.search(text)
    if m:
        mo, d, y, h, mi, s, ap = m.groups()
        hour = int(h) % 12 + (12 if ap.upper() == "P" else 0)   # 12 AM -> 0 点，12 PM -> 12 点
        return datetime.datetime(int(y), int(mo), int(d), hour, int(mi), int(s or 0))
    return None

log = logging.getLogger("dna")


# ------------------------- 基础设施 -------------------------
def setup_logging():
    try:                       # Windows 控制台/重定向的代码页不是中文时，中文会写成 ????
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    os.makedirs(LOG_DIR, exist_ok=True)
    log.setLevel(logging.INFO)
    fmt = logging.Formatter("%(asctime)s %(levelname)-5s %(message)s", "%Y-%m-%d %H:%M:%S")
    fh = TimedRotatingFileHandler(os.path.join(LOG_DIR, "collect.log"),
                                  when="midnight", backupCount=RETENTION_DAYS,
                                  encoding="utf-8")
    fh.setFormatter(fmt)
    log.addHandler(fh)
    if sys.stdout is not None:          # pythonw 下 stdout 可能是 None
        sh = logging.StreamHandler(sys.stdout)
        sh.setFormatter(fmt)
        log.addHandler(sh)


def acquire_single_instance():
    """崩溃时锁由操作系统自动释放，不会留下死锁。"""
    lock_path = os.path.join(DATA_ROOT, ".timer.lock")
    os.makedirs(DATA_ROOT, exist_ok=True)
    fh = open(lock_path, "a+")
    try:
        if os.name == "nt":
            import msvcrt
            msvcrt.locking(fh.fileno(), msvcrt.LK_NBLCK, 1)
        else:
            import fcntl
            fcntl.flock(fh.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
    except OSError:
        fh.close()
        return None
    return fh


def save_atomic(path, data: bytes):
    """先写 .tmp 再原子替换：中途崩溃不会留下半截文件冒充成品。"""
    tmp = path + ".tmp"
    with open(tmp, "wb") as f:
        f.write(data)
        f.flush()
        os.fsync(f.fileno())
    os.replace(tmp, path)


def clean_old_files(folder: str):
    cutoff = time.time() - RETENTION_DAYS * 86400
    if not os.path.isdir(folder):
        return
    for name in os.listdir(folder):
        path = os.path.join(folder, name)
        if not os.path.isfile(path) or not name.lower().endswith((".xls", ".xlsx")):
            continue
        if os.path.getmtime(path) < cutoff:
            try:
                os.remove(path)
                log.info("清理过期文件 %s", path)
            except OSError as exc:
                log.warning("清理失败 %s: %s", path, exc)


def check_same_stem_duplicates():
    """同一时段同时存在 .xls 和 .xlsx 会被看板重复计入，发现就告警。

    （只看同名同期的，别对历史遗留的一堆 .xlsx 每轮刷屏。）
    """
    for folder in (DIR_DAILY, DIR_HOURLY):
        for xlsx in glob.glob(os.path.join(folder, "*.xlsx")):
            xls = xlsx[:-1]
            if os.path.exists(xls):
                log.warning("同一时段重复：%s 与 %s 会被看板重复计入，请删掉其中一个",
                            os.path.basename(xls), os.path.basename(xlsx))


# ------------------------- 会话与表单 -------------------------
def make_session():
    sess = requests.Session()
    sess.auth = HttpNtlmAuth(DOMAIN_USER, DOMAIN_PASS)
    return sess


class _FormParser(HTMLParser):
    """把页面里的表单控件收成 {name: value}。

    等价于「浏览器点刷新时会提交的内容」。**必须按页面自己的默认值来**，不能照抄
    另一张报表页面的那份：实测导出的是报警还是操作数据，取决于筛选条件里的
    「事件类型」勾了哪一项，而两个报表页面的默认勾选是不同的。
    """

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.fields: dict[str, str] = {}
        self._select = None
        self._select_value = None
        self._select_picked = False
        self._textarea = None

    def handle_starttag(self, tag, attrs):
        attr = {k.lower(): (v if v is not None else "") for k, v in attrs}
        if "disabled" in attr:
            return
        if tag == "input":
            name = attr.get("name")
            if not name:
                return
            kind = attr.get("type", "text").lower()
            if kind in ("submit", "button", "image", "reset", "file"):
                return
            if kind in ("checkbox", "radio"):
                if "checked" in attr:
                    self.fields[name] = attr.get("value") or "on"
            else:
                self.fields[name] = attr.get("value", "")
        elif tag == "select":
            self._select = attr.get("name")
            self._select_value, self._select_picked = None, False
        elif tag == "option" and self._select:
            if "selected" in attr or not self._select_picked:
                self._select_value = attr.get("value", "")
                self._select_picked = True
        elif tag == "textarea":
            self._textarea = attr.get("name")
            if self._textarea:
                self.fields[self._textarea] = ""

    def handle_data(self, data):
        if self._textarea:
            self.fields[self._textarea] += data

    def handle_endtag(self, tag):
        if tag == "select" and self._select:
            self.fields[self._select] = self._select_value or ""
            self._select = None
        elif tag == "textarea":
            self._textarea = None


def build_form(page: str, start_dt: datetime.datetime, end_dt: datetime.datetime) -> dict:
    """按页面实际渲染出来的控件构建 POST 表单，只覆盖时间范围。

    拿不到足够控件时退回内置底稿 BASE_FORM —— 但那样筛选条件就是报警页面的那份，
    操作报表会导出成报警数据（TYPE_MISMATCH 那道校验会拦下来）。
    """
    parser = _FormParser()
    try:
        parser.feed(page)
    except Exception:
        log.warning("解析页面表单失败，退回内置底稿", exc_info=True)
    parsed = parser.fields

    if len(parsed) >= 20:
        form = dict(parsed)
        log.info("按页面实际控件构建表单：%d 个字段", len(form))
    else:
        log.warning("页面里只解析到 %d 个表单控件，退回内置底稿", len(parsed))
        form = dict(BASE_FORM)

    # 状态字段无论如何都要用页面上的（底稿里没有 __VIEWSTATE）
    for key in ("__VIEWSTATE", "__VIEWSTATEGENERATOR", "__EVENTVALIDATION", "__LASTFOCUS"):
        if parsed.get(key):
            form[key] = parsed[key]

    form["__EVENTTARGET"] = "refreshButton"
    form["__EVENTARGUMENT"] = ""

    # 时间字段：按后缀找，免得页面控件编号变了（ctl22 / ctl21）就写不进去。
    # 值用字面空格，别手工转义。
    for suffix, value in (("$StartTimeHidden", start_dt.strftime("%Y-%m-%d %H:%M:%S")),
                          ("$EndTimeHidden", end_dt.strftime("%Y-%m-%d %H:%M:%S")),
                          ("$startTimeField", field_time(start_dt)),
                          ("$endTimeField", field_time(end_dt))):
        names = [key for key in form if key.endswith(suffix)]
        if names:
            for key in names:
                form[key] = value
        else:
            form["ctl22" + suffix] = value
    return form


def field_time(dt: datetime.datetime) -> str:
    """复刻页面时间框的格式：10/10/2026 9:45:25 AM（小时不带前导零）。"""
    hour12 = dt.hour % 12 or 12
    ampm = "AM" if dt.hour < 12 else "PM"
    return f"{dt.month}/{dt.day}/{dt.year} {hour12}:{dt.minute:02d}:{dt.second:02d} {ampm}"


def _cell_text(sheet, r, c) -> str:
    """取单元格文本。日期型单元格必须换算 —— 直接 str() 会得到一串 Excel 序列号数字。"""
    import xlrd
    if r >= sheet.nrows or c >= sheet.ncols:
        return ""
    if sheet.cell_type(r, c) == xlrd.XL_CELL_DATE:
        dt = xlrd.xldate.xldate_as_datetime(sheet.cell_value(r, c), sheet.book.datemode)
        return dt.strftime("%Y/%m/%d %H:%M:%S")
    return str(sheet.cell_value(r, c))


def sheet_head_text(sheet, rows=12, cols=8) -> str:
    """把表头前几行拼成文本 —— 校验失败时打进日志，不用拿到文件也能看出真实版式。"""
    lines = [f"    工作表 [{sheet.name}] {sheet.nrows} 行 x {sheet.ncols} 列"]
    for r in range(min(sheet.nrows, rows)):
        cells = [_cell_text(sheet, r, c).strip()[:22] for c in range(min(sheet.ncols, cols))]
        lines.append(f"    r{r}: " + " | ".join(cells))
    return "\n".join(lines)


def _scan_sheet(sheet):
    """在一张工作表里找 Start / End 时间。

    不猜固定格子 —— 手工导出的样张和程序导出的版式未必一致（吃过这个亏）。
    扫前 15 行找 'Start:' / 'End:' 标签，取它右边最多 3 格里第一个像日期的值；
    万一标签也没有，就退化成「表头里按顺序出现的第 1、2 个日期」。
    """
    found: dict[str, datetime.datetime] = {}
    seen: list[datetime.datetime] = []
    for r in range(min(sheet.nrows, 15)):
        for c in range(min(sheet.ncols, 8)):
            text = _cell_text(sheet, r, c)
            parsed = parse_dt(text)
            if parsed is not None and len(seen) < 2:
                seen.append(parsed)
            for key in ("start", "end"):
                if key in found or not text.strip().lower().startswith(key + ":"):
                    continue
                for c2 in range(c, min(sheet.ncols, c + 4)):
                    parsed2 = parse_dt(_cell_text(sheet, r, c2))
                    if parsed2 is not None:
                        found[key] = parsed2
                        break
    if len(seen) >= 2:
        found.setdefault("start", seen[0])
        found.setdefault("end", seen[1])
    return found.get("start"), found.get("end")


def first_detail_type(sheet) -> str | None:
    """明细区第一行的 Type 值 —— 用来发现「报表选错了」。

    曾经出现过请求操作报表、拿回来的却是报警数据（行数、报表头全对，只有内容是错的）。
    脚本看不出报表选了哪个，但明细里的 Type 骗不了人。
    """
    header_row = None
    for r in range(min(sheet.nrows, 40)):
        values = [_cell_text(sheet, r, c).strip().lower() for c in range(min(sheet.ncols, 9))]
        if all(key in values for key in ("type", "tag", "description", "count")):
            header_row = r
            break
    if header_row is None:
        return None
    type_col = next((c for c in range(min(sheet.ncols, 9))
                     if _cell_text(sheet, header_row, c).strip().lower() == "type"), None)
    if type_col is None:
        return None
    for r in range(header_row + 1, min(sheet.nrows, header_row + 4)):
        value = _cell_text(sheet, r, type_col).strip()
        if value:
            return value
    return None


def read_report_range(data: bytes):
    """从 .xls 报表头读回真实时间段，并附带表头文本和明细类型供排错。

    不猜固定格子、也不假设只有一张表 —— 逐个工作表扫，谁先扫出 Start 就用谁。
    返回 (start, end, 表头文本, 明细第一行的 Type)。
    """
    import xlrd
    book = xlrd.open_workbook(file_contents=data, logfile=NULL_LOG)
    texts = []
    for idx in range(book.nsheets):
        sheet = book.sheet_by_index(idx)
        start, end = _scan_sheet(sheet)
        texts.append(sheet_head_text(sheet))
        if start is not None:
            return start, end, "\n".join(texts), first_detail_type(sheet)
    return None, None, "\n".join(texts), None


# ------------------------- 核心：一次采集 -------------------------
def fetch_report(sess, report_key: str, start_dt: datetime.datetime,
                 end_dt: datetime.datetime, dest_path: str):
    """走完 ①GET → ②POST → ③GET导出 三步，落盘一个 .xls。"""
    report_path = quote(REPORT_MAP[report_key], safe="")

    # ① GET 页面：跟随 302 拿会话段，从 HTML 取 VIEWSTATE
    resp = sess.get(f"{BASE_URL}/Viewer.aspx?report={report_path}", timeout=(10, 60))
    resp.raise_for_status()
    m = RE_SESSION.search(resp.url)
    if not m:
        raise RuntimeError(f"没拿到 cookieless 会话段，最终 URL = {resp.url}")
    sess_seg = m.group(0)
    viewer_url = f"{BASE_URL}/{sess_seg}/Viewer.aspx?report={report_path}"
    export_url = f"{BASE_URL}/{sess_seg}/ExcelPreview.aspx"

    # 诊断用：页面上到底加载的是哪个报表。曾经出现过「请求操作报表却拿回报警数据」，
    # 靠这行能看出是 report 参数没生效、还是导出接口给了旧文件。
    page_title = re.search(r"<title[^>]*>(.*?)</title>", resp.text, re.I | re.S)
    log.info("页面已加载 %s | 标题: %s", resp.url,
             page_title.group(1).strip() if page_title else "取不到")

    # 没有会话段时服务器会 302 到带 (S(xxx)) 的地址，而那个 302 有可能把 query
    # string 丢掉 —— 那样页面加载的就是默认报表，之后 POST/导出全跟着默认报表走
    # （表现为「时间范围变了、报表类型没变」）。检测到就带会话段重新加载一次。
    if "report=" not in resp.url:
        log.info("302 把 report 参数丢了，改用带会话段的地址重新加载")
        resp = sess.get(viewer_url, timeout=(10, 60))
        resp.raise_for_status()
        if "report=" not in resp.url:
            raise RuntimeError(f"report 参数始终没生效，最终 URL = {resp.url}")

    form = build_form(resp.text, start_dt, end_dt)
    if not form.get("__VIEWSTATE"):
        raise RuntimeError("页面里没有 __VIEWSTATE，可能被重定向到登录页")

    # ② POST 刷新：把时间范围写进服务端会话，响应是 HTML，丢弃
    posted = sess.post(viewer_url, data=form, timeout=(10, 120),
                       headers={"Referer": viewer_url})
    posted.raise_for_status()

    # ③ GET 导出：同一个会话，无参数，拿到 .xls
    exported = sess.get(export_url, timeout=(10, 180),
                        headers={"Referer": viewer_url})
    exported.raise_for_status()
    data = exported.content

    if not data.startswith(XLS_MAGIC):
        head = data[:300].decode("utf-8", "replace")
        raise RuntimeError(f"导出不是 .xls（头 {data[:8].hex()}），响应片段: {head}")

    got_start, got_end, head, detail_type = read_report_range(data)
    if got_start is None:
        # 把表头打进报错里：万一版式又变了，日志本身就能说明问题，不用再猜
        raise RuntimeError("读不出报表头里的 Start。报表头实际内容：\n" + head)
    if abs((got_start - start_dt).total_seconds()) > 60:
        raise RuntimeError(f"时间段不符：请求 {start_dt} ~ {end_dt}，文件里是 "
                           f"{got_start} ~ {got_end}\n报表头实际内容：\n" + head)

    # 报表选错了就拦下来 —— 宁可报错重试，也不能把报警数据当成操作数据存进去
    expected = EXPECTED_TYPE[report_key]
    if detail_type and detail_type.lower() != expected.lower():
        raise RuntimeError(f"报表选错了：请求 {REPORT_MAP[report_key]}（应为 {expected}），"
                           f"但导出内容里的 Type 是 {detail_type}\n报表头实际内容：\n" + head)

    save_atomic(dest_path, data)
    log.info("已保存 %s（%s ~ %s，%d 字节）",
             os.path.basename(dest_path), got_start, got_end, len(data))


def try_fetch(sess, report_key, start_dt, end_dt, dest_path, attempts=3):
    for attempt in range(1, attempts + 1):
        try:
            fetch_report(sess, report_key, start_dt, end_dt, dest_path)
            return True
        except Exception as exc:
            log.warning("第 %d/%d 次失败 %s: %s",
                        attempt, attempts, os.path.basename(dest_path), exc)
            if attempt < attempts:
                time.sleep(5 * attempt)
    log.error("放弃 %s", os.path.basename(dest_path))
    return False


# ------------------------- 采集计划 -------------------------
def needs_fetch(dest_path: str, bucket_end: datetime.datetime, is_current: bool) -> bool:
    """这个时段要不要采。

    核心规则：**文件在就不重复采**，只有三种情况例外 ——
      * 文件不存在（缺了就补）
      * 文件是空的/坏的（size 为 0）
      * 已走完的时段，但那份文件是在它走完**之前**采的（mtime < bucket_end），
        说明内容是残的 —— 补采一次把它补全，之后 mtime 就新过 bucket_end，不再采

    这样当前这一小时/当天第一次采到之后就不用每轮重抓，等它走完会自动补齐一次；
    稳态下每轮 0 个请求。
    """
    if not os.path.exists(dest_path):
        return not (is_current and CURRENT_BUCKET_MODE == "skip")
    if os.path.getsize(dest_path) == 0:
        return True
    if is_current:
        return CURRENT_BUCKET_MODE == "refresh"
    return os.path.getmtime(dest_path) < bucket_end.timestamp()


def collect_daily(sess):
    """最近 DAILY_CHECK_DAYS 个已走完的日子，外加今天。返回 (成功数, 失败数)。"""
    today = datetime.date.today()
    days = [today - datetime.timedelta(days=d) for d in range(1, DAILY_CHECK_DAYS + 1)]
    if CURRENT_BUCKET_MODE != "skip":
        days.insert(0, today)
    saved = failed = 0
    for day in days:
        start_dt = datetime.datetime.combine(day, datetime.time(0, 0, 0))
        end_dt = start_dt + datetime.timedelta(days=1)
        is_current = day == today
        for report_key, prefix in (("alarm_day", "Alarm"), ("oper_day", "Operation")):
            dest = os.path.join(DIR_DAILY, f"{prefix}_{day:%Y%m%d}.xls")
            if not needs_fetch(dest, end_dt, is_current):
                continue
            if try_fetch(sess, report_key, start_dt, end_dt, dest):
                saved += 1
            else:
                failed += 1
    return saved, failed


def collect_hourly(sess):
    """最近 HOURLY_CHECK_HOURS 个已走完的整点，外加当前这个小时。返回 (成功数, 失败数)。"""
    current = datetime.datetime.now().replace(minute=0, second=0, microsecond=0)
    hours = [current - datetime.timedelta(hours=h) for h in range(1, HOURLY_CHECK_HOURS + 1)]
    if CURRENT_BUCKET_MODE != "skip":
        hours.insert(0, current)
    saved = failed = 0
    for start_dt in hours:
        end_dt = start_dt + datetime.timedelta(hours=1)
        is_current = start_dt == current
        for report_key, prefix in (("alarm_hour", "Alarm"), ("oper_hour", "Operation")):
            dest = os.path.join(DIR_HOURLY, f"{prefix}_{start_dt:%Y%m%d_%H}00.xls")
            if not needs_fetch(dest, end_dt, is_current):
                continue
            if try_fetch(sess, report_key, start_dt, end_dt, dest):
                saved += 1
            else:
                failed += 1
    return saved, failed


def main_task():
    """跑一轮完整采集。

    **绝不向外抛异常** —— 它可能跑在看板进程的线程里，异常一冒出去线程就死了，
    而页面照常打开、看起来毫无异样。所有分支都就地兜住，并把状态写进 STATUS。
    """
    started = datetime.datetime.now()
    STATUS["running"] = True
    STATUS["last_start"] = started
    STATUS["last_error"] = ""
    log.info("========== 任务开始 %s ==========", started)

    saved = failed = 0
    try:
        try:
            clean_old_files(DIR_DAILY)
            clean_old_files(DIR_HOURLY)
            clean_old_files(QUARANTINE_DIR)
        except Exception as exc:
            log.exception("清理过期文件异常")
            STATUS["last_error"] = f"清理异常: {exc}"

        sess = make_session()
        try:
            for name, job in (("日报", collect_daily), ("时报", collect_hourly)):
                try:
                    got_saved, got_failed = job(sess)
                    saved += got_saved
                    failed += got_failed
                except Exception as exc:
                    log.exception("%s采集异常", name)
                    STATUS["last_error"] = f"{name}异常: {exc}"
        finally:
            sess.close()
    except Exception as exc:                       # make_session 等兜底
        log.exception("采集轮次异常")
        STATUS["last_error"] = f"{type(exc).__name__}: {exc}"
    finally:
        finished = datetime.datetime.now()
        STATUS["running"] = False
        STATUS["last_finish"] = finished
        STATUS["last_saved"] = saved
        STATUS["last_failed"] = failed
        STATUS["rounds"] += 1
        if failed == 0 and not STATUS["last_error"]:
            STATUS["last_ok"] = finished
        log.info("========== 任务结束，保存 %d 个、失败 %d 个，耗时 %.0f 秒 ==========",
                 saved, failed, (finished - started).total_seconds())


def seconds_until_next_round() -> float:
    """对齐到整点/半点再跑 —— 整点刚过就立刻去补上一个小时，不会拖到一个小时后。"""
    period = RUN_EVERY_MINUTES * 60
    return max(5.0, period - (time.time() % period))


def _run_forever():
    while True:
        try:
            main_task()
        except Exception:                  # main_task 已经不抛了，这是双保险
            log.exception("采集线程捕到异常，继续下一轮")
        time.sleep(seconds_until_next_round())


def start_background():
    """在看板进程里起一个守护线程跑采集。

    单实例锁保证「看板里起了线程」和「旁边另跑一个 timer_get_data.py」不会同时采。
    起不来就返回 None，看板照常服务。
    """
    setup_logging()
    os.makedirs(DIR_DAILY, exist_ok=True)
    os.makedirs(DIR_HOURLY, exist_ok=True)

    lock = acquire_single_instance()
    if lock is None:
        log.error("已有采集实例在运行，看板内不再重复启动")
        STATUS["enabled"] = False
        STATUS["last_error"] = "已有采集实例在运行"
        return None

    log.info("本机时间 %s（时区 %s），请确认与 DNA 服务器一致",
             datetime.datetime.now(), time.tzname)
    check_same_stem_duplicates()
    STATUS["enabled"] = True
    thread = threading.Thread(target=_run_forever, name="dna-collector", daemon=True)
    thread.start()
    log.info("看板内已启动采集线程，每 %d 分钟一轮", RUN_EVERY_MINUTES)
    return thread


def main():
    setup_logging()
    os.makedirs(DIR_DAILY, exist_ok=True)
    os.makedirs(DIR_HOURLY, exist_ok=True)

    lock = acquire_single_instance()
    if lock is None:
        log.error("已有实例在运行，本次退出")
        sys.exit(1)

    log.info("本机时间 %s（时区 %s），请确认与 DNA 服务器一致",
             datetime.datetime.now(), time.tzname)
    check_same_stem_duplicates()
    STATUS["enabled"] = True

    if "--once" in sys.argv:            # 配合任务计划程序：跑一轮就退出
        main_task()
        return

    log.info("已启动，每 %d 分钟执行一次，Ctrl+C 停止", RUN_EVERY_MINUTES)
    _run_forever()


if __name__ == "__main__":
    main()
