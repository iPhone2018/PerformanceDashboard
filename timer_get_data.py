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
    pip install requests requests-ntlm schedule xlrd
"""

import os
import re
import sys
import time
import html
import glob
import logging
import datetime
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

REPORT_MAP = {
    "alarm_day": "Users/administrator/Alarm-day",
    "alarm_hour": "Users/administrator/Alarm-hour",
    "oper_day": "Users/administrator/Operation-day",
    "oper_hour": "Users/administrator/Operation-hour",
}
VIEWSTATEGENERATOR_FALLBACK = "8BBF9466"

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
RE_DATE = re.compile(r"(\d{4})[/-](\d{1,2})[/-](\d{1,2})\s+(\d{1,2}):(\d{2}):(\d{2})")

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


def hidden_value(page: str, name: str):
    """取 <input type="hidden" name="..."> 的值，容忍属性顺序变化。"""
    for tag in re.findall(r"<input[^>]*>", page, re.I):
        if re.search(r'name="%s"' % re.escape(name), tag, re.I):
            m = re.search(r'value="([^"]*)"', tag, re.I)
            if m:
                return html.unescape(m.group(1))
    return None


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
    seen = []
    for r in range(min(sheet.nrows, 15)):
        for c in range(min(sheet.ncols, 8)):
            text = _cell_text(sheet, r, c)
            matched = RE_DATE.search(text)
            if matched and len(seen) < 2:
                seen.append(matched)
            for key in ("start", "end"):
                if key in found or not text.strip().lower().startswith(key + ":"):
                    continue
                for c2 in range(c, min(sheet.ncols, c + 4)):
                    m2 = RE_DATE.search(_cell_text(sheet, r, c2))
                    if m2:
                        found[key] = datetime.datetime(*[int(x) for x in m2.groups()])
                        break
    if len(seen) >= 2:
        for key, m in zip(("start", "end"), seen):
            found.setdefault(key, datetime.datetime(*[int(x) for x in m.groups()]))
    return found.get("start"), found.get("end")


def read_report_range(data: bytes):
    """从 .xls 报表头读回真实时间段，并附带各工作表表头文本供排错。

    不猜固定格子、也不假设只有一张表 —— 逐个工作表扫，谁先扫出 Start 就用谁。
    """
    import xlrd
    book = xlrd.open_workbook(file_contents=data, logfile=NULL_LOG)
    texts = []
    for idx in range(book.nsheets):
        sheet = book.sheet_by_index(idx)
        start, end = _scan_sheet(sheet)
        texts.append(sheet_head_text(sheet))
        if start is not None:
            return start, end, "\n".join(texts)
    return None, None, "\n".join(texts)


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

    viewstate = hidden_value(resp.text, "__VIEWSTATE")
    if not viewstate:
        raise RuntimeError("页面里没有 __VIEWSTATE，可能被重定向到登录页")
    viewstategen = hidden_value(resp.text, "__VIEWSTATEGENERATOR") or VIEWSTATEGENERATOR_FALLBACK

    form = dict(BASE_FORM)
    form["__VIEWSTATE"] = viewstate
    form["__VIEWSTATEGENERATOR"] = viewstategen
    # 字面空格，别手工转义
    form["ctl22$StartTimeHidden"] = start_dt.strftime("%Y-%m-%d %H:%M:%S")
    form["ctl22$EndTimeHidden"] = end_dt.strftime("%Y-%m-%d %H:%M:%S")
    form["ctl22$startTimeField"] = field_time(start_dt)
    form["ctl22$endTimeField"] = field_time(end_dt)

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

    got_start, got_end, head = read_report_range(data)
    if got_start is None:
        # 把表头打进报错里：万一版式又变了，日志本身就能说明问题，不用再猜
        raise RuntimeError("读不出报表头里的 Start。报表头实际内容：\n" + head)
    if abs((got_start - start_dt).total_seconds()) > 60:
        raise RuntimeError(f"时间段不符：请求 {start_dt} ~ {end_dt}，文件里是 "
                           f"{got_start} ~ {got_end}\n报表头实际内容：\n" + head)

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
    """最近 DAILY_CHECK_DAYS 个已走完的日子，外加今天。"""
    today = datetime.date.today()
    days = [today - datetime.timedelta(days=d) for d in range(1, DAILY_CHECK_DAYS + 1)]
    if CURRENT_BUCKET_MODE != "skip":
        days.insert(0, today)
    for day in days:
        start_dt = datetime.datetime.combine(day, datetime.time(0, 0, 0))
        end_dt = start_dt + datetime.timedelta(days=1)
        is_current = day == today
        for report_key, prefix in (("alarm_day", "Alarm"), ("oper_day", "Operation")):
            dest = os.path.join(DIR_DAILY, f"{prefix}_{day:%Y%m%d}.xls")
            if not needs_fetch(dest, end_dt, is_current):
                continue
            try_fetch(sess, report_key, start_dt, end_dt, dest)


def collect_hourly(sess):
    """最近 HOURLY_CHECK_HOURS 个已走完的整点，外加当前这个小时。"""
    current = datetime.datetime.now().replace(minute=0, second=0, microsecond=0)
    hours = [current - datetime.timedelta(hours=h) for h in range(1, HOURLY_CHECK_HOURS + 1)]
    if CURRENT_BUCKET_MODE != "skip":
        hours.insert(0, current)
    for start_dt in hours:
        end_dt = start_dt + datetime.timedelta(hours=1)
        is_current = start_dt == current
        for report_key, prefix in (("alarm_hour", "Alarm"), ("oper_hour", "Operation")):
            dest = os.path.join(DIR_HOURLY, f"{prefix}_{start_dt:%Y%m%d_%H}00.xls")
            if not needs_fetch(dest, end_dt, is_current):
                continue
            try_fetch(sess, report_key, start_dt, end_dt, dest)


def main_task():
    started = datetime.datetime.now()
    log.info("========== 任务开始 %s ==========", started)
    try:
        clean_old_files(DIR_DAILY)
        clean_old_files(DIR_HOURLY)
    except Exception:
        log.exception("清理过期文件异常")

    sess = make_session()
    try:
        for name, job in (("日报", collect_daily), ("时报", collect_hourly)):
            try:
                job(sess)
            except Exception:
                log.exception("%s采集异常", name)
    finally:
        sess.close()
    log.info("========== 任务结束，耗时 %.0f 秒 ==========",
             (datetime.datetime.now() - started).total_seconds())


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

    if "--once" in sys.argv:            # 配合任务计划程序：跑一轮就退出
        main_task()
        return

    import schedule
    schedule.every(RUN_EVERY_MINUTES).minutes.do(main_task)
    log.info("已启动，每 %d 分钟执行一次，Ctrl+C 停止", RUN_EVERY_MINUTES)
    main_task()
    while True:
        try:
            schedule.run_pending()
        except Exception:
            log.exception("调度异常")
        time.sleep(10)


if __name__ == "__main__":
    main()
