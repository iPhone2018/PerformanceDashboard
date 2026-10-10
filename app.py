# -*- coding: utf-8 -*-
"""Valmet DCS 操作 / 报警 Pareto 看板 —— 后端服务

数据来源（自动扫描，无需改代码）：
    data/hourly/{Operation|Alarm}_YYYYMMDD_HHMM.xlsx   按小时的 Pareto 明细
    data/daily/{Operation|Alarm}_YYYYMMDD.xlsx         按天的 Pareto 明细
文件版式与 DCS 导出的 Pareto Report 一致：顶部若干行报表头 + 明细表
（列固定为 Type / Tag / Description / Count），报表头里带 Start / End 时间。

接口：
    GET /                    看板页面
    GET /api/data            聚合数据（两条趋势 + 四张 TOP10 排行图）
    GET /api/download/hour   某小时的原始明细 Excel（key=YYYYMMDDHH）
    GET /api/download/day    某天的原始明细 Excel（key=YYYY-MM-DD）
    GET /api/download/window 某时段某一类事件的全部原始明细（range=24h|7d, type=Operation|Alarm）

统计口径：
    * 近 24 小时 -> 取 hourly 目录，以最新一个小时文件为基准往前 24 个整点；
    * 近 7 天    -> 取 daily 目录，以最新一天文件为基准往前 7 天；
    * 操作 / 报警两类事件全程分开统计，占比的分母是同时段同类事件总数。
    基准取“数据里最新的时间”，而不是机器当前时间，这样历史导出数据也能正常看。
"""
from __future__ import annotations

import glob
import io
import os
import re
import sys
from datetime import date, datetime, timedelta

import pandas as pd
from flask import Flask, jsonify, make_response, render_template, request

from data_paths import ENV_VAR, data_root, describe

# 输出强制 UTF-8。Windows 英文系统的控制台和重定向都是 cp1252，打印中文会直接
# UnicodeEncodeError 崩掉 —— 启动阶段就崩，页面根本起不来。放在任何 print 之前。
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 数据目录从 data_paths 取 —— 那份定义和采集端 timer_get_data.py 共用，
# 两边不可能指岔。"采集写到一个地方、看板读另一个地方"是最难查的一类问题：
# 两边都不报错，日志里采集还全是"已保存"，页面却是空的。
DATA_DIR = data_root()
HOURLY_DIR = os.path.join(DATA_DIR, "hourly")
DAILY_DIR = os.path.join(DATA_DIR, "daily")

HOUR_FMT = "%Y-%m-%d %H:%M"
DAY_FMT = "%Y-%m-%d"

HOURLY_FILE_RE = re.compile(r"^(?P<type>[A-Za-z]+)_(?P<day>\d{8})_(?P<hm>\d{4})\.xlsx?$")
DAILY_FILE_RE = re.compile(r"^(?P<type>[A-Za-z]+)_(?P<day>\d{8})\.xlsx?$")
START_TIME_RE = re.compile(r"Start:\s*(\d{4})[/-](\d{1,2})[/-](\d{1,2})[\sT]+(\d{1,2}):(\d{2})")
DATE_TIME_RE = re.compile(r"(\d{4})[/-](\d{1,2})[/-](\d{1,2})[\sT]+(\d{1,2}):(\d{2}):(\d{2})")
# ExcelPreview.aspx 导出的报表头是「月/日/年 12 小时制」：10/10/2026 12:00:00 AM
DATE_TIME_US_RE = re.compile(
    r"(\d{1,2})/(\d{1,2})/(\d{4})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AP])\.?\s*M\.?", re.I)


def _matched_to_datetime(matched: re.Match) -> datetime:
    """把上面三种正则的匹配结果统一成 datetime（靠分组个数区分）。"""
    groups = matched.groups()
    if len(groups) == 7:                       # 月/日/年 + 12 小时制
        mo, day, year, hour12, minute, second, ampm = groups
        hour = int(hour12) % 12 + (12 if ampm.upper() == "P" else 0)   # 12 AM→0点, 12 PM→12点
        return datetime(int(year), int(mo), int(day), hour, int(minute), int(second or 0))
    if len(groups) == 6:                       # 年/月/日 + 秒
        return datetime(*[int(x) for x in groups])
    year, mo, day, hour, minute = groups       # 年/月/日，无秒
    return datetime(int(year), int(mo), int(day), int(hour), int(minute))

REQUIRED_COLS = ("type", "tag", "description", "count")


class _NullLog:
    """吞掉 xlrd 往 stderr 刷的 OLE2/扇区告警。

    DCS 导出的 .xls 是第三方库生成的，xlrd 每次打开都会打两行告警；
    它是直接 print(file=logfile) 而不是 warnings.warn，所以只能从参数压。
    """

    def write(self, *_args):
        return 0

    def flush(self):
        pass


_NULL_LOG = _NullLog()

# 解析结果按 (mtime, size) 缓存，刷新页面不重复读盘
_FILE_CACHE: dict[str, tuple[tuple[int, int], tuple[pd.DataFrame, datetime | None]]] = {}

# ---------------------------------------------------------------------------
# 采集端（可选）
# ---------------------------------------------------------------------------
# timer_get_data.py 和看板共用同一份 data/，这里把它当模块用，在看板进程里起一个
# 守护线程跑采集 —— 部署只有一个进程。没装 requests / requests-ntlm / xlrd 时
# 看板照常工作（只是不采集），也可以用 DASHBOARD_COLLECT=0 显式关掉。
COLLECTOR_ENABLED = os.environ.get("DASHBOARD_COLLECT", "1").lower() not in ("0", "false", "no", "off")
try:
    import timer_get_data as collector
except Exception:                              # 缺依赖也要能开看板
    collector = None


def resource_dir() -> str:
    """模板和静态资源在哪。

    源码运行时就在脚本旁边；打包成 exe 后 PyInstaller 会把 --add-data 的内容
    解到 sys._MEIPASS，必须显式指过去，否则页面直接 404。
    """
    return getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))


app = Flask(
    __name__,
    template_folder=os.path.join(resource_dir(), "templates"),
    static_folder=os.path.join(resource_dir(), "static"),
)
app.json.ensure_ascii = False  # 接口返回中文而不是 \uXXXX
app.config["TEMPLATES_AUTO_RELOAD"] = True  # 改完模板刷新页面即可生效


# ---------------------------------------------------------------------------
# 一、Excel 解析
# ---------------------------------------------------------------------------
def _find_header_row(raw: pd.DataFrame) -> int | None:
    """Pareto Report 顶部有报表头，定位真正的中文/英文表头行。"""
    for i in range(min(len(raw), 40)):
        values = [str(v).strip().lower() for v in raw.iloc[i].tolist()]
        if all(col in values for col in REQUIRED_COLS):
            return i
    return None


def _parse_start_time(raw: pd.DataFrame) -> datetime | None:
    """从报表头里读 Start 时间，作为该文件统计区间的起点。

    标签和值可能同格也可能分成两格，时间格式也有两套：
      * 演示版式 / 手工样张：同格 "Start:  2026/09/29 00:00:00"，或 r2c1="Start:"、r2c2="2026/9/20 8:00:00"
      * 采集脚本导出的：r2c1="Start:"、r2c2="10/10/2026 12:00:00 AM"（月/日/年 + 12 小时制）
    """
    for i in range(min(len(raw), 12)):
        for c in range(min(raw.shape[1], 4)):
            text = str(raw.iat[i, c])
            if "start:" not in text.strip().lower():
                continue
            for pattern in (START_TIME_RE, DATE_TIME_RE, DATE_TIME_US_RE):   # 同格
                matched = pattern.search(text)
                if matched:
                    return _matched_to_datetime(matched)
            for c2 in range(c + 1, min(raw.shape[1], c + 4)):               # 标签右侧几格
                neighbour = str(raw.iat[i, c2])
                for pattern in (DATE_TIME_RE, DATE_TIME_US_RE):
                    matched = pattern.search(neighbour)
                    if matched:
                        return _matched_to_datetime(matched)
    return None


def _read_raw(path: str) -> pd.DataFrame:
    """读原始单元格。openpyxl 不认 logfile 参数，所以只给 .xls 传。"""
    if path.lower().endswith(".xls"):
        return pd.read_excel(path, header=None, dtype=object,
                             engine_kwargs={"logfile": _NULL_LOG})
    return pd.read_excel(path, header=None, dtype=object)


def _parse_pareto(path: str) -> tuple[pd.DataFrame, datetime | None]:
    """读一个 Pareto Excel -> (明细 DataFrame, 起始时间)。"""
    raw = _read_raw(path)
    if raw.empty:
        return pd.DataFrame(columns=["Type", "Tag", "Description", "Count"]), None

    header_row = _find_header_row(raw)
    if header_row is None:
        raise ValueError(f"{os.path.basename(path)} 未找到 Type/Tag/Description/Count 表头")

    header = [str(v).strip() for v in raw.iloc[header_row].tolist()]
    # 用位置索引取值，避免同名/空列名（Unnamed: n）互相覆盖
    pos = {}
    for idx, name in enumerate(header):
        pos.setdefault(name.lower(), idx)
    body = raw.iloc[header_row + 1:]

    def column(name: str) -> pd.Series:
        idx = pos.get(name)
        if idx is None:
            return pd.Series([None] * len(body), index=body.index)
        return body.iloc[:, idx]

    df = pd.DataFrame(
        {
            "Type": column("type"),
            "Tag": column("tag"),
            "Description": column("description"),
            "Count": column("count"),
        }
    )
    df = df.dropna(subset=["Tag", "Count"])
    df["Tag"] = df["Tag"].astype(str).str.strip()
    df["Description"] = df["Description"].fillna("").astype(str).str.strip()
    df["Type"] = df["Type"].fillna("Unknown").astype(str).str.strip().str.title()
    # 导出走 US 文化时，四位以上的计数会写成 1,234；to_numeric 遇到逗号会判成
    # 非法字符 → 静默归 0。先剥掉千分位再转。
    df["Count"] = pd.to_numeric(
        df["Count"].astype(str).str.replace(",", "", regex=False), errors="coerce"
    ).fillna(0).astype("int64")
    df = df[(df["Tag"] != "") & (df["Tag"].str.lower() != "nan") & (df["Count"] > 0)]
    return df.reset_index(drop=True), _parse_start_time(raw)


def read_pareto(path: str) -> tuple[pd.DataFrame, datetime | None]:
    """带缓存的读取。"""
    stat = os.stat(path)
    stamp = (stat.st_mtime_ns, stat.st_size)
    cached = _FILE_CACHE.get(path)
    if cached is not None and cached[0] == stamp:
        return cached[1]
    result = _parse_pareto(path)
    _FILE_CACHE[path] = (stamp, result)
    return result


def _excel_files(directory: str) -> list[str]:
    if not os.path.isdir(directory):
        return []
    files = [p for p in glob.glob(os.path.join(directory, "*.xls")) +
             glob.glob(os.path.join(directory, "*.xlsx"))
             if not os.path.basename(p).startswith("~$")]
    return sorted(files)


def _load_buckets(directory: str, pattern: re.Pattern, *, by_hour: bool,
                  warnings: list[str]) -> dict:
    """把一个目录下的文件合并成 {时间桶: {类型: DataFrame}}。

    桶里**没有**某个类型 = 该时段没有对应的 Excel 文件（数据缺失）；
    桶里**有**但 DataFrame 为空 = 文件存在、只是里面没有明细行（真实的 0 条事件）。
    这两种情况在看板上必须区分开，否则"没采到"会被画成"没有事件"。
    """
    buckets: dict = {}
    for path in _excel_files(directory):
        name = os.path.basename(path)
        matched = pattern.match(name)
        if matched is None:
            warnings.append(f"{name}: 文件名不符合命名规范，已跳过")
            continue
        try:
            df, start_time = read_pareto(path)
        except Exception as exc:  # 单文件坏掉不影响整体看板
            warnings.append(f"{name}: {exc}")
            continue

        if start_time is not None:
            key = start_time
        elif by_hour:  # 报表头里没写 Start，退回用文件名
            key = datetime.strptime(matched.group("day") + matched.group("hm"), "%Y%m%d%H%M")
        else:
            key = datetime.strptime(matched.group("day"), "%Y%m%d")
        # 小时桶对齐到整点、天桶对齐到日期
        key = key.replace(minute=0, second=0, microsecond=0)
        if not by_hour:
            key = key.date()

        file_type = _normalize_type(matched.group("type")) or matched.group("type").strip().title()
        slot = buckets.setdefault(key, {})
        slot.setdefault(file_type, _empty_detail())  # 先占位：文件存在（哪怕 0 行）
        for row_type, part in (df.groupby("Type", sort=False) if not df.empty else []):
            row_type = _normalize_type(row_type) or str(row_type).strip().title()
            if row_type in slot and not slot[row_type].empty:
                slot[row_type] = pd.concat([slot[row_type], part], ignore_index=True)
            else:
                slot[row_type] = part.reset_index(drop=True)
    return buckets


def load_hourly(warnings: list[str]) -> dict:
    return _load_buckets(HOURLY_DIR, HOURLY_FILE_RE, by_hour=True, warnings=warnings)


def load_daily(warnings: list[str]) -> dict:
    return _load_buckets(DAILY_DIR, DAILY_FILE_RE, by_hour=False, warnings=warnings)


# ---------------------------------------------------------------------------
# 二、聚合
# ---------------------------------------------------------------------------
TYPE_LABELS = {"Operation": "操作", "Alarm": "报警"}
DETAIL_COLUMNS = ["Type", "Tag", "Description", "Count"]
TIME_COLUMN = "时间"


def _empty_detail() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "Type": pd.Series(dtype="object"),
            "Tag": pd.Series(dtype="object"),
            "Description": pd.Series(dtype="object"),
            "Count": pd.Series(dtype="int64"),
        }
    )


def _bucket_frame(bucket: dict | None) -> pd.DataFrame:
    """把 {类型: DataFrame} 的时间桶拼回一张明细表（下载用）。"""
    if not bucket:
        return _empty_detail()
    frames = [frame for frame in bucket.values() if frame is not None and not frame.empty]
    return pd.concat(frames, ignore_index=True) if frames else _empty_detail()


def _normalize_type(value: str | None) -> str | None:
    """把 op / operation / alm / alarm 之类的写法统一成 Operation / Alarm。"""
    text = (value or "").strip().lower()
    if text.startswith("op"):
        return "Operation"
    if text.startswith("al"):
        return "Alarm"
    return None


def _slice_type(df: pd.DataFrame, type_name: str) -> pd.DataFrame:
    """按操作/报警把明细切开，两类事件分开统计。"""
    if df is None or df.empty:
        return _empty_detail()
    return df[df["Type"].str.lower() == type_name.lower()]


def _top_tags(df: pd.DataFrame, total: int, type_name: str, limit: int = 10) -> list[dict]:
    """按 Tag 汇总取 TOP N。占比 = 该 Tag 次数 / 同时段同类事件总数。"""
    if df is None or df.empty:
        return []
    grouped = (
        df.groupby("Tag", as_index=False)
        .agg(Count=("Count", "sum"), Description=("Description", "first"))
        .sort_values(["Count", "Tag"], ascending=[False, True])
    )
    return [
        {
            "tag": row.Tag,
            "description": row.Description,
            "type": type_name,
            "type_label": TYPE_LABELS.get(type_name, type_name),
            "count": int(row.Count),
            "pct": round(row.Count / total * 100, 1) if total else 0.0,
        }
        for row in grouped.head(limit).itertuples(index=False)
    ]


def _slot_value(slot: dict | None, type_name: str) -> tuple[int | None, str]:
    """取某个时段某个类型的次数。

    返回 (值, 状态)：missing = 没有这个文件（值 None，折线断开）；empty = 文件在但 0 条记录（值 0）。
    """
    if not slot or type_name not in slot:
        return None, "missing"
    frame = slot[type_name]
    if frame is None or frame.empty:
        return 0, "empty"
    return int(frame["Count"].sum()), "ok"


def _collect(directory_buckets: dict, slots: list, key_fmt: str, label_fmt: str,
             time_fmt: str) -> dict:
    """按槽位汇总一个窗口，同时记录缺失 / 空文件，供前端画缺口和提示。"""
    labels, keys, op, alarm, parts = [], [], [], [], []
    missing, empty_slots, covered = [], [], 0
    for slot in slots:
        data = directory_buckets.get(slot)
        label = slot.strftime(label_fmt)
        labels.append(label)
        keys.append(slot.strftime(key_fmt))
        op_value, op_state = _slot_value(data, "Operation")
        alarm_value, alarm_state = _slot_value(data, "Alarm")
        op.append(op_value)
        alarm.append(alarm_value)
        if data:
            covered += 1
            if "empty" in (op_state, alarm_state):
                empty_slots.append(label)
            for frame in data.values():
                if frame.empty:
                    continue
                part = frame.copy()
                part.insert(0, TIME_COLUMN, slot.strftime(time_fmt))
                parts.append(part)
        else:
            # 两种类型都没有文件：整个时段缺失
            missing.append(label)

    frame = pd.concat(parts, ignore_index=True) if parts else _empty_detail()
    if not parts:
        frame.insert(0, TIME_COLUMN, pd.Series(dtype="object"))

    return {
        "labels": labels, "keys": keys, "op": op, "alarm": alarm,
        "frame": frame, "missing": missing, "empty": empty_slots,
        "covered": covered, "total_slots": len(slots),
    }


def collect_windows(hourly: dict, daily: dict) -> dict:
    """把「近 24 小时」「近 7 天」两个窗口整理好，趋势图与 TOP10 共用同一份口径。

    基准取数据里最新的整点 / 最新的一天。缺文件的时段折线断开（不是 0），
    文件存在但没记录的时段按真实的 0 处理，两者都会在页面上单独标出。
    """
    hour_anchor = max(hourly) if hourly else None
    day_anchor = max(daily) if daily else (hour_anchor.date() if hour_anchor else date.today())

    hour_slots = [hour_anchor - timedelta(hours=i) for i in range(23, -1, -1)] if hour_anchor else []
    hour = _collect(hourly, hour_slots, "%Y%m%d%H", "%m-%d %H:00", "%Y-%m-%d %H:00")
    day = _collect(daily, [day_anchor - timedelta(days=i) for i in range(6, -1, -1)],
                   DAY_FMT, DAY_FMT, DAY_FMT)

    hour_range = "-"
    if hour_anchor is not None:
        # 整个 24 小时窗口：从最早那个整点到最后一个整点的结束时刻
        hour_range = (f"{(hour_anchor - timedelta(hours=23)).strftime('%m-%d %H:%M')} ~ "
                      f"{(hour_anchor + timedelta(hours=1)).strftime('%m-%d %H:%M')}")
    day_range = f"{(day_anchor - timedelta(days=6)).strftime(DAY_FMT)} ~ {day_anchor.strftime(DAY_FMT)}"

    return {
        "hour_anchor": hour_anchor,
        "day_anchor": day_anchor,
        "hour_frame": hour["frame"],
        "week_frame": day["frame"],
        "hour_range": hour_range,
        "day_range": day_range,
        "hour": hour,
        "day": day,
    }


def _preview(labels: list[str], limit: int = 3) -> str:
    head = "、".join(labels[:limit])
    return head + " 等" if len(labels) > limit else head


def _window_quality(window_info: dict, unit: str) -> str:
    """数据完整性描述：缺多少个时段、多少个文件是空的。"""
    problems = []
    missing = window_info.get("missing") or []
    empty = window_info.get("empty") or []
    if missing:
        problems.append(f"缺 {len(missing)} 个{unit}的数据（{_preview(missing)}）")
    if empty:
        problems.append(f"{len(empty)} 个文件里没有明细记录（{_preview(empty)}）")
    return "；".join(problems)


def _empty_dir_note(window_info: dict, directory: str, unit: str) -> str:
    """整个窗口一个文件都没有时的提示（否则页面上只有一张空图，看不出是没数据还是画错了）。"""
    if window_info["total_slots"]:
        return ""
    return f"data/{directory} 下没有任何符合命名规范的{unit}文件"


def _chart_block(window: str, type_name: str, frame: pd.DataFrame, range_label: str,
                 window_info: dict) -> dict:
    """一张 TOP10 柱状图需要的一切：标题、口径说明、数据完整性、下载入口。"""
    sub = _slice_type(frame, type_name)
    total = int(sub["Count"].sum()) if not sub.empty else 0
    type_label = TYPE_LABELS.get(type_name, type_name)
    window_label = "近 24 小时" if window == "24h" else "近 7 天"
    unit = "整点" if window == "24h" else "天"
    covered, total_slots = window_info["covered"], window_info["total_slots"]
    if total_slots:
        coverage = f"数据覆盖 {covered}/{total_slots} 个{unit}"
    else:  # 整个窗口一个文件都没有，说"覆盖 0/0"没有意义
        coverage = f"未找到{'小时' if window == '24h' else '天'}文件"
    return {
        "window": window,
        "type": type_name,
        "title": f"{window_label} TOP10 Tag · {type_label}",
        "note": (f"统计区间 {range_label} · {type_label}事件共 {total:,} 次 · "
                 f"{coverage} · 占比为该类事件内占比"),
        "data_note": _window_quality(window_info, unit),
        "total": total,
        "items": _top_tags(sub, total, type_name),
        "download_url": f"/api/download/window?range={window}&type={type_name}",
        "download_label": f"下载{window_label}全部{type_label}明细",
    }


def build_dashboard() -> dict:
    warnings: list[str] = []
    hourly = load_hourly(warnings)
    daily = load_daily(warnings)

    if not hourly and not daily:
        return {
            "ok": False,
            "message": f"{DATA_DIR} 下没有可用的 Pareto Excel，"
                       "请把导出的文件按命名规范放入 hourly/ 与 daily/ 后刷新。",
            "meta": {
                "data_dir": DATA_DIR,
                "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            },
            "warnings": warnings,
        }

    windows = collect_windows(hourly, daily)
    hour_anchor = windows["hour_anchor"]
    hour, day = windows["hour"], windows["day"]

    charts = [
        _chart_block("24h", "Operation", windows["hour_frame"], windows["hour_range"], hour),
        _chart_block("24h", "Alarm", windows["hour_frame"], windows["hour_range"], hour),
        _chart_block("7d", "Operation", windows["week_frame"], windows["day_range"], day),
        _chart_block("7d", "Alarm", windows["week_frame"], windows["day_range"], day),
    ]

    return {
        "ok": True,
        "meta": {
            "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "data_dir": DATA_DIR,  # 方便确认看板读的到底是哪个目录（也能让自测校验端口没串）
            "data_through": hour_anchor.strftime(HOUR_FMT) if hour_anchor else "-",
            "hour_window": {
                "label": windows["hour_range"],
                "points": len(hour["labels"]),
                "covered": hour["covered"],
                "missing": hour["missing"],
                "empty": hour["empty"],
                "quality": _window_quality(hour, "整点") or _empty_dir_note(hour, "hourly", "小时"),
            },
            "day_window": {
                "label": windows["day_range"],
                "days": len(day["labels"]),
                "covered": day["covered"],
                "missing": day["missing"],
                "empty": day["empty"],
                "quality": _window_quality(day, "天") or _empty_dir_note(day, "daily", "天"),
            },
            "files": {"hourly": len(_excel_files(HOURLY_DIR)), "daily": len(_excel_files(DAILY_DIR))},
            "warnings": warnings,
        },
        "hour": {"labels": hour["labels"], "keys": hour["keys"], "op": hour["op"], "alarm": hour["alarm"]},
        "day": {"labels": day["labels"], "keys": day["keys"], "op": day["op"], "alarm": day["alarm"]},
        "charts": charts,
    }


# ---------------------------------------------------------------------------
# 三、Excel 导出
# ---------------------------------------------------------------------------
def _autofit(writer: pd.ExcelWriter) -> None:
    """简单列宽自适应，方便直接打开查看。"""
    try:
        for sheet in writer.sheets.values():
            for column in sheet.columns:
                width = max((len(str(cell.value)) for cell in column if cell.value is not None), default=8)
                sheet.column_dimensions[column[0].column_letter].width = min(max(width + 2, 10), 46)
    except Exception:  # 列宽只是锦上添花，失败不影响导出
        pass


def _excel_response(sheets: dict[str, pd.DataFrame], filename: str):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name, index=False)
        _autofit(writer)
    buffer.seek(0)
    response = make_response(buffer.read())
    response.headers["Content-Type"] = (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    response.headers["Content-Disposition"] = f'attachment; filename="{filename}"'
    response.headers["Cache-Control"] = "no-store"
    return response


def _detail_sheets(df: pd.DataFrame, time_column: str | None = None) -> dict[str, pd.DataFrame]:
    """两张表：全部原始明细（不做任何汇总）+ Tag 汇总（口径与看板 TOP10 一致）。

    time_column 不为空时，明细里带上每条记录所属的小时/日期，方便跨时段核对。
    """
    columns = ([time_column] if time_column else []) + DETAIL_COLUMNS
    detail = df.copy()
    for col in columns:
        if col not in detail.columns:
            detail[col] = ""
    detail = detail[columns]
    detail["Count"] = pd.to_numeric(detail["Count"], errors="coerce").fillna(0).astype("int64")
    detail = detail[detail["Count"] > 0]

    sort_columns = ([time_column] if time_column else []) + ["Type", "Count", "Tag"]
    ascending = ([True] if time_column else []) + [True, False, True]
    detail = detail.sort_values(sort_columns, ascending=ascending)

    total = int(detail["Count"].sum()) if not detail.empty else 0
    # 与看板一致：只看 Tag 的总次数，类型合并展示
    summary = (
        detail.groupby("Tag", as_index=False)
        .agg(次数=("Count", "sum"),
             类型=("Type", lambda s: " + ".join(sorted(set(map(str, s))))),
             描述=("Description", "first"))
        .sort_values(["次数", "Tag"], ascending=[False, True])
    )
    # 占比保留 1 位小数，与页面柱状图上的数字完全一致，方便直接核对
    summary["占比%"] = (summary["次数"] / total * 100).round(1) if total else 0.0
    summary = summary[["Tag", "描述", "类型", "次数", "占比%"]]

    return {"全部原始明细": detail.reset_index(drop=True),
            "Tag汇总": summary.reset_index(drop=True)}


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/data")
def api_data():
    return jsonify(build_dashboard())


@app.route("/api/download/hour")
def download_hour():
    """下载某一个小时（趋势图上点圆点）的原始 Excel：该小时的 操作 + 报警 明细。"""
    key = (request.args.get("key") or "").strip()
    hourly = load_hourly([])
    if not hourly:
        return jsonify({"ok": False, "message": "hourly 目录下没有可用的 Pareto Excel"}), 404
    try:
        slot = datetime.strptime(key, "%Y%m%d%H")
    except ValueError:
        return jsonify({"ok": False, "message": f"key 参数格式应为 YYYYMMDDHH，收到：{key!r}"}), 400
    bucket = hourly.get(slot)
    if bucket is None:
        return jsonify({"ok": False, "message": f"{slot:%Y-%m-%d %H:00} 没有对应的 Pareto Excel"}), 404
    return _excel_response(_detail_sheets(_bucket_frame(bucket)), f"Pareto_Hour_{key}.xlsx")


@app.route("/api/download/day")
def download_day():
    """下载某一天（趋势图上点圆点）的原始 Excel：该天的 操作 + 报警 明细。"""
    key = (request.args.get("key") or "").strip()
    daily = load_daily([])
    if not daily:
        return jsonify({"ok": False, "message": "daily 目录下没有可用的 Pareto Excel"}), 404
    if key:
        try:
            day = datetime.strptime(key, DAY_FMT).date()
        except ValueError:
            return jsonify({"ok": False, "message": f"key 参数格式应为 YYYY-MM-DD，收到：{key!r}"}), 400
    else:
        day = max(daily)
    bucket = daily.get(day)
    if bucket is None:
        return jsonify({"ok": False, "message": f"{day:%Y-%m-%d} 没有对应的 Pareto Excel"}), 404
    return _excel_response(_detail_sheets(_bucket_frame(bucket)), f"Pareto_Day_{day:%Y%m%d}.xlsx")


@app.route("/api/download/window")
def download_window():
    """下载某个窗口内某一类事件的全部原始明细：range=24h|7d，type=Operation|Alarm。"""
    range_key = (request.args.get("range") or "24h").strip().lower()
    type_name = _normalize_type(request.args.get("type"))
    if range_key not in ("24h", "7d"):
        return jsonify({"ok": False, "message": f"range 只支持 24h / 7d，收到：{range_key!r}"}), 400
    if type_name is None:
        return jsonify({"ok": False, "message": "type 只支持 Operation / Alarm"}), 400

    windows = collect_windows(load_hourly([]), load_daily([]))
    if range_key == "24h":
        frame = windows["hour_frame"]
        range_label = windows["hour_range"]
        anchor = windows["hour_anchor"]
        suffix = f"{anchor:%Y%m%d%H}" if anchor else "none"
    else:
        frame = windows["week_frame"]
        range_label = windows["day_range"]
        anchor = windows["day_anchor"]
        suffix = f"{(anchor - timedelta(days=6)):%Y%m%d}_{anchor:%Y%m%d}"

    sub = _slice_type(frame, type_name)
    if sub.empty:
        return jsonify({"ok": False, "message": f"{range_label} 没有{TYPE_LABELS[type_name]}明细"}), 404

    sheets = _detail_sheets(sub, time_column=TIME_COLUMN)
    return _excel_response(sheets, f"Pareto_{range_key.upper()}_{type_name}_{suffix}.xlsx")


_DATA_DIR_REPORTED = False


def report_data_dir() -> str:
    """把"看板正在读哪个目录、里面有多少文件"打出来。

    "看板空白"最常见的原因就是读错目录，而这在页面上完全看不出来：采集端日志里
    全是"已保存"，页面却是空的。打这一行就能当场对上是不是同一处。

    启动时（__main__）和第一个请求进来时都会尝试调用，但只打一次。
    """
    global _DATA_DIR_REPORTED
    _DATA_DIR_REPORTED = True

    hourly = sorted(glob.glob(os.path.join(HOURLY_DIR, "*.xls*")))
    daily = sorted(glob.glob(os.path.join(DAILY_DIR, "*.xls*")))
    lines = [
        "=" * 64,
        f"  看板数据目录 : {describe()}",
        f"  时报目录     : {HOURLY_DIR}   （{len(hourly)} 个文件）",
        f"  日报目录     : {DAILY_DIR}   （{len(daily)} 个文件）",
    ]
    if hourly:
        lines.append(f"  时报最新     : {os.path.basename(hourly[-1])}")
    if daily:
        lines.append(f"  日报最新     : {os.path.basename(daily[-1])}")
    if not hourly and not daily:
        lines += [
            "  ⚠ 这两个目录里一个报表文件都没有，页面会是空的。",
            f"    如果采集端写到了别处，把这个环境变量指过去：{ENV_VAR}=<采集端的目录>",
        ]
    lines.append("=" * 64)
    text = "\n".join(lines)
    print(text, flush=True)
    return text


@app.before_request
def _report_data_dir_on_first_request():
    """用 waitress / gunicorn 之类的服务器启动时不会走 __main__，启动横幅打不出来。

    这里兜底补打一次。已经打过的（__main__ 那条路径）就跳过，不会每刷新一下刷一屏。
    """
    if not _DATA_DIR_REPORTED:
        report_data_dir()


if __name__ == "__main__":
    report_data_dir()

    # 采集线程要在 app.run() 之前起：daemon 线程不阻塞主线程，服务立刻可用，
    # 首轮采集在后台跑（冷启动约 30 秒，网络不通时会更久）。
    if COLLECTOR_ENABLED and collector is not None:
        collector.start_background()

    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5055)), debug=False)
