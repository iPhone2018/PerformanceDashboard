# -*- coding: utf-8 -*-
"""看板自测脚本

分三部分：
  1. 对着 data/ 里的真实数据跑一遍，用 pandas 独立重算后逐项比对接口返回值；
  2. 用 HTTP 拉三个下载接口产出的 Excel，核对行数 / 类型 / 合计 / 与图表是否一致；
  3. 另起服务喂一份"造出来的残缺数据"，验证缺整点、缺一类文件、空文件、缺整天、
     整个目录为空时的表现（这些情况在真实数据里不一定碰得到）。

不依赖 app.py 里的任何函数。

用法：
    ./venv/bin/python selftest.py
"""
from __future__ import annotations

import glob
import io
import json
import os
import re
import shutil
import socket
import subprocess
import sys
import tempfile
import time
import urllib.error
import urllib.request
from datetime import date, datetime, timedelta

import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PORT = int(os.environ.get("SELFTEST_PORT", 0)) or None   # 0/未设置 = 自动挑空闲端口
BASE_URL = ""                                            # 由 main() 决定

PASS, FAIL = [], []


def check(name: str, ok: bool, detail: str = "") -> None:
    (PASS if ok else FAIL).append(name)
    print(f"  {'✓' if ok else '✗'} {name}" + (f"  → {detail}" if detail and not ok else ""))


def section(title: str) -> None:
    print(f"\n=== {title} ===")


# ---------------------------------------------------------------------------
# 独立重算：只认最朴素的表头定位（第 0/1/2/3 列），与 app.py 的实现完全无关
# ---------------------------------------------------------------------------
def read_source(path: str) -> tuple[pd.DataFrame, datetime]:
    raw = pd.read_excel(path, header=None, dtype=object)
    header_row = None
    for i in range(min(len(raw), 40)):
        if (str(raw.iat[i, 0]).strip().lower() == "type"
                and str(raw.iat[i, 1]).strip().lower() == "tag"):
            header_row = i
            break
    assert header_row is not None, f"{path} 找不到表头"

    start = None
    for i in range(min(len(raw), 12)):
        m = re.search(r"Start:\s*(\d{4})/(\d{1,2})/(\d{1,2})\s+(\d{1,2}):(\d{2})", str(raw.iat[i, 1]))
        if m:
            start = datetime(*[int(x) for x in m.groups()])
            break
    assert start is not None, f"{path} 找不到 Start 时间"

    body = raw.iloc[header_row + 1:, :4]
    body.columns = ["Type", "Tag", "Description", "Count"]
    body = body.dropna(subset=["Tag", "Count"])
    body["Type"] = body["Type"].astype(str).str.strip()
    body["Tag"] = body["Tag"].astype(str).str.strip()
    body["Count"] = body["Count"].astype(int)
    body = body[(body["Count"] > 0) & (body["Tag"] != "")]
    return body.reset_index(drop=True), start


def source_hourly() -> dict:
    out: dict = {}
    for path in sorted(glob.glob(os.path.join(BASE_DIR, "data", "hourly", "*.xlsx"))):
        df, start = read_source(path)
        key = start.replace(minute=0, second=0, microsecond=0)
        out[key] = df if key not in out else pd.concat([out[key], df], ignore_index=True)
    return out


def source_daily() -> dict:
    out: dict = {}
    for path in sorted(glob.glob(os.path.join(BASE_DIR, "data", "daily", "*.xlsx"))):
        df, start = read_source(path)
        key = start.date()
        out[key] = df if key not in out else pd.concat([out[key], df], ignore_index=True)
    return out


def sum_type(df: pd.DataFrame, kind: str) -> int:
    if df is None or df.empty:
        return 0
    return int(df.loc[df["Type"].str.lower() == kind.lower(), "Count"].sum())


def top10_of_type(frame: pd.DataFrame, kind: str, total: int) -> list[tuple[str, int, float]]:
    sub = frame[frame["Type"].str.lower() == kind.lower()]
    grouped = (sub.groupby("Tag", as_index=False)["Count"].sum()
               .sort_values(["Count", "Tag"], ascending=[False, True]))
    return [(r.Tag, int(r.Count), round(r.Count / total * 100, 1))
            for r in grouped.head(10).itertuples(index=False)]


# ---------------------------------------------------------------------------
# HTTP / 服务进程
# ---------------------------------------------------------------------------
def http_get(path: str, base: str | None = None) -> tuple[int, dict, bytes]:
    try:
        with urllib.request.urlopen((base or BASE_URL) + path, timeout=60) as resp:
            return resp.status, dict(resp.headers), resp.read()
    except urllib.error.HTTPError as exc:
        return exc.code, dict(exc.headers), exc.read()


def read_book(blob: bytes) -> dict:
    return pd.read_excel(io.BytesIO(blob), sheet_name=None)


def free_port() -> int:
    with socket.socket() as sock:
        sock.bind(("127.0.0.1", 0))
        return int(sock.getsockname()[1])


def stop_server(proc: subprocess.Popen) -> None:
    proc.terminate()
    try:
        proc.wait(timeout=10)
    except subprocess.TimeoutExpired:
        proc.kill()


def start_server(data_dir: str | None, port: int) -> subprocess.Popen:
    expected = os.path.realpath(data_dir or os.path.join(BASE_DIR, "data"))
    env = dict(os.environ, PORT=str(port), PYTHONUNBUFFERED="1", DASHBOARD_DATA_DIR=expected)
    proc = subprocess.Popen([sys.executable, os.path.join(BASE_DIR, "app.py")],
                            cwd=BASE_DIR, env=env,
                            stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    base = f"http://127.0.0.1:{port}"
    deadline = time.time() + 60
    while time.time() < deadline:
        if proc.poll() is not None:
            raise RuntimeError(f"服务进程提前退出，返回码 {proc.returncode}\n"
                               f"{proc.stdout.read().decode(errors='replace')}")
        try:
            urllib.request.urlopen(base + "/api/data", timeout=3).read()
            break
        except Exception:
            time.sleep(0.4)
    else:
        stop_server(proc)
        raise RuntimeError("服务启动超时")

    # 端口若被别的服务（或上次没杀干净的进程）占着，会连到错误的实例上，
    # 后面所有断言都会对着别人的数据跑 —— 这里校验读的到底是不是我们要的目录。
    try:
        _, _, body = http_get("/api/data", base)
        raw_dir = json.loads(body).get("meta", {}).get("data_dir")
    except Exception as exc:  # noqa: BLE001
        stop_server(proc)
        raise RuntimeError(f"{base} 上的服务无法确认数据目录：{exc}") from exc
    served = os.path.realpath(raw_dir) if raw_dir else None
    if served != expected:
        stop_server(proc)
        raise RuntimeError(f"{base} 已被别的服务占用（它读的是 {served}，期望 {expected}）")
    return proc


# ---------------------------------------------------------------------------
# 第一部分：真实数据
# ---------------------------------------------------------------------------
def run_checks() -> None:
    hourly, daily = source_hourly(), source_daily()
    assert hourly and daily, "data/ 下没有数据文件"
    hour_anchor, day_anchor = max(hourly), max(daily)

    status, headers, body = http_get("/api/data")
    data = json.loads(body)

    section("接口 /api/data")
    check("HTTP 200", status == 200, f"实际 {status}")
    check("Content-Type 为 application/json", "application/json" in headers.get("Content-Type", ""))
    check("返回体合法且 ok=true", data.get("ok") is True)
    check("中文未被转义成 \\uXXXX", "\\u" not in body.decode("utf-8")[:2000])
    check("需求1：已去掉 4 个 KPI 数字", "kpi" not in data and "kpi" not in data.get("meta", {}))
    check("接口回报了数据目录（便于确认看的是哪份数据）", bool(data["meta"].get("data_dir")))

    # ---------------- 需求2：24 小时折线 + 点圆点下载 ----------------
    section("需求2：近24小时折线（点圆点下载该小时 Excel）")
    exp_slots = [hour_anchor - timedelta(hours=i) for i in range(23, -1, -1)]
    exp_h_labels = [s.strftime("%m-%d %H:00") for s in exp_slots]
    exp_h_keys = [s.strftime("%Y%m%d%H") for s in exp_slots]
    exp_h_op = [sum_type(hourly.get(s), "Operation") if s in hourly else None for s in exp_slots]
    exp_h_alarm = [sum_type(hourly.get(s), "Alarm") if s in hourly else None for s in exp_slots]

    check("横轴 24 个整点", len(data["hour"]["labels"]) == 24)
    check("横轴标签正确（旧→新）", data["hour"]["labels"] == exp_h_labels)
    check("操作序列正确", data["hour"]["op"] == exp_h_op)
    check("报警序列正确", data["hour"]["alarm"] == exp_h_alarm)
    check("每个整点带下载 key（YYYYMMDDHH）", data["hour"]["keys"] == exp_h_keys)
    check("数据齐全时缺数据列表为空",
          data["meta"]["hour_window"]["missing"] == [] and data["meta"]["hour_window"]["empty"] == [])
    check("数据齐全时覆盖数 = 24", data["meta"]["hour_window"]["covered"] == 24)

    last_key = exp_h_keys[-1]
    status, headers, blob = http_get(f"/api/download/hour?key={last_key}")
    check(f"点最后一个圆点 → 下载 {last_key} 的 Excel：HTTP 200", status == 200, str(status))
    check("下载的是 xlsx", headers.get("Content-Type", "").startswith(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
    check("文件名带小时", f"Pareto_Hour_{last_key}.xlsx" in headers.get("Content-Disposition", ""))
    book = read_book(blob)
    detail = book["全部原始明细"]
    src = hourly[hour_anchor]
    check("该小时 Excel：两张表齐全", set(book) == {"全部原始明细", "Tag汇总"}, str(list(book)))
    check("该小时 Excel：列结构与原始导出一致",
          list(detail.columns) == ["Type", "Tag", "Description", "Count"], str(list(detail.columns)))
    check("该小时 Excel：行数 = 该小时两个文件的行数之和", len(detail) == len(src),
          f"{len(detail)} vs {len(src)}")
    check("该小时 Excel：操作次数对得上",
          int(detail.loc[detail["Type"] == "Operation", "Count"].sum()) == exp_h_op[-1])
    check("该小时 Excel：报警次数对得上",
          int(detail.loc[detail["Type"] == "Alarm", "Count"].sum()) == exp_h_alarm[-1])

    status, _, _ = http_get("/api/download/hour?key=2020010100")
    check("点一个没有数据的小时 → 404 + 提示", status == 404, str(status))
    status, _, _ = http_get("/api/download/hour?key=abc")
    check("key 格式错误 → 400", status == 400, str(status))
    status, _, _ = http_get("/api/download/hour")
    check("key 缺失 → 400", status == 400, str(status))

    # ---------------- 需求2b：7 天折线 + 点圆点下载 ----------------
    section("需求2：近7天折线（点圆点下载该天 Excel）")
    exp_days = [day_anchor - timedelta(days=i) for i in range(6, -1, -1)]
    exp_d_labels = [d.strftime("%Y-%m-%d") for d in exp_days]
    exp_d_op = [sum_type(daily.get(d), "Operation") if d in daily else None for d in exp_days]
    exp_d_alarm = [sum_type(daily.get(d), "Alarm") if d in daily else None for d in exp_days]

    check("横轴 7 天", len(data["day"]["labels"]) == 7)
    check("横轴日期正确（旧→新）", data["day"]["labels"] == exp_d_labels)
    check("操作序列正确", data["day"]["op"] == exp_d_op)
    check("报警序列正确", data["day"]["alarm"] == exp_d_alarm)
    check("每天带下载 key（YYYY-MM-DD）", data["day"]["keys"] == exp_d_labels)
    check("7 天数据齐全", data["meta"]["day_window"]["covered"] == 7
          and data["meta"]["day_window"]["missing"] == [])

    mid_day = exp_days[2]
    status, headers, blob = http_get(f"/api/download/day?key={mid_day:%Y-%m-%d}")
    check(f"点中间某天圆点 → 下载 {mid_day} 的 Excel：HTTP 200", status == 200, str(status))
    check("文件名带日期", f"Pareto_Day_{mid_day:%Y%m%d}.xlsx" in headers.get("Content-Disposition", ""))
    detail = read_book(blob)["全部原始明细"]
    check("该天 Excel：行数 = 该天两个文件的行数之和", len(detail) == len(daily[mid_day]))
    check("该天 Excel：操作次数对得上",
          int(detail.loc[detail["Type"] == "Operation", "Count"].sum()) == exp_d_op[2])
    status, _, _ = http_get("/api/download/day?key=2000-01-01")
    check("没有数据的那天 → 404", status == 404, str(status))
    status, _, _ = http_get("/api/download/day?key=2026/09/29")
    check("日期格式错误 → 400", status == 400, str(status))

    # ---------------- 需求3+5：四张 TOP10 ----------------
    section("需求3+5：近24小时 TOP10 / 近7天 TOP10，操作与报警分开，共 4 张")
    hour_frame = pd.concat(
        [hourly[s].assign(时间=s.strftime("%Y-%m-%d %H:00")) for s in exp_slots if s in hourly],
        ignore_index=True)
    week_frame = pd.concat(
        [daily[d].assign(时间=d.strftime("%Y-%m-%d")) for d in exp_days if d in daily],
        ignore_index=True)

    charts = data.get("charts", [])
    check("共 4 张柱状图", len(charts) == 4, str(len(charts)))
    check("2×2 排布顺序：24h操作 / 24h报警 / 7天操作 / 7天报警",
          [(c["window"], c["type"]) for c in charts] ==
          [("24h", "Operation"), ("24h", "Alarm"), ("7d", "Operation"), ("7d", "Alarm")])

    for chart in charts:
        kind = chart["type"]
        frame = hour_frame if chart["window"] == "24h" else week_frame
        window_label = "近 24 小时" if chart["window"] == "24h" else "近 7 天"
        type_label = "操作" if kind == "Operation" else "报警"
        sub = frame[frame["Type"].str.lower() == kind.lower()]
        total = int(sub["Count"].sum())
        expected = top10_of_type(frame, kind, total)
        got = [(i["tag"], i["count"], i["pct"]) for i in chart["items"]]

        check(f"[{window_label}·{type_label}] 标题正确",
              chart["title"] == f"{window_label} TOP10 Tag · {type_label}", chart["title"])
        check(f"[{window_label}·{type_label}] 总量 = {total}", chart["total"] == total, str(chart["total"]))
        check(f"[{window_label}·{type_label}] TOP10 的 Tag/次数/占比 全部正确", got == expected,
              f"\n      接口 {got[:3]}\n      重算 {expected[:3]}")
        check(f"[{window_label}·{type_label}] 最多 10 条且按次数降序",
              len(got) <= 10 and all(got[i][1] >= got[i + 1][1] for i in range(len(got) - 1)))
        check(f"[{window_label}·{type_label}] 占比分母是同类事件总数",
              all(abs(c / total * 100 - p) < 0.051 for _, c, p in got))
        check(f"[{window_label}·{type_label}] 只统计本类型事件",
              all(i["type"] == kind and i["type_label"] == type_label for i in chart["items"]))
        check(f"[{window_label}·{type_label}] 带了下载入口",
              chart["download_url"].endswith(f"type={kind}"))
        check(f"[{window_label}·{type_label}] 数据齐全时无缺失提示", chart["data_note"] == "")

    frame_of = {"24h": hour_frame, "7d": week_frame}
    check("每张图的 Tag 都取自本类型明细（没有跨类型混入）",
          all(set(i["tag"] for i in c["items"]) <=
              set(frame_of[c["window"]].loc[
                      frame_of[c["window"]]["Type"].str.lower() == c["type"].lower(), "Tag"])
              for c in charts))

    op24 = {i["tag"]: i["count"] for i in charts[0]["items"]}
    al24 = {i["tag"]: i["count"] for i in charts[1]["items"]}
    shared = sorted(set(op24) & set(al24))
    if shared:
        tag = shared[0]
        both = int(hour_frame.loc[hour_frame["Tag"] == tag, "Count"].sum())
        check(f"同一个 Tag（{tag}）同时有操作和报警时分开统计、不合并",
              op24[tag] + al24[tag] == both, f"{op24[tag]} + {al24[tag]} vs {both}")
    else:
        check("24 小时窗口内操作/报警无同名 Tag（分开统计无歧义）", True)

    # ---------------- 分类型下载 ----------------
    section("需求5：按类型下载全部原始明细")
    for window, frame, kind in (("24h", hour_frame, "Operation"), ("24h", hour_frame, "Alarm"),
                                ("7d", week_frame, "Operation"), ("7d", week_frame, "Alarm")):
        sub = frame[frame["Type"].str.lower() == kind.lower()]
        status, headers, blob = http_get(f"/api/download/window?range={window}&type={kind}")
        label = f"{window}·{kind}"
        if status != 200:
            check(f"[{label}] 下载 200", False, str(status))
            continue
        book = read_book(blob)
        detail = book["全部原始明细"]
        check(f"[{label}] 下载 200 且文件可打开", True)
        check(f"[{label}] 类型列只有 {kind}", set(detail["Type"]) == {kind}, str(set(detail["Type"])))
        check(f"[{label}] 行数 = 该类事件全部明细行数", len(detail) == len(sub),
              f"{len(detail)} vs {len(sub)}")
        check(f"[{label}] 次数合计 = 该窗口该类总数",
              int(detail["Count"].sum()) == int(sub["Count"].sum()))
        check(f"[{label}] 明细带时间列且在最前", list(detail.columns)[0] == "时间")
        sheet_top = [(r["Tag"], int(r["次数"]), float(r["占比%"]))
                     for _, r in book["Tag汇总"].head(10).iterrows()]
        chart = next(c for c in charts if c["window"] == window and c["type"] == kind)
        check(f"[{label}] Excel 的 Tag 汇总与柱状图逐字一致（含占比小数位）",
              sheet_top == [(i["tag"], i["count"], i["pct"]) for i in chart["items"]],
              f"\n      Excel {sheet_top[:3]}\n      图表  {[(i['tag'], i['count'], i['pct']) for i in chart['items']][:3]}")

    status, _, _ = http_get("/api/download/window?range=24h&type=op")
    check("type 支持 op 简写", status == 200, str(status))
    status, _, _ = http_get("/api/download/window?range=24h&type=alm")
    check("type 支持 alm 简写", status == 200, str(status))
    status, _, _ = http_get("/api/download/window?range=30d&type=Operation")
    check("range 非法 → 400", status == 400, str(status))
    status, _, _ = http_get("/api/download/window?range=24h&type=Whatever")
    check("type 非法 → 400", status == 400, str(status))

    # ---------------- 需求1 / 4 / 6：页面 ----------------
    section("需求1/4/6：页面结构")
    status, _, page = http_get("/")
    html = page.decode("utf-8")
    check("页面 HTTP 200", status == 200)
    check("需求1：页面没有 KPI 卡片", "kpi-week-op" not in html and "kpi-hour-op" not in html
          and "kpi-row" not in html and "近 7 天总操作" not in html)
    check("需求4：页面没有表格视图", "table-day-top" not in html and "表格" not in html
          and "tableHtml" not in html)
    check("2 个折线容器", 'id="chart-hour"' in html and 'id="chart-day"' in html)
    check("4 个柱状容器", all(f'id="chart-top-{i}"' in html for i in range(4)))
    check("4 个下载按钮占位", all(f'id="top-dl-{i}"' in html for i in range(4)))
    check("折线图点击 → 下载该小时/该天",
          "/api/download/hour?key=" in html and "/api/download/day?key=" in html
          and "triggerDownload" in html)
    check("折线图绑定了 click 事件", html.count("Chart.on('click'") >= 2)
    check("需求6：柱子点击弹窗仍在", all(k in html for k in
          ("modal-tag", "modal-type", "modal-count", "modal-pct", "showModal")))
    check("缺数据时折线断开而不是连过去", "connectNulls: false" in html)
    check("缺数据时有灰带标记", "markArea" in html and "missingRanges" in html)
    check("每个面板都有数据完整性提示位", all(
        f'id="{i}"' in html for i in
        ("quality-hour", "quality-day", "top-quality-0", "top-quality-1", "top-quality-2", "top-quality-3")))
    check("下载有提示（toast）", 'id="toast"' in html and "toast(" in html)
    check("页面无 U+2011 等异形连字符", "‑" not in html and "‐" not in html and "­" not in html)
    check("ECharts 走本地 static（离线可用）", '/static/echarts.min.js' in html)

    status, _, js = http_get("/static/echarts.min.js")
    check("静态 echarts 可访问", status == 200 and len(js) > 100000, f"{status} / {len(js)} bytes")
    status, _, logo = http_get("/static/valmet_logo.png")
    check("静态 logo 可访问", status == 200 and logo[:4] == b"\x89PNG", str(status))


# ---------------------------------------------------------------------------
# 第二部分：数据缺失场景（喂造出来的残缺数据）
# ---------------------------------------------------------------------------
def write_pareto(path: str, kind: str, start: datetime, rows: list) -> None:
    """按真实 Pareto 版式写文件（rows 为空 = 只有一个表头的空导出）。"""
    data = [["", "Pareto Report"],
            ["", f"Start:  {start.strftime('%Y/%m/%d %H:%M:%S')}"],
            ["", f"End:    {(start + timedelta(hours=1)).strftime('%Y/%m/%d %H:%M:%S')}"],
            ["", "Print time:  2026/03/10 12:05"],
            ["", f"Rows retrieved:  {len(rows)}"],
            ["", ""],
            ["Type", "Tag", "Description", "Count"]]
    for tag, count in rows:
        data.append([kind, tag, f"desc {tag}", count])
    pd.DataFrame(data).to_excel(path, index=False, header=False)


def run_missing_data_checks() -> None:
    section("数据缺失：缺整点 / 只缺一类文件 / 空文件 / 缺整天 / 目录为空")

    root = tempfile.mkdtemp(prefix="dash-missing-")
    data_dir = os.path.join(root, "data")
    hourly_dir, daily_dir = os.path.join(data_dir, "hourly"), os.path.join(data_dir, "daily")
    os.makedirs(hourly_dir)
    os.makedirs(daily_dir)

    anchor = datetime(2026, 3, 10, 12, 0)
    slots = [anchor - timedelta(hours=i) for i in range(23, -1, -1)]      # 旧 → 新
    no_file = [slots[4], slots[5], slots[16]]       # 两个连续 + 一个孤立
    no_alarm_file = slots[9]                        # 只有 Operation 文件
    empty_file = slots[20]                          # 文件在，但一条明细都没有

    for slot in slots:
        if slot in no_file:
            continue
        if slot == no_alarm_file:
            write_pareto(os.path.join(hourly_dir, f"Operation_{slot:%Y%m%d_%H%M}.xlsx"),
                         "Operation", slot, [("OP-A", 5)])
            continue
        write_pareto(os.path.join(hourly_dir, f"Alarm_{slot:%Y%m%d_%H%M}.xlsx"), "Alarm", slot,
                     [] if slot == empty_file else [("AL-A", 10), ("AL-B", 2)])
        write_pareto(os.path.join(hourly_dir, f"Operation_{slot:%Y%m%d_%H%M}.xlsx"), "Operation", slot,
                     [] if slot == empty_file else [("OP-A", 7)])

    day_anchor = date(2026, 3, 10)
    days = [day_anchor - timedelta(days=i) for i in range(6, -1, -1)]
    day_no_file = [days[1], days[3]]
    day_empty = days[5]
    for d in days:
        if d in day_no_file:
            continue
        start = datetime.combine(d, datetime.min.time())
        write_pareto(os.path.join(daily_dir, f"Alarm_{d:%Y%m%d}.xlsx"), "Alarm", start,
                     [] if d == day_empty else [("AL-A", 100)])
        write_pareto(os.path.join(daily_dir, f"Operation_{d:%Y%m%d}.xlsx"), "Operation", start,
                     [] if d == day_empty else [("OP-A", 50)])

    port = free_port()
    base = f"http://127.0.0.1:{port}"
    proc = start_server(data_dir, port)
    try:
        _, _, body = http_get("/api/data", base)
        d = json.loads(body)
        hour, meta = d["hour"], d["meta"]["hour_window"]
        labels = [s.strftime("%m-%d %H:00") for s in slots]

        exp_op = [None if s in no_file else (5 if s == no_alarm_file else (0 if s == empty_file else 7))
                  for s in slots]
        exp_alarm = [None if (s in no_file or s == no_alarm_file) else (0 if s == empty_file else 12)
                     for s in slots]

        check("缺文件的整点：序列为 null（折线断开，而不是画成 0）", hour["op"] == exp_op, str(hour["op"]))
        check("只缺报警文件的整点：报警断线，但操作仍有数", hour["alarm"] == exp_alarm, str(hour["alarm"]))
        check("文件存在但 0 条记录：按真实的 0 处理，不是 null",
              hour["op"][slots.index(empty_file)] == 0 and hour["alarm"][slots.index(empty_file)] == 0)
        check("missing 只列完全没有文件的整点",
              meta["missing"] == [labels[slots.index(s)] for s in no_file], str(meta["missing"]))
        check("只缺一类文件的整点不算 missing（另一类还有数据）",
              labels[slots.index(no_alarm_file)] not in meta["missing"])
        check("空文件单独列在 empty 里",
              meta["empty"] == [labels[slots.index(empty_file)]], str(meta["empty"]))
        check("覆盖率 = 有文件的整点个数（24-3）", meta["covered"] == 21, str(meta["covered"]))
        check("提示文案写明缺几个整点并列出时间",
              "缺 3 个整点" in meta["quality"] and labels[slots.index(no_file[0])] in meta["quality"],
              meta["quality"])
        check("提示文案区分开『文件里没有明细记录』", "没有明细记录" in meta["quality"], meta["quality"])
        check("柱状图口径里带覆盖率", "数据覆盖 21/24 个整点" in d["charts"][0]["note"],
              d["charts"][0]["note"])
        check("柱状图单独给出缺失提示", "缺 3 个整点" in d["charts"][0]["data_note"])
        # 有操作文件的整点 = 24 - 3 个完全缺失 - 1 个空文件 = 20 个；其中 1 个只写了 5
        check("缺失的整点没被算进操作总量（19×7 + 5 = 138）",
              d["charts"][0]["total"] == 19 * 7 + 5, str(d["charts"][0]["total"]))
        check("缺失的整点没被算进报警总量（19×12）",
              d["charts"][1]["total"] == 19 * 12, str(d["charts"][1]["total"]))

        day, day_meta = d["day"], d["meta"]["day_window"]
        day_labels = [x.strftime("%Y-%m-%d") for x in days]
        check("缺天：序列为 null",
              day["op"] == [None if x in day_no_file else (0 if x == day_empty else 50) for x in days],
              str(day["op"]))
        check("缺天：missing 列表正确",
              day_meta["missing"] == [day_labels[days.index(x)] for x in day_no_file])
        check("缺天：空文件单独标记", day_meta["empty"] == [day_labels[days.index(day_empty)]])
        check("缺天：7 天趋势仍有 7 个位置（不塌缩）", len(day["labels"]) == 7)
        check("缺天：提示文案正确", "缺 2 个天" in day_meta["quality"], day_meta["quality"])

        key_missing = f"{no_file[0]:%Y%m%d%H}"
        status, _, _ = http_get(f"/api/download/hour?key={key_missing}", base)
        check("缺失整点的下载接口 → 404 而不是空文件", status == 404, str(status))
        status, _, blob = http_get(f"/api/download/hour?key={empty_file:%Y%m%d%H}", base)
        check("空文件那个小时仍可下载（文件确实存在）", status == 200, str(status))
        if status == 200:
            rows = read_book(blob)["全部原始明细"]
            check("空文件的 Excel 里 0 行明细", len(rows) == 0, str(len(rows)))
        status, _, blob = http_get("/api/download/window?range=24h&type=Operation", base)
        detail = read_book(blob)["全部原始明细"]
        check("按类型下载只含真实存在的数据",
              int(detail["Count"].sum()) == d["charts"][0]["total"],
              f'{int(detail["Count"].sum())} vs {d["charts"][0]["total"]}')
        status, _, _ = http_get(f"/api/download/day?key={day_no_file[0]:%Y-%m-%d}", base)
        check("缺的那天下载 → 404", status == 404, str(status))
    finally:
        stop_server(proc)
        shutil.rmtree(root, ignore_errors=True)

    # ---- 目录整个为空 ----
    empty_root = tempfile.mkdtemp(prefix="dash-empty-")
    os.makedirs(os.path.join(empty_root, "data", "hourly"))
    os.makedirs(os.path.join(empty_root, "data", "daily"))
    port = free_port()
    proc = start_server(os.path.join(empty_root, "data"), port)
    try:
        _, _, body = http_get("/api/data", f"http://127.0.0.1:{port}")
        d = json.loads(body)
        check("目录全空：ok=false 且给出中文提示",
              d["ok"] is False and "没有可用" in d["message"], json.dumps(d, ensure_ascii=False))
    finally:
        stop_server(proc)
        shutil.rmtree(empty_root, ignore_errors=True)

    # ---- 只有天文件、没有任何小时文件 ----
    half_root = tempfile.mkdtemp(prefix="dash-half-")
    os.makedirs(os.path.join(half_root, "data", "hourly"))
    os.makedirs(os.path.join(half_root, "data", "daily"))
    write_pareto(os.path.join(half_root, "data", "daily", "Alarm_20260310.xlsx"),
                 "Alarm", datetime(2026, 3, 10), [("AL-A", 9)])
    port = free_port()
    base = f"http://127.0.0.1:{port}"
    proc = start_server(os.path.join(half_root, "data"), port)
    try:
        _, _, body = http_get("/api/data", base)
        d = json.loads(body)
        check("没有任何小时文件：仍然 ok=true（只有 24 小时那侧空）", d["ok"] is True)
        check("没有任何小时文件：提示条说明原因",
              "data/hourly" in d["meta"]["hour_window"]["quality"], d["meta"]["hour_window"]["quality"])
        check("没有任何小时文件：24 小时趋势为空数组而不是报错", d["hour"]["labels"] == [])
        check("没有任何小时文件：口径说明可读（不出现 0/0）",
              "未找到小时文件" in d["charts"][0]["note"], d["charts"][0]["note"])
        check("天数据照常出数", d["charts"][3]["total"] == 9, str(d["charts"][3]["total"]))
    finally:
        stop_server(proc)
        shutil.rmtree(half_root, ignore_errors=True)


# ---------------------------------------------------------------------------
def main() -> int:
    global BASE_URL
    port = PORT or free_port()
    BASE_URL = f"http://127.0.0.1:{port}"
    print(f"启动服务 {BASE_URL} …")
    proc = start_server(None, port)
    try:
        run_checks()
        run_missing_data_checks()
    finally:
        stop_server(proc)

    print("\n" + "=" * 62)
    print(f"通过 {len(PASS)} 项，失败 {len(FAIL)} 项")
    if FAIL:
        print("失败项：")
        for name in FAIL:
            print("  -", name)
        return 1
    print("全部通过 ✅")
    return 0


if __name__ == "__main__":
    sys.exit(main())
