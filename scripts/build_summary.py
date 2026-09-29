# -*- coding: utf-8 -*-
"""策略最佳化參數彙整 — 產出腳本（重新產生 策略最佳化參數彙整_YYYYMMDD.xlsx）

用法
    python scripts\\build_summary.py                     # 產生到專案根目錄，檔名帶今天日期
    python scripts\\build_summary.py --out D:\\tmp\\a.xlsx # 指定輸出
    python scripts\\build_summary.py --source "D:\\..."   # 指定來源工作區

採集口徑（不可混用）
    - 主表「最佳」＝所有輪次中「淨利最高」的一輪全期績效。
    - 「R3 WFO 推薦參數」為前向最佳化各視窗的推薦值，僅供參考，與主表數值分欄。
    - 月報酬率＝報酬率(RoA%) ÷ 資料月數；資料月數由 WFO 視窗最早起點～最晚終點推估
      （每月 30.4375 天＝365.25/12），非原始報告記載值。

截圖
    把 MultiCharts「策略回測績效報告」的平倉權益曲線截圖存成
    screenshots\\<策略名稱>.png，重跑本腳本即自動嵌入「平倉權益曲線圖」分頁。
"""
from __future__ import annotations

import argparse
import os
import re
import sys
from datetime import date, datetime

import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import curated_notes as CN  # noqa: E402  人工策展註記（非來源可推導）
import compare_wsp as CW    # noqa: E402  .wsp 程式參數一致性比對

# ---------------------------------------------------------------------------
# 設定
# ---------------------------------------------------------------------------
PROJECT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCREENSHOT_DIR = os.path.join(PROJECT_DIR, "screenshots")
DEFAULT_SOURCE = r"H:\我的雲端硬碟\Muti-Agents 0924"
OLD_SOURCE = r"G:\我的雲端硬碟\Muti-Agents 0924"   # 舊成果檔記載的來源路徑
MC_WSP_DIR = r"E:\Jobe\群益交易程式"          # 最佳化原始工作區（截圖與參數比對用）

# 來源分頁命名不統一，以正規式對應
RE_DETAIL = re.compile(r"^R(\d+)(?:_最佳化明細|_全域粗掃|_精修|\s*明細)")
RE_WFO_R3 = re.compile(r"^R3[_ ]WFO(明細)?$")
RE_WFO_N = re.compile(r"^(WFO\d+)明細")

# 輪次 → 方法（人工對照表）
ROUND_METHOD = {
    "原單輪": "單次/定期最佳化（原單輪）",
    "R1": "R1 寬範圍篩選（基因演算法）",
    "R2": "R2 收斂精算",
    "R3": "R3 訊號精修",
    "R4": "R4 邊界延伸",
    "R5": "R5 平台驗證",
}

# 「所有參數設定值」要排除的欄位（比對來源原始欄名，區分大小寫）。
# 註：Avg Trade / Avg Winning Trade / Avg Losing Trade / Win-Loss Ratio /
#     Max Consecutive * / Avg Bars in * 僅列帶空格版本，故來源改用 camelCase
#     命名的策略（策略5 肯特納）會把這 8 個衍生指標一併帶出來。此為沿用舊腳本的
#     既有行為，刻意保留以維持既有成果檔逐格一致。
EXCLUDE_COLS = {
    "Rank",
    "Net Profit", "NetProfit",
    "Gross Profit", "GrossProfit",
    "Gross Loss", "GrossLoss",
    "Total Trades", "TotalTrades",
    "% Profitable", "%Profitable",
    "Winning Trades", "WinningTrades",
    "Losing Trades", "LosingTrades",
    "Avg Trade", "Avg Winning Trade", "Avg Losing Trade",
    "Win/Loss Ratio", "Max Consecutive Winners", "Max Consecutive Losers",
    "Avg Bars in Winner", "Avg Bars in Loser",
    "Max Intraday Drawdown", "MaxIntradayDrawDown",
    "Profit Factor", "ProfitFactor",
    "Return on Account", "ReturnOnAccount",
    "Custom Fitness Value", "CustomFitness",
}

# 欄位取值以正規化欄名比對（去掉非英數與 % 後轉小寫），以涵蓋
# 「Net Profit」/「NetProfit」兩套命名。
F_NET = "netprofit"
F_PF = "profitfactor"
F_TRADES = "totaltrades"
F_PCT = "%profitable"
F_DD = "maxintradaydrawdown"
F_ROA = "returnonaccount"
F_AVG = "avgtrade"

# 每月平均天數（365.25/12），用於資料月數推估
DAYS_PER_MONTH = 365.25 / 12.0

# ---------------------------------------------------------------------------
# 樣式
# ---------------------------------------------------------------------------
C_HEADER_FILL = PatternFill("solid", fgColor="FF1F3864")
C_PICK_FILL = PatternFill("solid", fgColor="FFD9E2F3")   # 全部輪次明細：★ 採用列
C_WARN_FILL = PatternFill("solid", fgColor="FFFFF2CC")    # 彙整主表：重大警訊策略
C_NEG_FILL = PatternFill("solid", fgColor="FFFFC7CE")     # 逐窗明細：OOS 為負
C_NOTRADE_FILL = PatternFill("solid", fgColor="FFF2F2F2")  # 逐窗明細：無交易收尾窗
C_GRAY = "FF808080"
C_DARKGRAY = "FF595959"
C_RED = "FFC00000"

F_TITLE = Font(bold=True, size=14)
F_NOTE_GRAY = Font(size=9, color=C_DARKGRAY)
F_NOTE_RED = Font(size=9, color=C_RED)
F_HEADER = Font(bold=True, size=10, color="FFFFFFFF")
F_BODY = Font(size=11)
F_BODY_BOLD = Font(bold=True, size=11)
F_BODY_GRAY = Font(size=11, color=C_GRAY)

A_HEADER = Alignment(horizontal="center", vertical="center", wrap_text=True)
A_HEADER_PLAIN = Alignment(horizontal="center", vertical="center")
A_BODY_TOP = Alignment(vertical="top")
A_BODY_TOP_WRAP = Alignment(vertical="top", wrap_text=True)

# 分頁版面
SHEET_SPECS = {
    "彙整主表": {
        "widths": {"A": 6, "B": 20.125, "C": 20.125, "D": 24, "E": 12, "F": 12, "G": 10,
                   "H": 10, "I": 9, "J": 13, "K": 12, "L": 9, "M": 12, "N": 62, "O": 52,
                   "P": 11, "Q": 12, "R": 22, "S": 70, "T": 20},
        "freeze": "C5",
        "wrap_cols": {"B", "C", "N", "O", "S"},
        "numfmt": {"F": "#,##0", "G": "0.00", "I": '0.00"%"', "J": "#,##0",
                   "K": '0.00"%"', "M": "0.00"},
    },
    "全部輪次明細": {
        "widths": {"A": 6, "B": 28, "C": 8, "D": 26, "E": 11, "F": 12, "G": 10,
                   "H": 10, "I": 9, "J": 13, "K": 12, "L": 12, "M": 62, "N": 20,
                   "O": 52, "P": 16},
        "freeze": "C4",
        "wrap_cols": {"M"},
        "numfmt": {"F": "#,##0", "G": "0.000", "I": '0.00"%"', "J": "#,##0",
                   "K": '0.00"%"', "L": "#,##0"},
    },
    "R3_WFO逐窗明細": {
        "widths": {"A": 6, "B": 28, "C": 10, "D": 7, "E": 12, "F": 14, "G": 12,
                   "H": 12, "I": 12, "J": 10, "K": 10, "L": 16, "M": 50},
        "freeze": "D4",
        "wrap_cols": {"M"},
        "numfmt": {},
    },
    "wsp參數比對": {
        "widths": {"A": 6, "B": 26, "C": 9, "D": 16, "E": 14, "F": 16, "G": 11, "H": 60},
        "freeze": "A4",
        "wrap_cols": {"H"},
        "numfmt": {},
    },
    "平倉權益曲線圖": {
        "widths": {"A": 30, "B": 20},
        "freeze": None,
        "wrap_cols": set(),
        "numfmt": {},
    },
    "資料來源與限制": {
        "widths": {"A": 26, "B": 120},
        "freeze": None,
        "wrap_cols": {"A", "B"},
        "numfmt": {},
    },
}

MAIN_HEADERS = [
    "編號", "策略名稱", "商品/週期", "資料區間(推估)", "最高淨利輪次", "淨利", "獲利因子",
    "總交易次數", "勝率%", "最大日內回撤", "報酬率 RoA%", "資料月數", "月報酬率%(估)",
    "所有參數設定值", "R3 WFO推薦參數(參考)", "OOS正/負窗", "OOS淨利總和",
    "Multichart交易程式參數是否一致", "特殊備註", "平倉權益曲線圖",
]
DETAIL_HEADERS = [
    "編號", "策略名稱", "輪次", "方法", "候選組合數", "淨利", "獲利因子", "總交易次數",
    "勝率%", "最大日內回撤", "報酬率 RoA%", "平均每筆淨利", "參數設定值",
    "資料來源分頁", "來源檔案", "★採為最高淨利輪",
]
WFO_HEADERS = [
    "編號", "策略名稱", "WFO", "窗口", "視窗起", "IS結束/OOS起", "視窗迄", "OOS淨利",
    "OOS獲利因子", "OOS交易數", "OOS勝率%", "OOS最大日內回撤", "該窗最佳參數",
]
WSP_HEADERS = [
    "編號", "策略名稱", "視窗", "input", "報告值", "wsp值", "判定", "說明",
]


# ---------------------------------------------------------------------------
# 小工具
# ---------------------------------------------------------------------------
def norm(text):
    """欄名正規化：去掉非英數與 % 後轉小寫。"""
    return re.sub(r"[^0-9a-zA-Z%]", "", str(text)).lower()


def rnd(value, digits=0):
    """四捨五入後，若為整數則轉 int，避免 68.0 這種尾數。"""
    if value is None:
        return None
    result = round(float(value), digits)
    if result == int(result):
        return int(result)
    return result


def fmt_param(value):
    """參數值的顯示格式：保留來源型別（int 不帶小數點、float 最多 6 位小數）。"""
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        return str(round(value, 6))
    return str(value)


def to_date(value):
    """把來源日期（datetime / date / 字串）轉成 date，無法解析時回 None。"""
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    for pattern in ("%Y-%m-%d", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S"):
        try:
            return datetime.strptime(text, pattern).date()
        except ValueError:
            continue
    return None


def date_text(value):
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")
    return str(value).strip().split(" ")[0]


def cell(row, cmap, field):
    idx = cmap.get(field)
    if idx is None or idx >= len(row):
        return None
    return row[idx]


def find_header(rows, field):
    """找出含指定正規化欄名的那一列（欄名列位置）。"""
    for i, row in enumerate(rows):
        for value in row:
            if isinstance(value, str) and norm(value) == field:
                return i
    return None


def build_cmap(header):
    return {norm(v): i for i, v in enumerate(header) if isinstance(v, str)}


# ---------------------------------------------------------------------------
# 來源解析
# ---------------------------------------------------------------------------
def read_detail_sheet(ws):
    """解析最佳化明細分頁：回傳該分頁淨利最高的一列。"""
    rows = list(ws.iter_rows(values_only=True))
    hi = find_header(rows, F_NET)
    if hi is None:
        return None
    header = rows[hi]
    cmap = build_cmap(header)
    net_idx = cmap[F_NET]
    data = [r for r in rows[hi + 1:]
            if net_idx < len(r) and isinstance(r[net_idx], (int, float))
            and not isinstance(r[net_idx], bool)]
    if not data:
        return None
    best = max(data, key=lambda r: r[net_idx])

    params = []
    for idx, name in enumerate(header):
        if not isinstance(name, str) or name in EXCLUDE_COLS:
            continue
        if idx >= len(best) or best[idx] is None:
            continue
        params.append("%s=%s" % (name, fmt_param(best[idx])))

    return {
        "candidates": len(data),
        "net_profit": rnd(cell(best, cmap, F_NET), 0),
        "profit_factor": rnd(cell(best, cmap, F_PF), 3),
        "total_trades": rnd(cell(best, cmap, F_TRADES), 0),
        "pct_profitable": rnd(cell(best, cmap, F_PCT), 2),
        "max_dd": rnd(cell(best, cmap, F_DD), 0),
        "roa": rnd(cell(best, cmap, F_ROA), 2),
        "avg_trade": rnd(cell(best, cmap, F_AVG), 0),
        "params": " / ".join(params),
    }


def read_wfo_sheet(ws, label):
    """解析前向最佳化（WFO）逐窗分頁。"""
    rows = list(ws.iter_rows(values_only=True))
    hi = find_header(rows, "windowstart")
    if hi is None:
        return None
    header = rows[hi]
    cmap = build_cmap(header)

    param_idx = [i for i, name in enumerate(header)
                 if isinstance(name, str) and name.startswith("Param_")]

    out = []
    for row in rows[hi + 1:]:
        window = cell(row, cmap, "window")
        if not isinstance(window, (int, float)) or isinstance(window, bool):
            continue
        parts = []
        for idx in param_idx:
            if idx < len(row) and row[idx] is not None:
                parts.append("%s=%s" % (header[idx][len("Param_"):], fmt_param(row[idx])))
        out.append({
            "label": label,
            "window": rnd(window, 0),
            "start": to_date(cell(row, cmap, "windowstart")),
            "is_end": to_date(cell(row, cmap, "isendoosstart1")),
            "end": to_date(cell(row, cmap, "windowend")),
            "net": rnd(cell(row, cmap, "oosnetprofit"), 0) or 0,
            "pf": rnd(cell(row, cmap, "oosprofitfactor"), 3) or 0,
            "trades": rnd(cell(row, cmap, "oostotaltrades"), 0) or 0,
            "pct": rnd(cell(row, cmap, "oos%profitable"), 2) or 0,
            "dd": rnd(cell(row, cmap, "oosmaxintradaydrawdown"), 0) or 0,
            "params": " / ".join(parts),
        })
    return out or None


def read_range_note(ws):
    """從「摘要」分頁抓「資料範圍」文字（僅供無 WFO 視窗的策略推估區間用）。"""
    for row in ws.iter_rows(values_only=True):
        if not row:
            continue
        key = str(row[0] or "").strip()
        if key in ("資料範圍", "資料區間") and len(row) > 1:
            return str(row[1] or "")
    return ""


def collect(source_dir):
    """掃描來源工作區所有報告 .xlsx，回傳輪次、WFO 視窗與策略索引。"""
    files = sorted(f for f in os.listdir(source_dir)
                   if f.lower().endswith(".xlsx") and not f.startswith("~$"))
    rounds, wfos = [], []
    strategies = {}

    for fname in files:
        m = re.match(r"^(\d+)-", fname)
        if not m:
            continue
        sid = int(m.group(1))
        name = fname.split("_")[0]
        is_wfo_file = "_WFO_" in fname
        info = strategies.setdefault(sid, {"id": sid, "name": name, "range_note": ""})

        wb = openpyxl.load_workbook(os.path.join(source_dir, fname),
                                    data_only=True, read_only=True)
        try:
            for order, ws in enumerate(wb.worksheets):
                title = ws.title
                if title == "摘要":
                    note = read_range_note(ws)
                    if note and not info["range_note"]:
                        info["range_note"] = note
                    continue

                m_detail = RE_DETAIL.match(title)
                if m_detail:
                    rid = int(m_detail.group(1))
                    label = "R%d" % rid
                    if rid == 1 and not is_wfo_file:
                        label = "原單輪"      # 非 WFO 報告的 R1 分頁＝原始單輪最佳化
                    data = read_detail_sheet(ws)
                    if data:
                        rounds.append(dict(data, sid=sid, name=name, round=label,
                                           method=ROUND_METHOD.get(label, label),
                                           sheet=title, file=fname,
                                           order=(sid, 1 if is_wfo_file else 0, order)))
                    continue

                if RE_WFO_R3.match(title):
                    wfo_label, family = title, "r3"
                else:
                    m_wfo = RE_WFO_N.match(title)
                    if not m_wfo:
                        continue
                    wfo_label, family = m_wfo.group(1), "wn"
                data = read_wfo_sheet(ws, wfo_label)
                if data:
                    wfos.extend(dict(w, sid=sid, name=name, family=family,
                                     sheet=title, file=fname) for w in data)
        finally:
            wb.close()

    rounds.sort(key=lambda r: r["order"])
    wfos.sort(key=lambda w: (0 if w["family"] == "r3" else 1, str(w["sid"]),
                             w["label"], w["window"]))
    return files, rounds, wfos, strategies


# ---------------------------------------------------------------------------
# 彙整主表欄位
# ---------------------------------------------------------------------------
def parse_range_from_text(text):
    """從『資料範圍』文字取出起訖兩組日期，例：2025/05/05 ～ 2026/09/25。"""
    found = re.findall(r"(\d{4})\s*[/年\-]\s*(\d{1,2})\s*[/月\-]\s*(\d{1,2})", text or "")
    if len(found) < 2:
        return None
    def iso(m):
        return date(int(m[0]), int(m[1]), int(m[2]))
    return iso(found[0]), iso(found[1])


def build_main_rows(files, rounds, wfos, strategies):
    """組出彙整主表一列一策略。"""
    by_sid_rounds = {}
    for r in rounds:
        by_sid_rounds.setdefault(r["sid"], []).append(r)
    by_sid_wfos = {}
    for w in wfos:
        by_sid_wfos.setdefault(w["sid"], []).append(w)

    picked = {}
    rows = []
    for sid in sorted(strategies):
        info = strategies[sid]
        sid_rounds = by_sid_rounds.get(sid, [])
        best = max(sid_rounds, key=lambda r: r["net_profit"])
        picked[sid] = (best["file"], best["sheet"])

        pick = CN.WFO_PICK.get(sid)
        sel = [w for w in by_sid_wfos.get(sid, []) if pick is None or w["label"] == pick]
        starts = [w["start"] for w in sel if w["start"]]
        ends = [w["end"] for w in sel if w["end"]]

        if starts and ends:
            span = (min(starts), max(ends))
        else:
            span = parse_range_from_text(info["range_note"])
        if span:
            data_range = "%s ~ %s" % (span[0].strftime("%Y-%m-%d"), span[1].strftime("%Y-%m-%d"))
        else:
            data_range = "未紀錄"

        months = monthly_return = None
        if span:
            months = rnd((span[1] - span[0]).days / DAYS_PER_MONTH, 1)
            monthly_return = rnd(best["roa"] / months, 2)

        traded = [w for w in sel if w["trades"] > 0]
        if traded:
            pos = sum(1 for w in traded if w["net"] >= 0)
            neg = len(traded) - pos
            oos_text, oos_sum = "%d正%d負" % (pos, neg), sum(w["net"] for w in traded)
        else:
            oos_text, oos_sum = "—", None

        rows.append([
            sid, info["name"], CN.MARKET_PERIOD.get(sid, ""), data_range, best["round"],
            best["net_profit"], best["profit_factor"], best["total_trades"],
            best["pct_profitable"], best["max_dd"], best["roa"], months, monthly_return,
            best["params"], CN.WFO_RECOMMEND.get(sid, ""), oos_text, oos_sum,
            "待比對", CN.SPECIAL_NOTES.get(sid, ""), "見「平倉權益曲線圖」分頁",
        ])
    return rows, picked


# ---------------------------------------------------------------------------
# 各分頁
# ---------------------------------------------------------------------------
def curve_note():
    """平倉權益曲線圖分頁的狀態說明（依本機現況動態更新）。"""
    wsp = ""
    if os.path.isdir(MC_WSP_DIR):
        times = [os.path.getmtime(os.path.join(MC_WSP_DIR, f))
                 for f in os.listdir(MC_WSP_DIR) if f.lower().endswith(".wsp")]
        if times:
            wsp = "，最新 .wsp 修改時間 %s" % datetime.fromtimestamp(max(times)).strftime("%Y-%m-%d")
    return ("狀態：待補。須開啟 Capital MultiCharts 載入策略工作區、於「策略回測績效報告」"
            "擷取平倉權益曲線後貼入本分頁；亦可將截圖存成 screenshots\\<策略名稱>.png，"
            "重跑 scripts\\build_summary.py 即自動嵌入。"
            "注意：最佳化於 %s\\ 執行，該磁碟目前已連線本機%s（Capital MultiCharts64 在 "
            "C:\\Capital\\）；載入工作區後仍須確認 input 現值與「所有參數設定值」一致。"
            % (MC_WSP_DIR, wsp))


def write_main_sheet(ws, main_rows, files, strategies, build_date, source_dir):
    spec = SHEET_SPECS["彙整主表"]
    ws["A1"] = "台指期策略最佳化參數彙整（%d個策略 × 淨利最高一輪全期績效）" % len(strategies)
    ws["A2"] = ("資料來源：%s（%d份最佳化報告.xlsx）｜產出日期：%s｜"
                "口徑：每個策略取所有輪次中「淨利最高」的一輪全期績效｜"
                "月報酬率＝報酬率(RoA) ÷ 資料月數（區間以WFO視窗最早起點~最晚終點推估）"
                % (source_dir, len(files), build_date))
    ws["A3"] = ("⚠「所有參數設定值」＝最高淨利那一輪的實際參數；「R3 WFO推薦參數(參考)」＝"
                "前向最佳化各視窗推薦值，兩者刻意不同，切勿混用。"
                "第18欄「程式參數一致性」＝主表參數 vs .wsp input 現值自動比對結果，"
                "逐項明細見「wsp參數比對」分頁。"
                "黃色底＝該策略有重大警訊需複核。")
    for i, head in enumerate(MAIN_HEADERS, start=1):
        ws.cell(row=4, column=i, value=head)

    for r, row in enumerate(main_rows, start=5):
        for c, value in enumerate(row, start=1):
            cellobj = ws.cell(row=r, column=c)
            if value is not None:
                cellobj.value = value
            letter = get_column_letter(c)
            cellobj.font = F_BODY
            cellobj.alignment = A_BODY_TOP_WRAP if letter in spec["wrap_cols"] else A_BODY_TOP
            if letter in spec["numfmt"]:
                cellobj.number_format = spec["numfmt"][letter]
        if row[0] in CN.WARN_STRATEGIES:
            for c in range(1, len(MAIN_HEADERS) + 1):
                ws.cell(row=r, column=c).fill = C_WARN_FILL
    return ws


def write_detail_sheet(ws, rounds, picked, strategies):
    spec = SHEET_SPECS["全部輪次明細"]
    ws["A1"] = "各策略全部最佳化輪次的最佳結果（★ 標示者為彙整主表採用）"
    for i, head in enumerate(DETAIL_HEADERS, start=1):
        ws.cell(row=3, column=i, value=head)

    for r, item in enumerate(rounds, start=4):
        star = picked.get(item["sid"], ("", "")) == (item["file"], item["sheet"])
        row = [item["sid"], item["name"], item["round"], item["method"], item["candidates"],
               item["net_profit"], item["profit_factor"], item["total_trades"],
               item["pct_profitable"], item["max_dd"], item["roa"], item["avg_trade"],
               item["params"], item["sheet"], item["file"], "★" if star else None]
        for c, value in enumerate(row, start=1):
            cellobj = ws.cell(row=r, column=c)
            if value is not None:
                cellobj.value = value
            letter = get_column_letter(c)
            cellobj.font = F_BODY
            cellobj.alignment = A_BODY_TOP_WRAP if letter in spec["wrap_cols"] else A_BODY_TOP
            if letter in spec["numfmt"]:
                cellobj.number_format = spec["numfmt"][letter]
        if star:
            for c in range(1, len(DETAIL_HEADERS) + 1):
                ws.cell(row=r, column=c).fill = C_PICK_FILL
    return ws


def write_wfo_sheet(ws, wfos, strategies):
    spec = SHEET_SPECS["R3_WFO逐窗明細"]
    ws["A1"] = "R3 前向最佳化（WFO）逐窗樣本外明細（最後一列為無交易的收尾窗，OOS為0）"
    for i, head in enumerate(WFO_HEADERS, start=1):
        ws.cell(row=3, column=i, value=head)

    for r, item in enumerate(wfos, start=4):
        row = [item["sid"], item["name"], item["label"], item["window"],
               date_text(item["start"]) if item["start"] else None,
               date_text(item["is_end"]) if item["is_end"] else None,
               date_text(item["end"]) if item["end"] else None,
               item["net"], item["pf"], item["trades"], item["pct"], item["dd"],
               item["params"]]
        for c, value in enumerate(row, start=1):
            if value is None:
                continue
            cellobj = ws.cell(row=r, column=c, value=value)
            letter = get_column_letter(c)
            cellobj.font = F_BODY
            cellobj.alignment = A_BODY_TOP_WRAP if letter in spec["wrap_cols"] else A_BODY_TOP
        if item["net"] < 0:
            ws.cell(row=r, column=8).fill = C_NEG_FILL
        if item["trades"] == 0:
            ws.cell(row=r, column=9).fill = C_NOTRADE_FILL
    return ws


def write_wsp_sheet(ws, wsp_rows):
    spec = SHEET_SPECS["wsp參數比對"]
    ws["A1"] = "MultiCharts .wsp 程式參數一致性比對（主表「所有參數設定值」vs .wsp 的 input 現值）"
    for i, head in enumerate(WSP_HEADERS, start=1):
        ws.cell(row=3, column=i, value=head)

    for r, item in enumerate(wsp_rows, start=4):
        for c, value in enumerate(item, start=1):
            cellobj = ws.cell(row=r, column=c, value=value)
            letter = get_column_letter(c)
            cellobj.font = F_BODY
            cellobj.alignment = A_BODY_TOP_WRAP if letter in spec["wrap_cols"] else A_BODY_TOP
        if len(item) >= 7:
            verdict = item[6]
            if verdict == "DIFF":
                ws.cell(row=r, column=7).fill = C_WARN_FILL
            elif verdict == "ONLY_WSP":
                ws.cell(row=r, column=7).fill = C_NOTRADE_FILL
            elif verdict == "ONLY_REPORT":
                ws.cell(row=r, column=7).fill = C_NEG_FILL
    return ws


def write_curve_sheet(ws, strategies, source_files):
    ws["A1"] = "平倉權益曲線圖（MultiCharts 策略回測績效報告截圖）"
    note = curve_note()
    if note:
        ws["A2"] = note
        ws.merge_cells("A2:H4")
        ws["A2"].font = F_NOTE_RED
        ws["A2"].alignment = A_BODY_TOP_WRAP

    for i, sid in enumerate(sorted(strategies)):
        name = strategies[sid]["name"]
        row = 6 + 2 * i
        ws.cell(row=row, column=1, value=name).font = F_BODY_BOLD
        png = os.path.join(SCREENSHOT_DIR, name + ".png")
        if os.path.isfile(png):
            ws.cell(row=row, column=2, value="已嵌入截圖").font = F_BODY
            img = XLImage(png)
            scale = min(1.0, 900.0 / img.width)
            img.width = int(img.width * scale)
            img.height = int(img.height * scale)
            img.anchor = "A%d" % (row + 1)
            ws.add_image(img)
        else:
            ws.cell(row=row, column=2, value="（待補截圖）").font = F_BODY_GRAY
    return ws


def fix_info_text(text, source_dir):
    """把舊成果檔中已過時／誤植的路徑與說明文字一併更新。

    說明分頁的敘述列取自 curated_notes.INFO_ITEMS（沿用既有成果檔文字），
    下列替換讓路徑與現況說明跟著本機環境走。
    """
    text = text.replace("Muto-Agents", "Muti-Agents")
    text = text.replace(OLD_SOURCE, source_dir)
    text = text.replace("該磁碟不在本機", "該磁碟目前已連線本機")
    text = text.replace(
        "欄位值皆直接取自各報告明細分頁最佳列，未人工改寫。",
        "欄位值皆直接取自各報告明細分頁最佳列，未人工改寫；"
        "本檔由 scripts\\build_summary.py 自動產生，可重跑更新。")
    return text


def write_info_sheet(ws, strategies, files, source_dir):
    ws["A1"] = "資料來源、口徑說明與限制"
    ws["A3"] = "項目"
    ws["B3"] = "說明"
    items = list(CN.INFO_ITEMS)
    items[0] = (items[0][0],
                "13 個編號策略（0~12），來源 Muti-Agents 0924 工作區的 %d 份最佳化報告 .xlsx。"
                % len(files))
    for r, (item, desc) in enumerate(items, start=4):
        a = ws.cell(row=r, column=1, value=item)
        a.font = F_BODY_BOLD
        a.alignment = A_BODY_TOP_WRAP
        b = ws.cell(row=r, column=2, value=fix_info_text(desc, source_dir))
        b.font = F_BODY
        b.alignment = A_BODY_TOP_WRAP
    return ws


# ---------------------------------------------------------------------------
# 樣式套用
# ---------------------------------------------------------------------------
def apply_sheet_layout(ws, title, header_rows, last_col, title_cell, note_fonts=()):
    spec = SHEET_SPECS[title]
    for letter, width in spec["widths"].items():
        ws.column_dimensions[letter].width = width
    if spec["freeze"]:
        ws.freeze_panes = spec["freeze"]

    ws[title_cell].font = F_TITLE
    for ref, font in note_fonts:
        ws[ref].font = font

    for r in header_rows:
        for c in range(1, last_col + 1):
            cellobj = ws.cell(row=r, column=c)
            cellobj.fill = C_HEADER_FILL
            cellobj.font = F_HEADER
            cellobj.alignment = A_HEADER if title in ("彙整主表", "全部輪次明細",
                                                      "R3_WFO逐窗明細") else A_HEADER_PLAIN
    return ws


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------
def build(source_dir, output, build_date):
    files, rounds, wfos, strategies = collect(source_dir)
    if not strategies:
        raise SystemExit("來源目錄找不到任何「<編號>-<策略名>」開頭的報告 .xlsx：%s" % source_dir)

    main_rows, picked = build_main_rows(files, rounds, wfos, strategies)

    # .wsp 程式參數一致性比對：填主表第 18 欄（索引 17）＋收集「wsp參數比對」分頁明細
    wsp_rows = []
    wsp_use_best = {}   # sid -> 是否用到 _BEST.wsp
    for row in main_rows:
        sid, name = row[0], row[1]
        p14 = row[13]
        wsp_path = CW.resolve_wsp(sid, name, MC_WSP_DIR)
        wsp_use_best[sid] = wsp_path is not None and "_BEST" in os.path.basename(wsp_path)
        summary, detail = CW.compare_strategy(str(name), p14, wsp_path)
        row[17] = summary
        for d in detail:
            wsp_rows.append([sid, name] + list(d))

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws_main = wb.create_sheet("彙整主表")
    ws_detail = wb.create_sheet("全部輪次明細")
    ws_wfo = wb.create_sheet("R3_WFO逐窗明細")
    ws_wsp = wb.create_sheet("wsp參數比對")
    ws_curve = wb.create_sheet("平倉權益曲線圖")
    ws_info = wb.create_sheet("資料來源與限制")

    write_main_sheet(ws_main, main_rows, files, strategies, build_date, source_dir)
    write_detail_sheet(ws_detail, rounds, picked, strategies)
    write_wfo_sheet(ws_wfo, wfos, strategies)
    write_wsp_sheet(ws_wsp, wsp_rows)
    write_curve_sheet(ws_curve, strategies, files)
    write_info_sheet(ws_info, strategies, files, source_dir)

    apply_sheet_layout(ws_main, "彙整主表", [4], len(MAIN_HEADERS), "A1",
                       [("A2", F_NOTE_GRAY), ("A3", F_NOTE_RED)])
    apply_sheet_layout(ws_detail, "全部輪次明細", [3], len(DETAIL_HEADERS), "A1")
    apply_sheet_layout(ws_wfo, "R3_WFO逐窗明細", [3], len(WFO_HEADERS), "A1")
    apply_sheet_layout(ws_wsp, "wsp參數比對", [3], len(WSP_HEADERS), "A1")
    apply_sheet_layout(ws_curve, "平倉權益曲線圖", [], 2, "A1")
    apply_sheet_layout(ws_info, "資料來源與限制", [3], 2, "A1")

    os.makedirs(os.path.dirname(os.path.abspath(output)), exist_ok=True)
    wb.save(output)
    return {"strategies": len(strategies), "rounds": len(rounds), "wfo": len(wfos),
            "files": len(files), "output": output}


def main():
    parser = argparse.ArgumentParser(description="重新產生策略最佳化參數彙整 Excel")
    parser.add_argument("--source", default=DEFAULT_SOURCE, help="來源工作區（最佳化報告 .xlsx）")
    parser.add_argument("--out", default=None, help="輸出 .xlsx 路徑")
    parser.add_argument("--date", default=None, help="標示於主表的產出日期（預設今天）")
    args = parser.parse_args()

    build_date = args.date or date.today().strftime("%Y-%m-%d")
    output = args.out or os.path.join(
        PROJECT_DIR, "策略最佳化參數彙整_%s.xlsx" % date.today().strftime("%Y%m%d"))
    info = build(args.source, output, build_date)
    print("策略 %d 個／輪次 %d 列／WFO 視窗 %d 列／來源檔 %d 份" % (
        info["strategies"], info["rounds"], info["wfo"], info["files"]))
    print("已輸出：%s" % info["output"])


if __name__ == "__main__":
    main()
