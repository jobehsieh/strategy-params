# -*- coding: utf-8 -*-
"""wsp 程式參數比對工具

比對「彙整主表的『所有參數設定值』（＝最高淨利輪的實際參數，取自最佳化報告）」
與「MultiCharts .wsp 工作檔中策略 SignalObject 的 input 現值」。

產出：
    1. 主表第 18 欄「MultiCharts 程式參數一致性」的摘要文字
    2. 新增分頁「wsp參數比對」逐參數明細

用法（由 build_summary.py 呼叫；亦可獨立執行）
    python scripts\\compare_wsp.py --main <彙整 xlsx> --wsp-dir <E:\\Jobe\\群益交易程式>
"""
from __future__ import annotations

import argparse
import io
import os
import re

# ---------------------------------------------------------------------------
# .wsp 解析
# ---------------------------------------------------------------------------
RE_WSP_INPUT = re.compile(
    r"^\[Wsp\\(Window_\d+)\\ChartManager\\(?:Chart_\d+\\Series_\d+\\IndicatorHelper"
    r"|Strategy\\SignalObject_\d+)\\SignalHelper\\InputHelper\\(Input_.*?)\]",
    re.MULTILINE
)
RE_INPUT_SECTION = re.compile(
    r"^\[Wsp\\(Window_\d+)\\ChartManager\\Strategy\\SignalObject_\d+"
    r"\\SignalHelper\\InputHelper\\(Input_[^\]]+)\]\s*\n(.*?)(?=\n\[|\Z)",
    re.DOTALL | re.MULTILINE
)
RE_WSP_VALUE = re.compile(r"(?m)^\s*Value\s*=\s*'([^']*)'\s*$")


def parse_wsp_intputs(path):
    """回傳 {window: {input_name: value_string}}。

    只收 Strategy 的 SignalObject（策略訊號），不收 IndicatorHelper（指標）。
    """
    if not os.path.isfile(path):
        return None
    raw = open(path, "rb").read()
    text = raw.decode("big5", "replace")
    out = {}
    for m in RE_INPUT_SECTION.finditer(text):
        win = m.group(1)
        iname = m.group(2)  # 例如 Input_Man
        block = m.group(3)
        vm = RE_WSP_VALUE.search(block)
        val = vm.group(1) if vm else None
        out.setdefault(win, {})[iname[len("Input_"):]] = val
    return out


# ---------------------------------------------------------------------------
# 主表參數解析
# ---------------------------------------------------------------------------
# 策略 5 的主表「所有參數設定值」混入了衍生績效指標（非策略 input），要排除
DERIVED_COLS = {
    "AvgTrade", "AvgWinningTrade", "AvgLosingTrade", "WinLossRatio",
    "MaxConsecWinners", "MaxConsecLosers", "AvgBarsInWinner", "AvgBarsInLoser",
}

# 公式型／枚舉型 input 的特殊比對規則
# 報告以數值表示、.wsp 以運算式或布林字串表示
PRICE_EXPRS = {
    "(h+l)*0.5": 0,   # Median 價格（程式預設），報告最佳化表以 0 匯出
}


def parse_report_params(p14_text):
    """'N1=30 / UP=1 / ...' → {name: raw_string}"
    """
    if not p14_text:
        return {}
    out = {}
    for part in p14_text.split(" / "):
        part = part.strip()
        if not part:
            continue
        if "=" not in part:
            continue
        k, v = part.split("=", 1)
        k, v = k.strip(), v.strip()
        if k in DERIVED_COLS:
            continue
        out[k] = v
    return out


def _to_num(s):
    s = str(s).strip().replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


def _typ(s):
    """回傳 (型別, 值)：bool / num / str。型別由字串內容決定。"""
    s = str(s).strip()
    low = s.lower()
    if low in ("true", "false"):
        return ("bool", low)
    n = _to_num(s)
    if n is not None:
        # 純整數顯示不帶小數點
        if n == int(n):
            return ("num", int(n))
        return ("num", n)
    return ("str", re.sub(r"\s+", "", low))


def compare_value(report_raw, wsp_raw):
    """回傳 (判定, 說明)。判定: OK / DIFF / SPECIAL / ONLY_REPORT / ONLY_WSP"""
    if wsp_raw is None:
        return "ONLY_REPORT", "%s（wsp 無此 input）" % report_raw
    # Price 公式型特例：報告枚舉值 vs .wsp 運算式
    wsp_key = re.sub(r"\s+", "", wsp_raw.lower())
    if wsp_key in PRICE_EXPRS and _to_num(report_raw) == PRICE_EXPRS[wsp_key]:
        return "SPECIAL", "報告 %s（枚舉值）＝wsp %s（運算式），同義" % (
            report_raw, wsp_raw)
    rt, rv = _typ(report_raw)
    wt, wv = _typ(wsp_raw)
    # 布林 vs 數字：wsp 若為布林，報告 1/0 視同 true/false
    if wt == "bool" and rt == "num":
        return _judge_num_bool(rv, wv, report_raw, wsp_raw)
    if rt == "bool" and wt == "num":
        return _judge_num_bool(wv, rv, wsp_raw, report_raw)
    if rt == wt:
        if rt == "num":
            if rv == wv:
                return "OK", str(rv)
            return "DIFF", "%s ≠ %s" % (report_raw, wsp_raw)
        if rt == "str":
            if rv == wv:
                return "OK", wsp_raw
            return "DIFF", "%s ≠ %s" % (report_raw, wsp_raw)
        return "OK", wsp_raw
    return "DIFF", "%s ≠ %s" % (report_raw, wsp_raw)


def _judge_num_bool(num_val, bool_s, num_disp, bool_disp):
    target = "true" if num_val == 1 else ("false" if num_val == 0 else None)
    if target is None:
        return "DIFF", "%s（數字） vs %s（布林），無法比對" % (num_disp, bool_disp)
    if target == bool_s:
        return "OK", "%s(=%s)" % (num_disp, target)
    return "DIFF", "%s vs %s" % (num_disp, bool_disp)


# ---------------------------------------------------------------------------
# .wsp 檔名對應
# ---------------------------------------------------------------------------
def _norm_wsp_body(s):
    """去副檔名、檔頭編號、_BEST 尾綴與空白，回傳策略名稱主體。"""
    s = re.sub(r"(?i)\.wsp\s*$", "", str(s).strip())
    s = re.sub(r"^\d+\s*[-–—]\s*", "", s)
    s = re.sub(r"(?i)_best\s*$", "", s)
    return re.sub(r"\s+", "", s).lower()


def resolve_wsp(sid, name, wsp_dir):
    """依主表策略找到工作區檔：優先 <sid>-<策略名>_BEST.wsp，其次 <sid>-*.wsp。

    回傳絕對路徑，找不到回 None。
    """
    if not os.path.isdir(wsp_dir):
        return None
    base = _norm_wsp_body(name)
    matching = []
    for f in sorted(os.listdir(wsp_dir)):
        if not f.lower().endswith(".wsp"):
            continue
        if not re.match(r"^\d+\s*[-–—]", f):
            continue
        if _norm_wsp_body(f) == base:
            matching.append(f)
    if not matching:
        return None
    best = [f for f in matching if "_BEST" in f.upper()]
    return os.path.join(wsp_dir, (best or matching)[0])


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------
def compare_strategy(sid_name, p14_text, wsp_path):
    """回傳 (summary_text, detail_rows)"""
    detail = []
    report = parse_report_params(p14_text)
    wsp = parse_wsp_intputs(wsp_path)
    if wsp is None:
        return ("無法比對（找不到 .wsp）", detail)
    if not wsp:
        return ("無法比對（.wsp 無策略 input）", detail)

    diffs, specials, only_wsp = [], [], []
    for win, inputs in sorted(wsp.items()):
        for name, wv in sorted(inputs.items()):
            rv = None
            for rn, rvv in report.items():
                if rn.lower() == name.lower():
                    rv = rvv
                    break
            if rv is None:
                only_wsp.append((win, name, wv))
                detail.append((win, name, "-", wv, "ONLY_WSP"))
                continue
            verdict, note = compare_value(rv, wv)
            detail.append((win, name, rv, wv, verdict, note))
            if verdict == "DIFF":
                diffs.append((win, name, rv, wv))
            elif verdict == "SPECIAL":
                specials.append((win, name, rv, wv))

    # 報告有、wsp 完全沒有的 input
    for rname, rv in sorted(report.items()):
        found = any(rname.lower() == n.lower() for _, ins in wsp.items() for n in ins)
        if not found:
            detail.append(("-", rname, rv, "-", "ONLY_REPORT",
                           "報告有此參數，wsp 無對應 input"))
            diffs.append(("-", rname, rv, "(wsp 無)"))

    if not diffs and not only_wsp:
        return ("一致", detail)
    bits = []
    if diffs:
        bits.append("差異 %d 項" % len(diffs))
    if only_wsp:
        bits.append("wsp 獨有 %d 項" % len(only_wsp))
    if specials:
        bits.append("型別特例 %d 項" % len(specials))
    return ("不一致（%s）" % "、".join(bits), detail)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--main", default=r"H:\我的雲端硬碟\最佳化參數彙整\策略最佳化參數彙整_20260928.xlsx")
    ap.add_argument("--wsp-dir", default=r"E:\Jobe\群益交易程式")
    args = ap.parse_args()

    import openpyxl
    wb = openpyxl.load_workbook(args.main)
    ws = wb["彙整主表"]
    out = io.open(r"C:\Users\Administrator\AppData\Local\Temp\opencode\wsp_compare.txt",
                  "w", encoding="utf-8")
    for r in range(5, ws.max_row + 1):
        sid = ws.cell(row=r, column=1).value
        name = ws.cell(row=r, column=2).value
        p14 = ws.cell(row=r, column=14).value
        wsp_path = resolve_wsp(sid, name, args.wsp_dir)
        summary, detail = compare_strategy(name, p14, wsp_path)
        out.write("SID=%s %s\n  wsp=%s\n  %s\n" % (sid, name,
                  os.path.basename(wsp_path) if os.path.isfile(wsp_path) else "(找不到)",
                  summary))
        for row in detail:
            out.write("    %s\n" % " | ".join(str(x) for x in row))
    out.close()
    print("done")


if __name__ == "__main__":
    main()