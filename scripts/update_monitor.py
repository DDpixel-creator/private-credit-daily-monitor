import argparse
import csv
import io
import shutil
import statistics
from datetime import datetime
from pathlib import Path
from typing import Dict

import openpyxl
import requests
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

STOOQ_SYMBOLS = [
    "SPY", "XLF", "BLK", "BX", "APO", "ARES", "OWL", "KKR",
    "BIZD", "HYG", "BKLN", "JBBB", "LQD",
]

TRACKED_IDS = ["EW-02", "EW-03", "EW-04", "TR-06", "SY-01", "SY-04"]

HEADERS = [
    "编号",
    "一级模块",
    "二级模块",
    "监控指标",
    "说明",
    "更新频率",
    "最近更新时间",
    "状态",
    "状态依据（具体数据+来源）",
    "分数",
    "备注 / 下一步动作",
]


def fetch_text(url: str) -> str:
    headers = {"User-Agent": "Mozilla/5.0"}
    resp = requests.get(url, headers=headers, timeout=20)
    resp.raise_for_status()
    return resp.text


def stooq_last(symbol: str) -> Dict[str, float | str]:
    url = f"https://stooq.com/q/d/l/?s={symbol.lower()}.us&i=d"
    txt = fetch_text(url).strip().splitlines()
    rows = list(csv.DictReader(io.StringIO("\n".join(txt))))
    rows = [r for r in rows if r.get("Close")]

    if len(rows) < 2:
        raise RuntimeError(f"Not enough rows for Stooq symbol {symbol}")

    last = rows[-1]
    prev = rows[-6] if len(rows) >= 6 else rows[0]

    close_now = float(last["Close"])
    close_prev = float(prev["Close"])
    chg5d = (close_now / close_prev - 1.0) * 100 if close_prev else 0.0

    return {
        "date": last["Date"],
        "chg5d": chg5d,
        "close": close_now,
    }


def fred_last(series: str) -> Dict[str, float | str]:
    url = f"https://fred.stlouisfed.org/graph/fredgraph.csv?id={series}"
    txt = fetch_text(url).strip().splitlines()
    rows = list(csv.DictReader(io.StringIO("\n".join(txt))))
    rows = [r for r in rows if r.get(series) and r[series] != "."]

    if not rows:
        raise RuntimeError(f"No rows for FRED series {series}")

    last = rows[-1]
    return {
        "date": last["observation_date"],
        "value": float(last[series]),
    }


def create_minimal_workbook(path: Path) -> None:
    wb = Workbook()
    dash = wb.active
    dash.title = "Dashboard"
    dash["A1"] = "Private Credit Daily Monitor"
    dash["A3"] = "最近更新日期"
    dash["B3"] = ""

    checklist = wb.create_sheet("Checklist")
    for idx, header in enumerate(HEADERS, start=1):
        checklist.cell(row=5, column=idx, value=header)

    rows = [
        ("EW-02", "市场定价", "上市资管股", "资管股相对收益", "5日相对表现"),
        ("EW-03", "市场定价", "BDC", "BDC 相对高收益债", "5日相对表现"),
        ("EW-04", "市场定价", "信用压力", "HY/Loan/CLO 弱化", "综合弱项数"),
        ("TR-06", "融资条件", "短端利率", "SOFR", "利率水平"),
        ("SY-01", "系统环境", "金融条件", "ANFCI", "金融条件压力"),
        ("SY-04", "系统环境", "信用偏好", "IG vs HY", "LQD-HYG 与 IG OAS"),
    ]
    for row_idx, item in enumerate(rows, start=6):
        checklist.cell(row=row_idx, column=1, value=item[0])
        checklist.cell(row=row_idx, column=2, value=item[1])
        checklist.cell(row=row_idx, column=3, value=item[2])
        checklist.cell(row=row_idx, column=4, value=item[3])
        checklist.cell(row=row_idx, column=5, value=item[4])

    wb.save(path)


def ensure_master_template(master: Path, asset_template: Path) -> None:
    if master.exists():
        return

    if asset_template.exists():
        shutil.copy2(asset_template, master)
        return

    create_minimal_workbook(master)


def ensure_sheet(wb: openpyxl.Workbook, title: str) -> Worksheet:
    if title in wb.sheetnames:
        return wb[title]
    return wb.create_sheet(title)


def ensure_checklist_structure(ws: Worksheet) -> None:
    for idx, header in enumerate(HEADERS, start=1):
        if ws.cell(row=5, column=idx).value != header:
            ws.cell(row=5, column=idx, value=header)

    existing_ids = {
        ws.cell(row=r, column=1).value: r
        for r in range(6, ws.max_row + 1)
        if ws.cell(row=r, column=1).value
    }

    defaults = [
        ("EW-02", "市场定价", "上市资管股", "资管股相对收益", "5日相对表现"),
        ("EW-03", "市场定价", "BDC", "BDC 相对高收益债", "5日相对表现"),
        ("EW-04", "市场定价", "信用压力", "HY/Loan/CLO 弱化", "综合弱项数"),
        ("TR-06", "融资条件", "短端利率", "SOFR", "利率水平"),
        ("SY-01", "系统环境", "金融条件", "ANFCI", "金融条件压力"),
        ("SY-04", "系统环境", "信用偏好", "LQD-HYG 与 IG OAS", "相对偏好"),
    ]

    next_row = max(ws.max_row + 1, 6)
    for item in defaults:
        if item[0] not in existing_ids:
            row = next_row
            next_row += 1
            ws.cell(row=row, column=1, value=item[0])
            ws.cell(row=row, column=2, value=item[1])
            ws.cell(row=row, column=3, value=item[2])
            ws.cell(row=row, column=4, value=item[3])
            ws.cell(row=row, column=5, value=item[4])


def build_row_map(ws: Worksheet) -> Dict[str, int]:
    result: Dict[str, int] = {}
    for r in range(6, ws.max_row + 1):
        key = ws.cell(row=r, column=1).value
        if key:
            result[str(key)] = r
    return result


def set_status(ws: Worksheet, row: int, status: str, evidence: str, note: str = "") -> None:
    ws.cell(row=row, column=8, value=status)
    ws.cell(row=row, column=9, value=evidence)
    score = {"绿灯": 0, "黄灯": 1, "红灯": 2, "待更新": ""}.get(status, "")
    ws.cell(row=row, column=10, value=score)
    ws.cell(row=row, column=11, value=note)


def main() -> None:
    parser = argparse.ArgumentParser(description="Update private credit monitor workbook")
    parser.add_argument("--workspace", default=".", help="Directory for outputs")
    parser.add_argument(
        "--master",
        default="private_credit_monitor_master_template.xlsx",
        help="Master workbook filename inside workspace",
    )
    args = parser.parse_args()

    workspace = Path(args.workspace).resolve()
    workspace.mkdir(parents=True, exist_ok=True)

    master = workspace / args.master
    latest = workspace / "private_credit_monitor_latest.xlsx"
    daily = workspace / f"private_credit_monitor_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
    asset_template = Path(__file__).resolve().parents[1] / "assets" / "private_credit_monitor_template.xlsx"

    ensure_master_template(master, asset_template)

    wb = openpyxl.load_workbook(master)
    dashboard = ensure_sheet(wb, "Dashboard")
    checklist = ensure_sheet(wb, "Checklist")
    ensure_checklist_structure(checklist)

    row_map = build_row_map(checklist)

    market_data = {symbol: stooq_last(symbol) for symbol in STOOQ_SYMBOLS}
    fred = {
        "HY_OAS": fred_last("BAMLH0A0HYM2"),
        "IG_OAS": fred_last("BAMLC0A0CM"),
        "SOFR": fred_last("SOFR"),
        "ANFCI": fred_last("ANFCI"),
    }

    for rid in TRACKED_IDS:
        if rid not in row_map:
            raise KeyError(f"Required row id missing: {rid}")

    mgr_symbols = ["BLK", "BX", "APO", "ARES", "OWL", "KKR"]
    mgr_avg = statistics.mean(float(market_data[s]["chg5d"]) for s in mgr_symbols)
    bench_avg = statistics.mean([float(market_data["SPY"]["chg5d"]), float(market_data["XLF"]["chg5d"])])
    rel_mgr = mgr_avg - bench_avg

    row = row_map["EW-02"]
    status = "绿灯" if rel_mgr > -1 else ("黄灯" if rel_mgr > -3 else "红灯")
    evidence = (
        f"5日相对收益：资管股均值 {mgr_avg:.2f}% vs SPY/XLF 均值 {bench_avg:.2f}%，"
        f"差值 {rel_mgr:.2f}%。来源: Stooq ({market_data['BLK']['date']})."
    )
    set_status(checklist, row, status, evidence)

    rel_bdc = float(market_data["BIZD"]["chg5d"]) - float(market_data["HYG"]["chg5d"])
    row = row_map["EW-03"]
    status = "绿灯" if rel_bdc >= -1 else ("黄灯" if rel_bdc >= -3 else "红灯")
    evidence = (
        f"5日 BIZD {market_data['BIZD']['chg5d']:.2f}% vs HYG {market_data['HYG']['chg5d']:.2f}%，"
        f"相对 {rel_bdc:.2f}%。来源: Stooq ({market_data['BIZD']['date']})."
    )
    set_status(checklist, row, status, evidence)

    hy_oas = float(fred["HY_OAS"]["value"])
    loan_5d = float(market_data["BKLN"]["chg5d"])
    clo_5d = float(market_data["JBBB"]["chg5d"])
    weak_count = sum([
        1 if hy_oas > 4.0 else 0,
        1 if loan_5d < 0 else 0,
        1 if clo_5d < 0 else 0,
    ])

    row = row_map["EW-04"]
    status = "绿灯" if weak_count == 0 else ("黄灯" if weak_count == 1 else "红灯")
    evidence = (
        f"HY OAS={hy_oas:.2f}% ({fred['HY_OAS']['date']}), "
        f"BKLN 5日 {loan_5d:.2f}%, JBBB 5日 {clo_5d:.2f}%，弱项 {weak_count}/3。来源: FRED + Stooq."
    )
    set_status(checklist, row, status, evidence)

    sofr = float(fred["SOFR"]["value"])
    row = row_map["TR-06"]
    status = "绿灯" if sofr < 4 else ("黄灯" if sofr < 5 else "红灯")
    evidence = f"SOFR={sofr:.3f}% ({fred['SOFR']['date']})。来源: FRED."
    set_status(checklist, row, status, evidence)

    anfci = float(fred["ANFCI"]["value"])
    row = row_map["SY-01"]
    status = "绿灯" if anfci < 0 else ("黄灯" if anfci < 0.5 else "红灯")
    evidence = f"ANFCI={anfci:.3f} ({fred['ANFCI']['date']})。来源: FRED."
    set_status(checklist, row, status, evidence)

    ig_oas = float(fred["IG_OAS"]["value"])
    lqd_minus_hyg = float(market_data["LQD"]["chg5d"]) - float(market_data["HYG"]["chg5d"])
    row = row_map["SY-04"]
    status = "绿灯" if ig_oas < 1.4 and lqd_minus_hyg >= 0 else ("黄灯" if ig_oas < 1.8 else "红灯")
    evidence = (
        f"IG OAS={ig_oas:.2f}% ({fred['IG_OAS']['date']}); "
        f"LQD-HYG 5日相对={lqd_minus_hyg:.2f}%。来源: FRED + Stooq."
    )
    set_status(checklist, row, status, evidence)

    dashboard["A1"] = "Private Credit Daily Monitor"
    dashboard["A3"] = "最近更新日期"
    dashboard["B3"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    wb.save(master)
    wb.save(latest)
    wb.save(daily)

    print(f"Master updated: {master}")
    print(f"Latest snapshot: {latest}")
    print(f"Daily archive: {daily}")


if __name__ == "__main__":
    main()
