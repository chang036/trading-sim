# -*- coding: utf-8 -*-
"""
建立「投資模擬器」Excel 試算表。

用法：
    python build_xlsx.py                          # 產生空白範本（含範例列）
    python build_xlsx.py backup.json              # 用網頁版備份 JSON 填好資料
    python build_xlsx.py backup.json -o my.xlsx   # 指定輸出檔名

試算表分頁：
    儀表板    總入金 / 總出金 / 現金 / 持倉市值 / 總資產 / 報酬
    出入金    日期、類型（入金/出金）、金額、備註
    買賣紀錄  日期、類型、標的、數量、成交價、報價幣別、匯率、手續費、備註
    持倉      各標的數量、平均成本、現價、市值、未實現/已實現損益
    設定      基準幣別、匯率
    說明      使用方式
"""
import argparse
import json
import os
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule

# ---------- 樣式 ----------
NAVY = "1E293B"
BLUE = "2563EB"
LIGHT = "F1F5F9"
WHITE = "FFFFFF"
GREEN = "047857"
RED = "BE123C"

HDR_FILL = PatternFill("solid", fgColor=NAVY)
HDR_FONT = Font(color=WHITE, bold=True, size=10.5)
TITLE_FONT = Font(bold=True, size=16, color=NAVY)
SUB_FONT = Font(size=9.5, color="64748B")
BOLD = Font(bold=True)
THIN = Side(style="thin", color="D8E0EA")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
INPUT_FILL = PatternFill("solid", fgColor="FFFBEB")   # 黃底 = 手動輸入
CALC_FILL = PatternFill("solid", fgColor=LIGHT)       # 灰底 = 公式
MONEY = '#,##0.00'
PCT = '0.00%'
DATE = 'yyyy/mm/dd'
DATETIME = 'yyyy/mm/dd hh:mm'

FIRST, LAST = 2, 301  # 資料列範圍（含範例）


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER
    ws.row_dimensions[row].height = 24


def set_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def box(ws, r1, c1, r2, c2):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(row=r, column=c).border = BORDER


# ---------- 分頁 ----------
def build_settings(wb, base, fx):
    ws = wb.create_sheet("設定")
    ws["A1"] = "基準幣別"
    ws["B1"] = base
    ws["A2"] = "匯率（1 USD = ? TWD）"
    ws["B2"] = fx
    ws["A4"] = "說明"
    ws["A5"] = "基準幣別可填 TWD 或 USD；匯率只用於換算 USD 資產。"
    ws["A6"] = "改匯率後，所有 USD 資產的市值會跟著變，但歷史成本不變。"
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 12
    for r in (1, 2):
        ws.cell(row=r, column=1).font = BOLD
        ws.cell(row=r, column=2).fill = INPUT_FILL
        ws.cell(row=r, column=2).border = BORDER
        ws.cell(row=r, column=2).alignment = Alignment(horizontal="center")
    ws["A4"].font = BOLD
    ws["A5"].font = SUB_FONT
    ws["A6"].font = SUB_FONT
    dv = DataValidation(type="list", formula1='"TWD,USD"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("B1")
    ws.sheet_view.showGridLines = False
    return ws


def build_cashflow(wb, rows):
    ws = wb.create_sheet("出入金")
    heads = ["日期", "類型", "金額", "備註"]
    for i, h in enumerate(heads, start=1):
        ws.cell(row=1, column=i, value=h)
    style_header(ws, 1, len(heads))

    for r in range(FIRST, LAST + 1):
        ws.cell(row=r, column=1).number_format = DATE
        ws.cell(row=r, column=3).number_format = MONEY
        for c in range(1, 5):
            cell = ws.cell(row=r, column=c)
            cell.border = BORDER
            cell.fill = INPUT_FILL

    for i, row in enumerate(rows):
        r = FIRST + i
        ws.cell(row=r, column=1, value=row["date"])
        ws.cell(row=r, column=2, value=row["type"])
        ws.cell(row=r, column=3, value=row["amount"])
        ws.cell(row=r, column=4, value=row.get("note", ""))

    dv = DataValidation(type="list", formula1='"入金,出金"', allow_blank=True,
                        showErrorMessage=True, errorTitle="類型錯誤",
                        error="請填「入金」或「出金」")
    ws.add_data_validation(dv)
    dv.add(f"B{FIRST}:B{LAST}")

    set_widths(ws, [14, 10, 16, 40])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:D{LAST}"
    return ws


def build_trades(wb, rows):
    ws = wb.create_sheet("買賣紀錄")
    heads = ["日期", "類型", "標的", "數量", "成交價", "報價幣別", "匯率",
             "手續費", "買進成本", "賣出收入", "已實現損益",
             "持股(交易後)", "持倉成本(交易後)", "平均成本", "備註"]
    for i, h in enumerate(heads, start=1):
        ws.cell(row=1, column=i, value=h)
    style_header(ws, 1, len(heads))

    for r in range(FIRST, LAST + 1):
        # 手動輸入欄位
        for c in (1, 2, 3, 4, 5, 6, 8, 15):
            cell = ws.cell(row=r, column=c)
            cell.fill = INPUT_FILL
            cell.border = BORDER
        # 公式欄位
        ws[f"G{r}"] = (f'=IF($F{r}="","",IF($F{r}=設定!$B$1,1,'
                       f'IF(設定!$B$1="TWD",設定!$B$2,1/設定!$B$2)))')
        ws[f"I{r}"] = f'=IF($B{r}="買進",$D{r}*$E{r}*$G{r}+$H{r},"")'
        ws[f"J{r}"] = f'=IF($B{r}="賣出",$D{r}*$E{r}*$G{r}-$H{r},"")'
        ws[f"K{r}"] = (f'=IF($B{r}="賣出",$J{r}-IFERROR('
                       f'SUMIFS($M$1:M{r-1},$C$1:C{r-1},$C{r})/'
                       f'SUMIFS($L$1:L{r-1},$C$1:C{r-1},$C{r})*$D{r},0),"")')
        ws[f"L{r}"] = (f'=IF($B{r}="","",SUMIFS($D$1:D{r},$C$1:C{r},$C{r},$B$1:B{r},"買進")'
                       f'-SUMIFS($D$1:D{r},$C$1:C{r},$C{r},$B$1:B{r},"賣出"))')
        ws[f"M{r}"] = (f'=IF($B{r}="","",IF($B{r}="買進",'
                       f'IFERROR(SUMIFS($M$1:M{r-1},$C$1:C{r-1},$C{r}),0)+$I{r},'
                       f'IFERROR(SUMIFS($M$1:M{r-1},$C$1:C{r-1},$C{r})/'
                       f'SUMIFS($L$1:L{r-1},$C$1:C{r-1},$C{r}),0)*$L{r}))')
        ws[f"N{r}"] = f'=IFERROR($M{r}/$L{r},"")'
        for c in range(1, len(heads) + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = BORDER
            if c in (9, 10, 11, 12, 13, 14):
                cell.fill = CALC_FILL
        ws.cell(row=r, column=1).number_format = DATETIME
        for c in (5, 7, 8, 9, 10, 11, 13, 14):
            ws.cell(row=r, column=c).number_format = MONEY
        ws.cell(row=r, column=4).number_format = '0.########'

    for i, row in enumerate(rows):
        r = FIRST + i
        ws.cell(row=r, column=1, value=row["date"])
        ws.cell(row=r, column=2, value=row["type"])
        ws.cell(row=r, column=3, value=row["symbol"])
        ws.cell(row=r, column=4, value=row["qty"])
        ws.cell(row=r, column=5, value=row["price"])
        ws.cell(row=r, column=6, value=row["ccy"])
        ws.cell(row=r, column=7, value=row["fx"])
        ws.cell(row=r, column=8, value=row["fee"])
        ws.cell(row=r, column=15, value=row.get("note", ""))

    dv_type = DataValidation(type="list", formula1='"買進,賣出"', allow_blank=True,
                             showErrorMessage=True, errorTitle="類型錯誤",
                             error="請填「買進」或「賣出」")
    dv_ccy = DataValidation(type="list", formula1='"USD,TWD"', allow_blank=True,
                            showErrorMessage=True, errorTitle="幣別錯誤",
                            error="請填 USD 或 TWD")
    ws.add_data_validation(dv_type)
    ws.add_data_validation(dv_ccy)
    dv_type.add(f"B{FIRST}:B{LAST}")
    dv_ccy.add(f"F{FIRST}:F{LAST}")

    set_widths(ws, [16, 8, 11, 12, 12, 10, 9, 11, 13, 13, 13, 14, 17, 12, 24])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:O{LAST}"
    return ws


def build_positions(wb, assets):
    ws = wb.create_sheet("持倉")
    heads = ["標的", "名稱", "數量", "平均成本", "現價", "報價幣別", "匯率",
             "市值", "未實現損益", "已實現損益", "報酬率"]
    for i, h in enumerate(heads, start=1):
        ws.cell(row=1, column=i, value=h)
    style_header(ws, 1, len(heads))

    for i, a in enumerate(assets):
        r = FIRST + i
        ws.cell(row=r, column=1, value=a["symbol"])
        ws.cell(row=r, column=2, value=a["name"])
        ws.cell(row=r, column=5, value=a["price"])
        ws.cell(row=r, column=6, value=a["ccy"])
        ws[f"C{r}"] = (f'=SUMIFS(買賣紀錄!$D:$D,買賣紀錄!$C:$C,$A{r},買賣紀錄!$B:$B,"買進")'
                       f'-SUMIFS(買賣紀錄!$D:$D,買賣紀錄!$C:$C,$A{r},買賣紀錄!$B:$B,"賣出")')
        ws[f"D{r}"] = (f'=IFERROR(LOOKUP(2,1/(買賣紀錄!$C$2:$C${LAST}=$A{r}),'
                       f'買賣紀錄!$N$2:$N${LAST}),"")')
        ws[f"G{r}"] = (f'=IF($F{r}="","",IF($F{r}=設定!$B$1,1,'
                       f'IF(設定!$B$1="TWD",設定!$B$2,1/設定!$B$2)))')
        ws[f"H{r}"] = f'=IF($C{r}=0,0,$C{r}*$E{r}*$G{r})'
        ws[f"I{r}"] = f'=IF($C{r}=0,0,$C{r}*($E{r}*$G{r}-$D{r}))'
        ws[f"J{r}"] = f'=SUMIFS(買賣紀錄!$K:$K,買賣紀錄!$C:$C,$A{r})'
        ws[f"K{r}"] = f'=IFERROR($I{r}/($C{r}*$D{r}),"")'
        for c in range(1, len(heads) + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = BORDER
            cell.fill = CALC_FILL
        for c in (1, 2, 5, 6):
            ws.cell(row=r, column=c).fill = INPUT_FILL
        ws.cell(row=r, column=3).number_format = '0.########'
        for c in (4, 5, 8, 9, 10):
            ws.cell(row=r, column=c).number_format = MONEY
        ws.cell(row=r, column=11).number_format = PCT

    # 空白列（可自行新增標的）
    for r in range(FIRST + len(assets), FIRST + 20):
        ws.cell(row=r, column=1).fill = INPUT_FILL
        ws.cell(row=r, column=2).fill = INPUT_FILL
        ws.cell(row=r, column=5).fill = INPUT_FILL
        ws.cell(row=r, column=6).fill = INPUT_FILL
        ws[f"C{r}"] = (f'=IF($A{r}="",0,SUMIFS(買賣紀錄!$D:$D,買賣紀錄!$C:$C,$A{r},買賣紀錄!$B:$B,"買進")'
                       f'-SUMIFS(買賣紀錄!$D:$D,買賣紀錄!$C:$C,$A{r},買賣紀錄!$B:$B,"賣出"))')
        ws[f"D{r}"] = (f'=IFERROR(LOOKUP(2,1/(買賣紀錄!$C$2:$C${LAST}=$A{r}),'
                       f'買賣紀錄!$N$2:$N${LAST}),"")')
        ws[f"G{r}"] = (f'=IF($F{r}="","",IF($F{r}=設定!$B$1,1,'
                       f'IF(設定!$B$1="TWD",設定!$B$2,1/設定!$B$2)))')
        ws[f"H{r}"] = f'=IF($C{r}=0,0,$C{r}*$E{r}*$G{r})'
        ws[f"I{r}"] = f'=IF($C{r}=0,0,$C{r}*($E{r}*$G{r}-$D{r}))'
        ws[f"J{r}"] = f'=SUMIFS(買賣紀錄!$K:$K,買賣紀錄!$C:$C,$A{r})'
        ws[f"K{r}"] = f'=IFERROR($I{r}/($C{r}*$D{r}),"")'
        for c in range(1, len(heads) + 1):
            ws.cell(row=r, column=c).border = BORDER
        ws.cell(row=r, column=3).number_format = '0.########'
        for c in (4, 5, 8, 9, 10):
            ws.cell(row=r, column=c).number_format = MONEY
        ws.cell(row=r, column=11).number_format = PCT

    set_widths(ws, [11, 24, 12, 12, 12, 10, 9, 15, 15, 15, 11])
    ws.freeze_panes = "A2"
    return ws


def build_dashboard(wb):
    ws = wb.create_sheet("儀表板", 0)
    ws.sheet_view.showGridLines = False
    ws["A1"] = "投資模擬器 · 損益儀表板"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = "黃底欄位可手動輸入；灰底為公式，請勿覆蓋。"
    ws["A2"].font = SUB_FONT

    rows = [
        ("基準幣別", '=設定!$B$1', None),
        ("匯率（1 USD = ? TWD）", '=設定!$B$2', '0.00'),
        (None, None, None),
        ("總入金", '=SUMIFS(出入金!$C:$C,出入金!$B:$B,"入金")', MONEY),
        ("總出金", '=SUMIFS(出入金!$C:$C,出入金!$B:$B,"出金")', MONEY),
        ("淨投入", "=B7-B8", MONEY),
        ("現金餘額", "=B9-SUM(買賣紀錄!$I:$I)+SUM(買賣紀錄!$J:$J)", MONEY),
        ("持倉市值", "=SUM(持倉!$H:$H)", MONEY),
        ("總資產", "=B10+B11", MONEY),
        (None, None, None),
        ("累計報酬", "=B12-B9", MONEY),
        ("報酬率", '=IFERROR(B14/B9,"")', PCT),
        ("已實現損益", "=SUM(買賣紀錄!$K:$K)", MONEY),
        ("未實現損益", "=SUM(持倉!$I:$I)", MONEY),
        ("累計手續費", "=SUM(買賣紀錄!$H:$H)", MONEY),
    ]

    r = 4
    for label, formula, fmt in rows:
        if label is None:
            r += 1
            continue
        ws.cell(row=r, column=1, value=label).font = BOLD
        ws.cell(row=r, column=1).border = BORDER
        cell = ws.cell(row=r, column=2, value=formula)
        cell.border = BORDER
        cell.fill = CALC_FILL
        cell.alignment = Alignment(horizontal="right")
        if fmt:
            cell.number_format = fmt
        ws.cell(row=r, column=1).fill = PatternFill("solid", fgColor=WHITE)
        r += 1

    ws["A19"] = "計算方式"
    ws["A19"].font = BOLD
    notes = [
        "現金 = 淨投入 − 買進成本 + 賣出收入（含手續費）",
        "持倉成本採移動平均法；賣出時認列已實現損益",
        "總報酬 = 總資產 − 淨投入（淨投入 = 入金 − 出金）",
        "手續費一律以基準幣別計收",
    ]
    for i, n in enumerate(notes, start=20):
        ws.cell(row=i, column=1, value="· " + n).font = SUB_FONT

    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 20

    for ref in ("B14", "B15", "B16", "B17"):
        ws.conditional_formatting.add(
            ref, CellIsRule(operator="lessThan", formula=["0"], font=Font(color=RED, bold=True)))
        ws.conditional_formatting.add(
            ref, CellIsRule(operator="greaterThan", formula=["0"], font=Font(color=GREEN, bold=True)))
    return ws


def build_help(wb):
    ws = wb.create_sheet("說明")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 100
    lines = [
        ("投資模擬器 · Excel 版使用說明", TITLE_FONT),
        ("", None),
        ("1. 先在「設定」填基準幣別（TWD 或 USD）與匯率。", None),
        ("2. 「出入金」分頁：每一筆存入或提出都記一列，金額用基準幣別。", None),
        ("3. 「買賣紀錄」分頁：每一筆買進或賣出記一列。", None),
        ("   · 匯率欄填「1 單位報價幣別 = ? 基準幣別」，例如 USD→TWD 填 31.5、TWD 標的填 1。", None),
        ("   · 右側灰色欄位（買進成本、已實現損益、平均成本…）會自動計算，不要手動改。", None),
        ("4. 「持倉」分頁：數量與平均成本自動帶入；只要更新「現價」就會算出市值與未實現損益。", None),
        ("5. 「儀表板」會自動彙總總入金、總出金、現金、持倉市值、總資產與報酬率。", None),
        ("", None),
        ("小技巧", BOLD),
        ("· 想加新標的：在「持倉」下方空白列填代號、名稱、現價、幣別即可。", None),
        ("· 匯入網頁版資料：在網頁按「設定 → 備份 JSON」，再執行", None),
        ("    python build_xlsx.py trading-sim-backup-xxxx.json", None),
        ("  就會產生一份填好資料的 Excel。", None),
        ("· Excel 版為「基準幣別總帳」：能正確算出入金、買賣、已實現/未實現損益（基準幣），", None),
        ("  但不會分開追蹤 TWD/USD 現金錢包與匯率損益；完整多幣別請用網頁版。", None),
        ("· 這份試算表只做模擬，不會連結任何真實帳戶或下單。", None),
    ]
    for i, (text, font) in enumerate(lines, start=1):
        c = ws.cell(row=i, column=1, value=text)
        if font:
            c.font = font
        c.alignment = Alignment(vertical="center")
    return ws


# ---------- 公式中的分頁名稱加引號 ----------
def quote_sheets(wb):
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, str) and v.startswith("="):
                    for name in ("出入金", "買賣紀錄", "持倉", "設定"):
                        v = v.replace(name + "!", "'" + name + "'!")
                    cell.value = v


# ---------- 資料 ----------
DEFAULT_ASSETS = [
    {"symbol": "BTC", "name": "Bitcoin", "price": 59700, "ccy": "USD"},
    {"symbol": "ETH", "name": "Ethereum", "price": 1676, "ccy": "USD"},
    {"symbol": "SOL", "name": "Solana", "price": 70.15, "ccy": "USD"},
    {"symbol": "VOO", "name": "Vanguard S&P 500 ETF", "price": 690, "ccy": "USD"},
    {"symbol": "VT", "name": "Vanguard Total World", "price": 150, "ccy": "USD"},
    {"symbol": "0050", "name": "元大台灣50", "price": 84.45, "ccy": "TWD"},
]


def sample_data():
    from datetime import timedelta
    now = datetime.now().replace(microsecond=0)
    d = lambda days: now - timedelta(days=days)
    cash = [
        {"date": d(30), "type": "入金", "amount": 300000, "note": "期初資金（範例，可刪）"},
        {"date": d(3), "type": "出金", "amount": 20000, "note": "生活費（範例，可刪）"},
    ]
    trades = [
        {"date": d(28), "type": "買進", "symbol": "BTC", "qty": 0.02, "price": 55000,
         "ccy": "USD", "fx": 31.5, "fee": 34.65, "note": "範例，可刪"},
        {"date": d(24), "type": "買進", "symbol": "VOO", "qty": 4, "price": 650,
         "ccy": "USD", "fx": 31.5, "fee": 65.52, "note": "範例，可刪"},
        {"date": d(14), "type": "買進", "symbol": "0050", "qty": 200, "price": 80,
         "ccy": "TWD", "fx": 1, "fee": 12.8, "note": "範例，可刪"},
        {"date": d(7), "type": "賣出", "symbol": "BTC", "qty": 0.005, "price": 62000,
         "ccy": "USD", "fx": 31.5, "fee": 9.77, "note": "範例，可刪"},
    ]
    return cash, trades


def from_backup(path):
    with open(path, "r", encoding="utf-8") as f:
        st = json.load(f)
    cash, trades = [], []
    ledger = sorted(st.get("ledger", []), key=lambda e: (e.get("t", ""), e.get("seq", 0)))
    for e in ledger:
        try:
            dt = datetime.fromisoformat(str(e.get("t", "")).replace("Z", "+00:00")).replace(tzinfo=None)
        except Exception:
            dt = datetime.now()
        typ = e.get("type")
        if typ in ("deposit", "withdraw"):
            amt_base = float(e.get("amountBase", e.get("amount", 0)))
            note = e.get("note", "")
            ccy = e.get("ccy")
            if ccy and ccy != st.get("base", "TWD"):
                note = (f"{ccy} {float(e.get('amount',0)):,.2f} " + note).strip()
            cash.append({
                "date": dt,
                "type": "入金" if typ == "deposit" else "出金",
                "amount": amt_base,
                "note": note,
            })
        elif typ == "convert":
            # Excel 版為單一基準幣總帳；換匯以基準幣現金影響近似記錄
            base = st.get("base", "TWD")
            frm, to = e.get("from"), e.get("to")
            if frm == base:
                cash.append({"date": dt, "type": "出金", "amount": float(e.get("amountFrom", 0)),
                             "note": f"換匯 {frm}->{to}（Excel 單幣別近似）"})
            if to == base:
                cash.append({"date": dt, "type": "入金", "amount": float(e.get("amountTo", 0)),
                             "note": f"換匯 {frm}->{to}（Excel 單幣別近似）"})
        elif typ in ("buy", "sell"):
            price = float(e.get("price", 0)) or 1
            pb = float(e.get("priceBase", 0))
            fx = pb / price if price else 1
            trades.append({
                "date": dt,
                "type": "買進" if typ == "buy" else "賣出",
                "symbol": e.get("symbol", ""),
                "qty": float(e.get("qty", 0)),
                "price": price,
                "ccy": e.get("ccy", "USD"),
                "fx": round(fx, 6),
                "fee": float(e.get("fee", 0)),
                "note": e.get("note", ""),
            })
    assets = []
    for a in st.get("assets", []):
        assets.append({"symbol": a.get("symbol"), "name": a.get("name", ""),
                       "price": float(a.get("price", 0)), "ccy": a.get("ccy", "USD")})
    return st.get("base", "TWD"), float(st.get("fx", 31.5)), cash, trades, assets


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("backup", nargs="?", help="網頁版備份 JSON")
    ap.add_argument("-o", "--out", default=None)
    args = ap.parse_args()

    if args.backup:
        base, fx, cash, trades, assets = from_backup(args.backup)
        out = args.out or ("投資模擬器_" + datetime.now().strftime("%Y%m%d") + ".xlsx")
    else:
        base, fx = "TWD", 31.5
        cash, trades = sample_data()
        assets = DEFAULT_ASSETS
        out = args.out or "投資模擬器_範本.xlsx"

    wb = Workbook()
    wb.remove(wb.active)

    build_dashboard(wb)
    build_cashflow(wb, cash)
    build_trades(wb, trades)
    build_positions(wb, assets)
    build_settings(wb, base, fx)
    build_help(wb)

    wb.calculation.fullCalcOnLoad = True
    quote_sheets(wb)
    wb.active = 0
    out_path = os.path.abspath(out)
    wb.save(out_path)
    print("已建立：", out_path)
    print(f"  分頁：儀表板 / 出入金 / 買賣紀錄 / 持倉 / 設定 / 說明")
    print(f"  出入金 {len(cash)} 筆、買賣 {len(trades)} 筆、標的 {len(assets)} 檔")


if __name__ == "__main__":
    main()
