# -*- coding: utf-8 -*-
"""
stockprice_core.py
StockPrice_ver2.py の計算ロジックを Web アプリから呼べるよう関数化したもの。

- CSV (IR BANK 形式の fy-balance-sheet.csv / fy-profit-and-loss.csv) の読み込み
- 銘柄コードから IR BANK の CSV を取得（失敗時は Yahoo Finance の財務データで代替）
- Yahoo Finance から株価・流動資産を取得
- 資産価値・事業価値・理論株価・上昇余地・PSR・時価総額/正味価値・投資判断の計算
- テンプレート StockPrice.xlsx への書き込みと Excel バイト列の生成
"""

import csv
import datetime as dt
import io
import os
import re
from dataclasses import dataclass, field

import openpyxl as px
import requests

TEMPLATE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "StockPrice.xlsx")

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36"
}

UPSIDE_THRESHOLD = 30  # 投資の下限（上昇余地 %）。テンプレートの E10 と同じ値
ROE_ASSUMED_PER = 150  # テンプレート F4 の値（事業価値の倍率）

# 証券コード（例: 7203, 130A）。"2024/03" のような年度表記は除外する
CODE_RE = re.compile(r"(?<![0-9A-Za-z/])([0-9][0-9A-Z]{2}[0-9A-Z])(?![0-9A-Za-z/])")


# ============================================================
# 共通ユーティリティ
# ============================================================
def safe_float(val):
    """文字列・数値・None が混在するデータから安全に float を取得する（取得できなければ None）"""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace(",", "").replace("¥", "").replace("$", "").replace("%", "")
    if s in ("", "-", "－", "—"):
        return None
    s = re.sub(r"[^0-9.\-]", "", s)
    try:
        return float(s) if s not in ("", "-", ".") else None
    except ValueError:
        return None


def decode_csv(data):
    """CSV のバイト列を文字列に変換（UTF-8 / BOM付き / Shift_JIS に対応）"""
    if isinstance(data, str):
        return data
    for enc in ("utf-8-sig", "cp932"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace")


def read_csv_rows(data):
    """CSV を行のリスト（各行は文字列リスト）として返す。空行も保持する（元スクリプトの行番号と合わせるため）"""
    text = decode_csv(data)
    return [row for row in csv.reader(io.StringIO(text), delimiter=",")]


def _last_row_index(rows):
    """A 列が空でない最終行のインデックス（元スクリプトの maxRow 判定と同じ）"""
    for i in reversed(range(len(rows))):
        if rows[i] and rows[i][0].strip() != "":
            return i
    raise ValueError("CSV にデータ行がありません。")


def _header_index(rows, names, default):
    """見出し行から列番号を探す。見つからなければ元スクリプトの固定列番号を使う"""
    for row in rows[:10]:
        for idx, cell in enumerate(row):
            if cell.strip() in names:
                return idx
    return default


def _cell(rows, r, c):
    try:
        return rows[r][c]
    except IndexError:
        return None


# ============================================================
# 入力データ
# ============================================================
@dataclass
class FinancialInputs:
    code: str = ""
    company: str = ""           # Excel の A1 / ファイル名に使う表示名
    bps: float = None           # D2
    equity_ratio: float = None  # F2（小数。例 0.45）
    eps: float = None           # D4
    roa: float = None           # E4（小数）
    sales: float = None         # G12（円）
    total_assets: float = None  # 円
    net_assets: float = None    # 円
    shareholders_equity: float = None  # 円
    current_assets_million: float = None  # G14（百万円）
    stock_price: float = None   # C8（円）
    shares_outstanding: float = None
    source: str = ""
    raw_pl: bytes = None        # 取得した PL CSV（IR レポートで再利用）
    notes: list = field(default_factory=list)


def parse_irbank_csv(bs_data, pl_data):
    """IR BANK 形式の BS/PL CSV から指標を取り出す（StockPrice_ver2.py のセル3前半と同じ列）"""
    bs = read_csv_rows(bs_data)
    pl = read_csv_rows(pl_data)
    inp = FinancialInputs(source="CSV")

    # ---- BS ----
    r = _last_row_index(bs)
    c_total = _header_index(bs, ("総資産",), 1)
    c_net = _header_index(bs, ("純資産",), 2)
    c_equity = _header_index(bs, ("株主資本",), 3)
    c_bps = _header_index(bs, ("BPS",), 7)
    c_ratio = _header_index(bs, ("自己資本比率",), 8)

    inp.bps = safe_float(_cell(bs, r, c_bps))
    ratio = safe_float(_cell(bs, r, c_ratio))
    inp.equity_ratio = ratio / 100 if ratio is not None else None
    inp.total_assets = safe_float(_cell(bs, r, c_total))
    inp.net_assets = safe_float(_cell(bs, r, c_net))
    inp.shareholders_equity = safe_float(_cell(bs, r, c_equity))
    if inp.bps is None:
        raise ValueError(f"BS の最終行（{r + 1} 行目）から BPS を読み取れません。ファイルを確認してください。")
    if inp.equity_ratio is None:
        raise ValueError(f"BS の最終行（{r + 1} 行目）から自己資本比率を読み取れません。")

    # ---- PL ----
    m = _last_row_index(pl)
    c_sales = _header_index(pl, ("売上高", "営業収益", "経常収益"), 1)
    c_eps = _header_index(pl, ("EPS",), 5)
    c_roa = _header_index(pl, ("ROA",), 7)

    is_forecast = any("予想" in str(v) for v in pl[m])
    if is_forecast:
        inp.notes.append(f"PL の最終行（{m + 1} 行目）は予想データです。EPS は予想値を使います。")

    inp.eps = safe_float(_cell(pl, m, c_eps))
    roa = safe_float(_cell(pl, m, c_roa))
    if roa is None and m - 1 >= 0:
        roa = safe_float(_cell(pl, m - 1, c_roa))
        if roa is not None:
            inp.notes.append("最終行に ROA が無いため、1 つ前の行の ROA を使いました。")
    inp.roa = roa / 100 if roa is not None else None

    # 売上高：元スクリプトと同じく maxRow-1 行目（予想行がある前提の最新実績行）
    sales_row = m - 1 if is_forecast else m
    inp.sales = safe_float(_cell(pl, sales_row, c_sales))
    if inp.sales is None and sales_row != m - 1:
        inp.sales = safe_float(_cell(pl, m - 1, c_sales))

    if inp.eps is None:
        raise ValueError(f"PL の最終行（{m + 1} 行目）から EPS を読み取れません。")
    if inp.roa is None:
        raise ValueError("PL から ROA を読み取れません。")

    # ---- 証券コード・会社名（元スクリプトは PL の A3） ----
    label = None
    if len(pl) > 2 and pl[2] and CODE_RE.search(pl[2][0]):  # 元スクリプトと同じ PL の A3 を優先
        label = pl[2][0].strip()
    for row in ([] if label else pl[:5] + bs[:5]):
        for cell in row:
            if CODE_RE.search(cell or ""):
                label = cell.strip()
                break
        if label:
            break
    if label is None and len(pl) > 2 and pl[2]:
        label = pl[2][0].strip()
    if label:
        mcode = CODE_RE.search(label)
        inp.code = mcode.group(1) if mcode else label[:4]
        inp.company = label
    return inp


# ============================================================
# 外部データ取得
# ============================================================
def normalize_code(code):
    code = (code or "").strip().upper().replace(".T", "")
    code = code.translate(str.maketrans("０１２３４５６７８９ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ",
                                        "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"))
    if not re.fullmatch(r"[0-9][0-9A-Z]{2}[0-9A-Z]", code):
        raise ValueError("銘柄コードは 4 桁（例: 7203、130A）で入力してください。")
    return code


def fetch_irbank_csv(code):
    """IR BANK から fy-balance-sheet.csv / fy-profit-and-loss.csv を取得する"""
    out = {}
    for name in ("fy-balance-sheet.csv", "fy-profit-and-loss.csv"):
        url = f"https://f.irbank.net/files/{code}/{name}"
        resp = requests.get(url, headers=HEADERS, timeout=20)
        resp.raise_for_status()
        if b"<html" in resp.content[:500].lower():
            raise ValueError(f"IR BANK から CSV ではなく HTML が返されました: {url}")
        out[name] = resp.content
    return out["fy-balance-sheet.csv"], out["fy-profit-and-loss.csv"]


def fetch_stock_price(code):
    """Yahoo Finance v8 chart API で株価と銘柄名を取得する"""
    url = f"https://query1.finance.yahoo.com/v8/finance/chart/{code}.T?interval=1d&range=1d"
    resp = requests.get(url, headers=HEADERS, timeout=15)
    resp.raise_for_status()
    result = resp.json().get("chart", {}).get("result") or []
    if not result:
        raise ValueError("Yahoo Finance から株価データを取得できませんでした。")
    meta = result[0].get("meta", {})
    price = meta.get("regularMarketPrice")
    if price is None:
        raise ValueError("Yahoo Finance の応答に株価がありません。")
    name = meta.get("longName") or meta.get("shortName") or ""
    return float(price), name


def fetch_yahoo_timeseries(code, types, years=5):
    """Yahoo Finance fundamentals-timeseries から年次データの最新値を取得する"""
    now_ts = int(dt.datetime.now().timestamp())
    url = f"https://query1.finance.yahoo.com/ws/fundamentals-timeseries/v1/finance/timeseries/{code}.T"
    params = {"type": ",".join(types), "period1": now_ts - years * 365 * 86400, "period2": now_ts}
    resp = requests.get(url, params=params, headers=HEADERS, timeout=15)
    resp.raise_for_status()
    latest = {}
    for item in resp.json().get("timeseries", {}).get("result", []) or []:
        for t in types:
            values = [v for v in (item.get(t) or []) if v]
            if values:
                v = max(values, key=lambda x: x.get("asOfDate", ""))
                raw = v.get("reportedValue", {}).get("raw")
                if raw is not None:
                    latest[t] = float(raw)
    return latest


def fetch_current_assets_million(code):
    data = fetch_yahoo_timeseries(code, ["annualCurrentAssets"])
    ca = data.get("annualCurrentAssets")
    return int(round(ca / 1_000_000, 0)) if ca is not None else None


def inputs_from_yahoo(code):
    """IR BANK が使えない場合の代替：Yahoo Finance の年次財務データから同じ指標を組み立てる"""
    types = [
        "annualTotalAssets", "annualStockholdersEquity", "annualTotalEquityGrossMinorityInterest",
        "annualOrdinarySharesNumber", "annualShareIssued", "annualBasicEPS", "annualDilutedEPS",
        "annualNetIncome", "annualNetIncomeCommonStockholders", "annualTotalRevenue",
        "annualCurrentAssets",
    ]
    d = fetch_yahoo_timeseries(code, types)
    inp = FinancialInputs(code=code, source="Yahoo Finance（財務データから算出）")
    total_assets = d.get("annualTotalAssets")
    equity = d.get("annualStockholdersEquity")
    shares = d.get("annualOrdinarySharesNumber") or d.get("annualShareIssued")
    net_income = d.get("annualNetIncomeCommonStockholders") or d.get("annualNetIncome")
    if not (total_assets and equity and shares):
        raise ValueError("Yahoo Finance から総資産・自己資本・発行済株式数を取得できませんでした。")
    inp.total_assets = total_assets
    inp.net_assets = d.get("annualTotalEquityGrossMinorityInterest") or equity
    inp.shareholders_equity = equity
    inp.shares_outstanding = shares
    inp.bps = equity / shares
    inp.equity_ratio = equity / total_assets
    inp.eps = d.get("annualBasicEPS") or d.get("annualDilutedEPS") or (net_income / shares if net_income else None)
    inp.roa = net_income / total_assets if net_income is not None else None
    inp.sales = d.get("annualTotalRevenue")
    ca = d.get("annualCurrentAssets")
    inp.current_assets_million = int(round(ca / 1_000_000, 0)) if ca is not None else None
    inp.notes.append("IR BANK の CSV を取得できなかったため、Yahoo Finance の直近年次実績から算出しました"
                     "（EPS は実績値。IR BANK の予想 EPS とは異なる場合があります）。")
    if inp.eps is None or inp.roa is None:
        raise ValueError("Yahoo Finance から EPS / 純利益を取得できませんでした。")
    return inp


def load_inputs_by_code(code):
    """銘柄コードから入力データを取得（IR BANK → 失敗時 Yahoo Finance）"""
    code = normalize_code(code)
    errors = []
    try:
        bs, pl = fetch_irbank_csv(code)
        inp = parse_irbank_csv(bs, pl)
        inp.code = code
        inp.source = "IR BANK（CSV）"
        inp.raw_pl = pl
        return inp, errors
    except Exception as e:  # noqa: BLE001
        errors.append(f"IR BANK からの取得に失敗: {e}")
    inp = inputs_from_yahoo(code)
    return inp, errors


def enrich_with_market_data(inp):
    """株価・流動資産を Yahoo Finance から補完する。取得できなかった項目は警告として返す"""
    warnings = []
    if inp.stock_price is None and inp.code:
        try:
            price, name = fetch_stock_price(inp.code)
            inp.stock_price = price
            if not inp.company:
                inp.company = f"{inp.code} {name}".strip()
        except Exception as e:  # noqa: BLE001
            warnings.append(f"株価を取得できませんでした（{e}）。画面で株価を入力してください。")
    if inp.current_assets_million is None and inp.code:
        try:
            inp.current_assets_million = fetch_current_assets_million(inp.code)
            if inp.current_assets_million is None:
                warnings.append("流動資産を取得できませんでした。画面で入力すると「時価総額/正味価値」を計算します。")
        except Exception as e:  # noqa: BLE001
            warnings.append(f"流動資産を取得できませんでした（{e}）。画面で入力すると「時価総額/正味価値」を計算します。")
    if not inp.company:
        inp.company = inp.code
    return warnings


# ============================================================
# 計算（StockPrice_ver2.py セル4 と同じロジック）
# ============================================================
def calculate_e2(val):
    """割引評価率"""
    if val is None:
        return None
    if val >= 0.8:
        return 0.8
    elif val >= 0.67:
        return 0.75
    elif val >= 0.5:
        return 0.7
    elif val >= 0.33:
        return 0.65
    elif val >= 0.1:
        return 0.6
    return 0.5


def calculate_g4(val):
    """財務レバレッジ補正"""
    if val is None:
        return None
    if val >= 0.66:
        return 1
    elif val >= 0.5:
        return 1.2
    elif val >= 0.4:
        return 1.36
    return 1.5


@dataclass
class Result:
    discount_rate: float = None      # E2
    asset_value: float = None        # C2
    leverage: float = None           # G4
    business_value: float = None     # C4
    theoretical_price: float = None  # C6
    stock_price: float = None        # C8
    upside_pct: float = None         # C10 × 100
    market_cap_million: float = None # E12
    sales_million: float = None      # F12
    psr: float = None                # C12
    liabilities_million: float = None  # H14
    current_assets_million: float = None  # G14
    net_value_million: float = None  # F14
    mc_to_net_value: float = None    # C14
    equity_ok: bool = None           # G2
    upside_ok: bool = None           # D10
    psr_ok: bool = None              # D12
    mc_net_ok: bool = None           # D14
    recommend: bool = None
    advice: str = ""
    summary: str = ""


def calculate(inp):
    r = Result()
    ratio = inp.equity_ratio

    r.discount_rate = calculate_e2(ratio)
    r.asset_value = inp.bps * r.discount_rate if inp.bps is not None and r.discount_rate is not None else 0
    r.leverage = calculate_g4(ratio)
    if None not in (inp.eps, inp.roa, r.leverage):
        r.business_value = inp.eps * inp.roa * ROE_ASSUMED_PER * r.leverage
    else:
        r.business_value = 0
    r.theoretical_price = r.asset_value + r.business_value
    r.stock_price = inp.stock_price
    r.equity_ok = ratio > 0.5 if ratio is not None else None

    if r.stock_price:
        r.upside_pct = (r.theoretical_price / r.stock_price - 1) * 100
        r.upside_ok = r.upside_pct > UPSIDE_THRESHOLD

    # 時価総額 = 株価 × 発行済株式数（発行済株式数 ≈ 株主資本 / BPS）
    shares = inp.shares_outstanding
    if shares is None and inp.shareholders_equity and inp.bps:
        shares = inp.shareholders_equity / inp.bps
        inp.shares_outstanding = shares
    if r.stock_price and shares:
        r.market_cap_million = int(round(r.stock_price * shares / 1_000_000, 0))

    if inp.sales is not None:
        r.sales_million = int(round(inp.sales / 1_000_000, 0))
    if r.market_cap_million and r.sales_million:
        r.psr = r.market_cap_million / r.sales_million
        r.psr_ok = r.psr < 1

    if inp.total_assets is not None and inp.net_assets is not None:
        r.liabilities_million = int(round((inp.total_assets - inp.net_assets) / 1_000_000, 0))
    r.current_assets_million = inp.current_assets_million
    if r.current_assets_million is not None and r.liabilities_million is not None:
        r.net_value_million = r.current_assets_million - r.liabilities_million
        if r.net_value_million != 0 and r.market_cap_million:
            r.mc_to_net_value = r.market_cap_million / r.net_value_million
            r.mc_net_ok = 0 < r.mc_to_net_value < 1

    # 投資判断（元スクリプトと同じ：上昇余地 30% 以上で推奨）
    ratio_display = f"{ratio * 100:.1f}%" if ratio is not None else "-"
    if r.upside_pct is not None and r.psr is not None:
        r.recommend = r.upside_pct >= UPSIDE_THRESHOLD
        mark, advice = ("【★推奨】", "投資を推奨します") if r.recommend else ("【▲非推奨】", "投資を推奨しません（様子見）")
        r.advice = f"{mark} {advice}"
        r.summary = (f"自己資本比率 '{ratio_display}'、理論株価に対しての上昇余地 '{r.upside_pct:.1f}%'、"
                     f"PSRが'{r.psr:.1f}倍'であることから、{mark} {advice}。")
    else:
        r.recommend = None
        r.advice = "判断保留"
        r.summary = "必要なデータが不足しているため、投資判断は保留します。"
    return r


# ============================================================
# Excel 出力（テンプレート StockPrice.xlsx に書き込み）
# ============================================================
def build_excel(inp, r, template_path=TEMPLATE_PATH):
    wb = px.load_workbook(template_path)
    ws = wb.active

    ws["A1"] = inp.company or inp.code
    ws["D2"] = inp.bps
    ws["F2"] = inp.equity_ratio
    ws["F2"].number_format = "0.00%"
    ws["E2"] = r.discount_rate
    ws["C2"] = r.asset_value
    ws["D4"] = inp.eps
    ws["E4"] = inp.roa
    ws["G4"] = r.leverage
    ws["C4"] = r.business_value
    ws["C6"] = r.theoretical_price
    if r.stock_price is not None:
        ws["C8"] = r.stock_price
    if r.market_cap_million is not None:
        ws["E12"] = r.market_cap_million
    if inp.sales is not None:
        ws["G12"] = inp.sales
    ws["F12"] = r.sales_million if r.sales_million is not None else ""
    if r.liabilities_million is not None:
        ws["H14"] = r.liabilities_million
        ws["H14"].number_format = "#,##0"
    if r.current_assets_million is not None:
        ws["G14"] = r.current_assets_million
        ws["G14"].number_format = "#,##0"

    # C10/C12/C14 などの数式はテンプレートのまま残し、Excel で開いたときに再計算させる
    wb.calculation.fullCalcOnLoad = True

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def excel_file_name(inp):
    name = (inp.company or inp.code or "result").strip()
    name = re.sub(r'[\\/:*?"<>|]', "_", name)
    return f"StockPrice_{name}.xlsx"
