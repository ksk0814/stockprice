# -*- coding: utf-8 -*-
"""
理論株価・IR レポート Webアプリ（Streamlit）
起動: python3 -m streamlit run app.py
"""

import re

import altair as alt
import pandas as pd
import streamlit as st

import ir_report_core as ir
import stockprice_core as core

# 古い Streamlit（Python 3.8 で入る 1.3x 系など）でも動くよう、バージョンで引数を切り替える
ST_VER = tuple(int(x) for x in re.findall(r"\d+", st.__version__)[:2])
STRETCH = {"width": "stretch"} if ST_VER >= (1, 46) else {"use_container_width": True}
NUM_FMT = "localized" if ST_VER >= (1, 46) else "%d"

st.set_page_config(page_title="理論株価・IRレポート", page_icon="📈", layout="wide")

st.title("📈 理論株価・IR レポート")
st.caption("銘柄コードを入れるだけで、理論株価（資産価値＋事業価値）と業績の成長性・過去の PER 水準を並べて、割安か割高かを判断する材料をそろえます。")

for k, v in {"inputs": None, "report": None, "messages": [], "price": 0.0, "ca": 0.0, "theory_error": None}.items():
    st.session_state.setdefault(k, v)


# ============================================================
# 入力
# ============================================================
def load(code=None, bs_file=None, pl_file=None):
    msgs, inp, rep = [], None, None
    theory_error = None

    with st.spinner("財務データを取得しています…"):
        try:
            if code is not None:
                inp, m = core.load_inputs_by_code(code)
                msgs += m
            else:
                inp = core.parse_irbank_csv(bs_file.getvalue(), pl_file.getvalue())
                inp.raw_pl = pl_file.getvalue()
        except Exception as e:  # noqa: BLE001
            theory_error = f"理論株価の計算に必要なデータを読み込めませんでした: {e}"

    with st.spinner("株価・流動資産を Yahoo Finance から取得しています…"):
        if inp is not None:
            msgs += core.enrich_with_market_data(inp)

    with st.spinner("業績の推移と期末株価を取得しています…"):
        pl_bytes = inp.raw_pl if inp is not None else (pl_file.getvalue() if pl_file else None)
        report_code = code or (inp.code if inp else None)
        try:
            rep, m = ir.load_report(report_code, pl_data=pl_bytes)
            msgs += m
        except Exception as e:  # noqa: BLE001
            msgs.append(f"IR レポート（業績推移）を作成できませんでした: {e}")

    if inp is None and rep is None:
        st.error("データを取得できませんでした。銘柄コードを確認するか、CSV をアップロードしてください。")
        for m in [x for x in msgs + [theory_error] if x]:
            st.caption(m)
        return
    if rep is not None and inp is not None:
        rep.company = rep.company or inp.company
    st.session_state.update(
        inputs=inp, report=rep, messages=msgs, theory_error=theory_error,
        price=float((inp.stock_price if inp else None) or (rep.current_price if rep else None) or 0.0),
        ca=float((inp.current_assets_million if inp else None) or 0.0),
    )


mode = st.radio("データの入力方法", ["銘柄コードを入力", "CSV ファイルをアップロード"], horizontal=True,
                label_visibility="collapsed")
if mode == "銘柄コードを入力":
    with st.form("code_form"):
        c1, c2 = st.columns([3, 1], **({"vertical_alignment": "bottom"} if ST_VER >= (1, 36) else {}))
        code = c1.text_input("銘柄コード（4桁）", placeholder="例: 7203、130A")
        submitted = c2.form_submit_button("分析する", type="primary", **STRETCH)
    if submitted:
        try:
            load(code=core.normalize_code(code))
        except ValueError as e:
            st.error(str(e))
else:
    st.markdown("IR BANK からダウンロードした **fy-balance-sheet.csv** と **fy-profit-and-loss.csv** を選んでください。"
                "損益計算書だけでも IR レポートは作れます。")
    c1, c2 = st.columns(2)
    bs_file = c1.file_uploader("貸借対照表（fy-balance-sheet.csv）", type=["csv"])
    pl_file = c2.file_uploader("損益計算書（fy-profit-and-loss.csv）", type=["csv"])
    if st.button("分析する", type="primary", disabled=not pl_file):
        if bs_file:
            load(bs_file=bs_file, pl_file=pl_file)
        else:
            try:
                rep, m = ir.load_report(None, pl_data=pl_file.getvalue())
                st.session_state.update(inputs=None, report=rep, messages=m, price=float(rep.current_price or 0),
                                        theory_error="貸借対照表の CSV が無いため、理論株価は計算していません。")
            except Exception as e:  # noqa: BLE001
                st.error(f"CSV を読み込めませんでした: {e}")

inp, rep = st.session_state.inputs, st.session_state.report
if inp is None and rep is None:
    st.info("銘柄コードを入力して「分析する」を押してください。")
    st.stop()


# ============================================================
# 計算（株価は画面で修正できるので、その値で毎回再計算する）
# ============================================================
price = st.session_state.price or None
r = None
if inp is not None:
    inp.stock_price = price
    inp.current_assets_million = st.session_state.ca or None
    r = core.calculate(inp)
if rep is not None:
    ir.set_current_price(rep, price)
label, reasons = ir.combined_advice(r.upside_pct if r else None, rep, core.UPSIDE_THRESHOLD)


def fmt(v, spec, suffix=""):
    return "-" if v is None else f"{v:{spec}}{suffix}"


def judge(ok):
    return "-" if ok is None else ("OK" if ok else "NG")


XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
name = (inp.company if inp else "") or (rep.company if rep else "") or (inp.code if inp else rep.code)
st.subheader(name)

for m in st.session_state.messages:
    st.warning(m)
if st.session_state.theory_error:
    st.warning(st.session_state.theory_error)
if not price:
    st.info("現在の株価を取得できませんでした。下の「現在の株価」に入力すると、上昇余地や PER を計算します。")
st.columns([1, 2])[0].number_input("現在の株価（円）", min_value=0.0, step=1.0, key="price",
                                    help="自動取得した値です。指値などで試算したいときは書き換えると全体が再計算されます")

# ---- 要点（どのタブでも見える） ----
k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("現在の株価", fmt(price, ",.0f", " 円"))
k2.metric("理論株価", fmt(r.theoretical_price if r else None, ",.0f", " 円"),
          delta=fmt(r.upside_pct, "+.1f", "%") if r and r.upside_pct is not None else None,
          help="資産価値＋事業価値。下の数字は現在の株価に対する上昇余地")
k3.metric("PER 水準から見た目安", fmt(rep.per_fair_price if rep else None, ",.0f", " 円"),
          delta=fmt(rep.per_upside_pct, "+.1f", "%") if rep and rep.per_upside_pct is not None else None,
          help="過去の期末 PER の中央値 × 直近 EPS（予想があれば予想）")
k4.metric("成長性（直近3期）", rep.growth_label if rep else "-",
          help=f"直近3期平均の 増収率＋増益率。{ir.GROWTH_HIGH:.0f}% 以上かつ両方プラスで「高成長」")
verdict = label.split("・")[0]
k5.metric("総合評価", verdict, help=f"{label}（詳しくは「総合判断」タブ）")

tab_sum, tab_theory, tab_ir = st.tabs(["🧭 総合判断", "💰 理論株価", "📊 業績・成長性（IR レポート）"])

# ============================================================
# 総合判断
# ============================================================
with tab_sum:
    box = st.success if "割安" in label else st.error if "割高" in label else st.info
    box(f"**総合評価：{label}**")
    for t in reasons:
        st.markdown(f"- {t}")

    st.markdown("#### 判断の目安")
    rows = []
    if r is not None:
        rows.append(["理論株価との比較（上昇余地）", fmt(r.upside_pct, ".1f", "%"), f"{core.UPSIDE_THRESHOLD}% 以上で割安",
                     judge(r.upside_ok)])
        rows.append(["自己資本比率", f"{inp.equity_ratio * 100:.1f}%", "50% 超で安全性が高い", judge(r.equity_ok)])
        rows.append(["PSR", fmt(r.psr, ".2f", " 倍"), "1.0 倍未満で割安", judge(r.psr_ok)])
    if rep is not None:
        per_judge = "-" if rep.per_upside_pct is None else (
            "割安" if rep.per_upside_pct >= 10 else "割高" if rep.per_upside_pct <= -10 else "同水準")
        rows.append(["過去 PER との比較", fmt(rep.current_per, ".1f", " 倍"),
                     f"中央値 {fmt(rep.per_median, '.1f', ' 倍')} より低ければ割安", per_judge])
        rows.append(["成長性（直近3期平均）", fmt(rep.growth_score_3y, ".1f", "%"), f"{ir.GROWTH_HIGH:.0f}% 以上で高成長",
                     rep.growth_label or "-"])
        rows.append(["EPS 年平均成長率", fmt(rep.eps_cagr, ".1f", "%"),
                     f"過去 {rep.eps_cagr_years} 年" if rep.eps_cagr_years else "-", "-"])
        corr = rep.eps_price_corr
        rows.append(["EPS と株価の相関", fmt(corr, ".2f"), "0.7 以上で業績連動性が高い",
                     "-" if corr is None else ("高い" if corr >= 0.7 else "中程度" if corr >= 0.3 else "低い")])
    st.dataframe(pd.DataFrame(rows, columns=["観点", "値", "目安", "判定"]), hide_index=True, **STRETCH)
    st.caption("理論株価は資産と現在の収益力、PER 水準は市場がこれまでつけてきた評価、成長性は今後の利益の伸びを見ています。"
               "3 つがそろって割安を示すほど確度が高く、食い違うときは理由（一時的な減益、業態の変化など）を確認してください。"
               "あくまで助言であり、投資は自己責任でお願いします。")

    if r is not None and rep is not None:
        st.download_button("📥 まとめて Excel（理論株価＋IR レポート）", type="primary",
                           data=ir.add_ir_sheet_to_workbook(core.build_excel(inp, r), rep),
                           file_name=core.excel_file_name(inp), mime=XLSX)

# ============================================================
# 理論株価
# ============================================================
with tab_theory:
    if inp is None:
        st.info(st.session_state.theory_error or "理論株価の計算に必要なデータがありません。")
    else:
        for n in inp.notes:
            st.caption(f"ℹ️ {n}")
        c1, c2 = st.columns(2)
        c1.number_input("流動資産（百万円）", step=1.0, key="ca", help="0 のときは「時価総額/正味価値」を計算しません")
        c2.text_input("データ元", value=inp.source, disabled=True)

        m1, m2, m3 = st.columns(3)
        m1.metric("資産価値", fmt(r.asset_value, ",.1f", " 円"), help="BPS × 割引評価率")
        m2.metric("事業価値", fmt(r.business_value, ",.1f", " 円"), help=f"EPS × ROA × {core.ROE_ASSUMED_PER} × 財務レバレッジ補正")
        m3.metric("時価総額", fmt(r.market_cap_million, ",.0f", " 百万円"))

        st.markdown("#### 指標と判定")
        st.dataframe(pd.DataFrame({
            "指標": ["① 自己資本比率", "② 上昇余地", "③ PSR", "④ 時価総額/正味価値"],
            "値": [f"{inp.equity_ratio * 100:.1f}%", fmt(r.upside_pct, ".1f", "%"), fmt(r.psr, ".2f", " 倍"),
                  fmt(r.mc_to_net_value, ".2f", " 倍")],
            "基準": ["50% 超", f"{core.UPSIDE_THRESHOLD}% 超", "1.0 倍未満", "1.0 倍未満"],
            "判定": [judge(r.equity_ok), judge(r.upside_ok), judge(r.psr_ok), judge(r.mc_net_ok)],
        }), hide_index=True, **STRETCH)

        st.markdown("#### 理論株価だけで見た投資判断")
        (st.success if r.recommend else st.error if r.recommend is False else st.warning)(r.summary)

        with st.expander("計算の内訳"):
            st.markdown(f"""
- 割引評価率（自己資本比率から）: **{r.discount_rate}**
- 資産価値 = BPS {inp.bps:,.2f} × {r.discount_rate} = **{r.asset_value:,.2f} 円**
- 財務レバレッジ補正: **{r.leverage}**
- 事業価値 = EPS {inp.eps:,.2f} × ROA {inp.roa:.4f} × {core.ROE_ASSUMED_PER} × {r.leverage} = **{r.business_value:,.2f} 円**
- 理論株価 = 資産価値 + 事業価値 = **{r.theoretical_price:,.2f} 円**
- 売上高: {fmt(r.sales_million, ",.0f", " 百万円")}／負債（総資産−純資産）: {fmt(r.liabilities_million, ",.0f", " 百万円")}／流動資産: {fmt(r.current_assets_million, ",.0f", " 百万円")}／正味価値: {fmt(r.net_value_million, ",.0f", " 百万円")}
""")
        st.download_button("📥 理論株価の Excel（StockPrice テンプレート）", data=core.build_excel(inp, r),
                           file_name=core.excel_file_name(inp), mime=XLSX)

# ============================================================
# IR レポート
# ============================================================
with tab_ir:
    if rep is None:
        st.info("業績の推移データを取得できませんでした。")
    else:
        for n in rep.notes:
            st.caption(f"ℹ️ {n}")
        st.caption(f"データ元：{rep.source}。色の薄い棒と「予想」の行は会社予想です。増収率・増益率（営業利益）は前期比です。")

        df = pd.DataFrame([{
            "年度": x.label + ("（予想）" if x.forecast and "予" not in x.label else ""),
            "売上高": x.sales / 1e6 if x.sales is not None else None,
            "営業利益": x.op_income / 1e6 if x.op_income is not None else None,
            "増収率": x.sales_growth, "増益率": x.op_growth, "成長性": x.growth_score, "区分": x.pattern or "-",
            "純利益": x.net_income / 1e6 if x.net_income is not None else None,
            "EPS": x.eps, "期末株価": x.price, "PER": x.per, "予想": x.forecast,
        } for x in rep.rows])
        # 欠損（None）が混じっても数値列として扱えるようにする（None → NaN）
        num_cols = ["売上高", "営業利益", "増収率", "増益率", "成長性", "純利益", "EPS", "期末株価", "PER"]
        df[num_cols] = df[num_cols].apply(pd.to_numeric, errors="coerce").astype(float)

        g1, g2, g3, g4 = st.columns(4)
        g1.metric("直近3期平均 増収率", fmt(rep.avg_sales_growth_3y, ".1f", "%"))
        g2.metric("直近3期平均 増益率", fmt(rep.avg_op_growth_3y, ".1f", "%"))
        g3.metric("EPS 年平均成長率", fmt(rep.eps_cagr, ".1f", "%"),
                  help=f"過去 {rep.eps_cagr_years} 年" if rep.eps_cagr_years else None)
        g4.metric("EPS と株価の相関", fmt(rep.eps_price_corr, ".2f"))

        order = list(df["年度"])
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**売上高と営業利益の推移（百万円）**")
            long = df.melt(id_vars=["年度", "予想"], value_vars=["売上高", "営業利益"], var_name="項目",
                           value_name="百万円").dropna(subset=["百万円"])
            st.altair_chart(alt.Chart(long).mark_bar().encode(
                x=alt.X("年度:N", sort=order, title=None),
                xOffset="項目:N",
                y=alt.Y("百万円:Q", title=None),
                color=alt.Color("項目:N", legend=alt.Legend(orient="top", title=None)),
                opacity=alt.condition("datum.予想", alt.value(0.4), alt.value(1)),
                tooltip=["年度", "項目", alt.Tooltip("百万円:Q", format=",.0f")],
            ), **STRETCH)
        with c2:
            st.markdown("**増収率・増益率（前期比 %）**")
            long = df.melt(id_vars=["年度"], value_vars=["増収率", "増益率"], var_name="項目",
                           value_name="率").dropna(subset=["率"])
            st.altair_chart(alt.Chart(long).mark_line(point=True).encode(
                x=alt.X("年度:N", sort=order, title=None),
                y=alt.Y("率:Q", title="%"),
                color=alt.Color("項目:N", legend=alt.Legend(orient="top", title=None)),
                tooltip=["年度", "項目", alt.Tooltip("率:Q", format=".1f")],
            ) + alt.Chart(pd.DataFrame({"y": [0]})).mark_rule(color="gray").encode(y="y:Q"),
                **STRETCH)

        c3, c4 = st.columns(2)
        pts = df.dropna(subset=["EPS", "期末株価"])
        with c3:
            st.markdown("**EPS と期末株価の相関**")
            if pts.empty:
                st.caption("期末株価が無いため表示できません。")
            else:
                base = alt.Chart(pts).encode(x=alt.X("EPS:Q", title="EPS（円）"),
                                             y=alt.Y("期末株価:Q", title="株価（円）"))
                st.altair_chart(base.mark_circle(size=80)
                                + base.mark_text(dx=6, align="left", fontSize=10).encode(text="年度")
                                + base.transform_regression("EPS", "期末株価").mark_line(strokeDash=[4, 4], color="gray"),
                                **STRETCH)
        with c4:
            st.markdown("**期末 PER の推移**")
            per = df.dropna(subset=["PER"])
            if per.empty:
                st.caption("PER を計算できる年度がありません。")
            else:
                chart = alt.Chart(per).mark_line(point=True).encode(
                    x=alt.X("年度:N", sort=order, title=None), y=alt.Y("PER:Q", title="倍"),
                    tooltip=["年度", alt.Tooltip("PER:Q", format=".1f")])
                lines = [(n, v) for n, v in (("過去の中央値", rep.per_median), ("現在", rep.current_per)) if v]
                rules = pd.DataFrame({"名称": [n for n, _ in lines], "y": [v for _, v in lines]})
                st.altair_chart(chart + alt.Chart(rules).mark_rule(strokeDash=[4, 4]).encode(
                    y="y:Q", color=alt.Color("名称:N", scale=alt.Scale(domain=["過去の中央値", "現在"], range=["#888888", "#e8590c"]),
                                             legend=alt.Legend(orient="top", title=None))),
                    **STRETCH)

        st.markdown("**年度別データ**")
        table = df.drop(columns=["予想"])
        for c in ("売上高", "営業利益", "純利益", "期末株価"):
            table[c] = table[c].round(0)
        st.dataframe(
            table, hide_index=True, **STRETCH,
            column_config={
                "売上高": st.column_config.NumberColumn("売上高（百万円）", format=NUM_FMT),
                "営業利益": st.column_config.NumberColumn("営業利益（百万円）", format=NUM_FMT),
                "純利益": st.column_config.NumberColumn("純利益（百万円）", format=NUM_FMT),
                "増収率": st.column_config.NumberColumn(format="%.1f%%"),
                "増益率": st.column_config.NumberColumn(format="%.1f%%"),
                "成長性": st.column_config.NumberColumn("成長性（増収率＋増益率）", format="%.1f%%"),
                "EPS": st.column_config.NumberColumn(format="%.1f"),
                "期末株価": st.column_config.NumberColumn(format=NUM_FMT),
                "PER": st.column_config.NumberColumn(format="%.1f"),
            })
        st.download_button("📥 IR レポートの Excel", data=ir.build_ir_excel(rep), file_name=ir.ir_file_name(rep), mime=XLSX)
