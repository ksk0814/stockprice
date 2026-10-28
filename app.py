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
st.caption("銘柄コードを入れるだけで、理論株価（残余利益モデル）と業績の成長性・過去の PER 水準を並べて、割安か割高かを判断する材料をそろえます。")
st.caption(f"⚠️ {core.DISCLAIMER}")

EXAMPLE_PRICE = core.residual_income_value(1000, 0.12)[0]
with st.expander("💡 理論株価の考え方（かんたん解説）", expanded=True):
    st.markdown(f"""
理論株価は **「会社が今持っている財産」＋「これから稼ぐ“もうけの上乗せ分”」** で計算します（残余利益モデル）。

1. **今持っている財産**：1 株あたりの純資産（BPS）。会社を今たたんだら株主に残るお金の目安です。
2. **もうけの上乗せ分**：株主は預けたお金に年 **{core.COST_OF_EQUITY * 100:.0f}%** くらいの利回りを期待しています（株主資本コスト）。
   会社がそれ **以上** に稼げば超えた分だけ価値が上乗せされ、届かなければ差し引かれます。
3. **控えめに見積もる**：高いもうけはライバルとの競争でだんだん薄れるものとして、**{core.FADE_YEARS} 年**かけて上乗せ分が 0 になると仮定し、
   将来のお金は今の価値に割り引いて合計します。

**例**：純資産 1,000 円・利益率（ROE）12% の会社は、期待の {core.COST_OF_EQUITY * 100:.0f}% を上回る分として 1 年目に約 40 円を上乗せで稼ぎます。
この上乗せが少しずつ小さくなりながら続くと考え、今の価値に直して合計すると約 {EXAMPLE_PRICE - 1000:,.0f} 円。
純資産と足した **約 {EXAMPLE_PRICE:,.0f} 円** が理論株価です。

株価が理論株価より **{core.UPSIDE_THRESHOLD}% 以上安ければ「割安圏」、{core.UPSIDE_THRESHOLD}% 以上高ければ「割高圏」** と表示します。
前提（{core.COST_OF_EQUITY * 100:.0f}% や {core.FADE_YEARS} 年）は「理論株価」タブで変えられ、結果にどのくらいの幅があるかも確認できます。
""")

DEFAULTS = {"inputs": None, "report": None, "messages": [], "price": 0.0, "ca": 0.0, "theory_error": None,
            "coe": core.COST_OF_EQUITY * 100, "years": core.FADE_YEARS, "payout": core.DEFAULT_PAYOUT * 100,
            "roe_basis": core.ROE_BASIS_FORECAST}
for k, v in DEFAULTS.items():
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
        payout=float((inp.payout_ratio if inp and inp.payout_ratio is not None else core.DEFAULT_PAYOUT) * 100),
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
    r = core.calculate(inp, coe=st.session_state.coe / 100, years=st.session_state.years,
                       payout=st.session_state.payout / 100, roe_basis=st.session_state.roe_basis)
if rep is not None:
    ir.set_current_price(rep, price)
label, reasons = ir.combined_advice(r.gap_pct if r else None, rep, core.UPSIDE_THRESHOLD)


def fmt(v, spec, suffix=""):
    return "-" if v is None else f"{v:{spec}}{suffix}"


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
          help="残余利益モデル（BPS ＋ 資本コストを超える利益の現在価値）。下の数字は理論株価が現在の株価をどれだけ上回るか")
k3.metric("PER 水準から見た目安", fmt(rep.per_fair_price if rep else None, ",.0f", " 円"),
          delta=fmt(rep.per_upside_pct, "+.1f", "%") if rep and rep.per_upside_pct is not None else None,
          help="過去の期末 PER の中央値 × 直近 EPS（予想があれば予想）")
k4.metric("成長性（直近3期）", rep.growth_label if rep else "-",
          help=f"直近3期平均の 増収率＋増益率。{ir.GROWTH_HIGH:.0f}% 以上かつ両方プラスで「高成長」")
verdict = label.split("（")[0]
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
        rows.append(["株価の理論株価からの乖離", fmt(r.gap_pct, "+.1f", "%"),
                     f"−{core.UPSIDE_THRESHOLD}% 以下で割安圏、+{core.UPSIDE_THRESHOLD}% 以上で割高圏", r.verdict or "-"])
        rows.append(["PBR", fmt(r.pbr, ".2f", " 倍"), "1.0 倍未満は株価が純資産を下回る", "-"])
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
    st.caption("理論株価は純資産と資本コストを上回る収益力、PER 水準は市場がこれまでつけてきた評価、成長性は今後の利益の伸びを見ています。"
               "3 つがそろうほど確度が高く、食い違うときは理由（一時的な損益、業態の変化など）を確認してください。")
    st.caption(f"⚠️ {core.DISCLAIMER}")

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
        for n in inp.notes + r.notes:
            st.caption(f"ℹ️ {n}")

        st.markdown("#### 前提条件")
        p1, p2, p3, p4 = st.columns(4)
        p1.number_input("株主資本コスト r（%）", min_value=1.0, max_value=20.0, step=0.5, key="coe",
                        help="株主が求める最低限の利回り。既定の 8% は伊藤レポート（経済産業省, 2014）が示した ROE の最低ライン")
        p2.number_input("ROE が r に収れんする年数", min_value=1, max_value=30, step=1, key="years",
                        help="今の ROE（超過利益）が競争で薄れ、r と同じになるまでの年数。長いほど理論株価は ROE の影響を強く受けます")
        p3.number_input("配当性向（%）", min_value=0.0, max_value=100.0, step=5.0, key="payout",
                        help="利益のうち配当に回す割合。残りは純資産に積み上がります。取得できない場合は 30% と仮定")
        p4.selectbox("1年目の ROE", core.ROE_BASES, key="roe_basis",
                     help="一時的な損益で ROE がぶれる銘柄は「過去3期平均 ROE」を選ぶと安定します")
        c1, c2 = st.columns(2)
        c1.number_input("流動資産（百万円）", step=1.0, key="ca", help="0 のときは「時価総額/正味流動資産」を計算しません")
        c2.text_input("データ元", value=inp.source, disabled=True)

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("BPS（純資産）", fmt(r.book_value, ",.1f", " 円"), help="1 株当たり純資産（直近実績）")
        m2.metric("残余利益の現在価値", fmt(r.pv_residual_income, ",.1f", " 円"),
                  help="資本コストを上回る（下回る）利益の合計を現在価値に割り引いたもの")
        m3.metric("理論株価", fmt(r.theoretical_price, ",.0f", " 円"),
                  help=f"株主資本コスト ±1% での幅：{fmt(r.price_low, ',.0f')}〜{fmt(r.price_high, ',.0f')} 円")
        m4.metric("ROE（1年目）", fmt(r.roe_used * 100 if r.roe_used is not None else None, ".1f", "%"), help=r.roe_basis)

        st.markdown("#### 株価との比較")
        (st.success if r.recommend else st.error if r.recommend is False else st.info)(r.summary)

        if r.sensitivity:
            st.markdown("#### 前提を変えたときの理論株価（円）")
            st.dataframe(pd.DataFrame(
                [[f"{c * 100:.0f}%"] + [round(r.sensitivity[(c, y)]) if r.sensitivity[(c, y)] is not None else None
                                        for y in core.SENS_YEARS] for c in core.SENS_COE],
                columns=["株主資本コスト ＼ 収れん年数"] + [f"{y} 年" for y in core.SENS_YEARS]), hide_index=True, **STRETCH)
            st.caption("理論株価は前提によって変わります。1 つの数字ではなく、この幅の中のどこに株価があるかを見てください。")

        st.markdown("#### 参考指標")
        st.dataframe(pd.DataFrame({
            "指標": ["PBR", "PER", "PSR", "自己資本比率", "時価総額 ÷ 正味流動資産"],
            "値": [fmt(r.pbr, ".2f", " 倍"), fmt(r.per, ".1f", " 倍"), fmt(r.psr, ".2f", " 倍"),
                  fmt(inp.equity_ratio * 100 if inp.equity_ratio is not None else None, ".1f", "%"),
                  fmt(r.mc_to_net_value, ".2f", " 倍")],
            "読み方": ["1 倍未満は株価が純資産を下回る（東証が改善を要請している水準）", "利益の何年分の株価か。業種平均と比べる",
                     "売上高に対する株価水準", "高いほど財務の安全性が高い",
                     "1 倍未満は清算価値より安い水準（Graham のネットネット）"],
        }), hide_index=True, **STRETCH)

        with st.expander("計算の内訳と算定方法"):
            st.markdown(f"""
**理論株価 = BPS + Σ (ROE_t − r) × 期首 BPS_t ÷ (1 + r)^t**（残余利益モデル）

- BPS（直近実績）: **{fmt(inp.bps, ",.2f", " 円")}**
- 1年目の ROE: **{fmt(r.roe_used * 100 if r.roe_used is not None else None, ".2f", "%")}**（{r.roe_basis}）
- 株主資本コスト r: **{r.cost_of_equity * 100:.1f}%**、収れん年数 T: **{r.years} 年**、配当性向: **{r.payout * 100:.1f}%**
- ROE は T 年かけて r まで直線的に下がり、それ以降の超過利益は 0 と置いています（保守的な前提）。
- 純資産は「利益 − 配当」だけ毎年増えると仮定しています（クリーンサープラス関係）。
""")
            if r.ri_rows:
                st.dataframe(pd.DataFrame([{
                    "年": x["t"], "ROE": f"{x['roe'] * 100:.2f}%", "期首 BPS": round(x["book"], 2),
                    "残余利益": round(x["ri"], 2), "割引係数": round(x["df"], 4), "現在価値": round(x["pv"], 2),
                } for x in r.ri_rows]), hide_index=True, **STRETCH)
            st.markdown("""
**この手法を使う理由**：会計上の純資産と利益だけで企業価値を表せる理論式で（Ohlson 1995）、
将来の大部分を見積もりに頼る DCF や配当割引モデルより、実際の株価をよく説明することが実証研究で示されています
（Francis, Olsson & Oswald 2000 ほか）。日本公認会計士協会『企業価値評価ガイドライン』にもインカム・アプローチの一つ（残余利益法）として掲載されています。

**出典**
""" + "\n".join(f"- {s}" for s in core.METHOD_SOURCES))
        st.download_button("📥 理論株価の Excel（計算式入り）", data=core.build_excel(inp, r),
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
