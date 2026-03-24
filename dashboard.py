"""Wantedly Analytics ダッシュボード（Google Sheets から読み込み）"""

import json
from datetime import datetime, timedelta

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from google.oauth2.service_account import Credentials
import gspread

st.set_page_config(
    page_title="Wantedly Analytics",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

SHEET_ID = "1Xz3VSILhUrcP_wm5Wi4URqklSEvEzzM0x9fJkAZuGhE"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]


# ── Google Sheets 接続 ────────────────────
@st.cache_resource(ttl=300)
def get_client():
    if "gcp_service_account" in st.secrets:
        creds = Credentials.from_service_account_info(
            st.secrets["gcp_service_account"], scopes=SCOPES
        )
    else:
        # ローカル実行時は credentials.json を使用
        from pathlib import Path
        creds_path = Path(__file__).parent / "credentials.json"
        creds = Credentials.from_service_account_file(str(creds_path), scopes=SCOPES)
    return gspread.authorize(creds)


@st.cache_data(ttl=300)
def load_sheet(sheet_name: str) -> pd.DataFrame:
    client = get_client()
    ss = client.open_by_key(SHEET_ID)
    ws = ss.worksheet(sheet_name)
    data = ws.get_all_values()
    if not data or len(data) < 2:
        return pd.DataFrame()
    return pd.DataFrame(data[1:], columns=data[0])


# ── スタイル ──────────────────────────────
st.markdown("""
<style>
.metric-card {
    background: #f8f9fa;
    border-radius: 10px;
    padding: 16px 20px;
    border-left: 4px solid #2e5899;
    margin-bottom: 8px;
}
.metric-title { font-size: 13px; color: #666; margin-bottom: 4px; }
.metric-value { font-size: 28px; font-weight: bold; color: #2e5899; }
.metric-sub   { font-size: 12px; color: #999; }
</style>
""", unsafe_allow_html=True)


def metric_card(title: str, value: str, sub: str = "") -> None:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-title">{title}</div>
        <div class="metric-value">{value}</div>
        <div class="metric-sub">{sub}</div>
    </div>
    """, unsafe_allow_html=True)


# ── メイン ────────────────────────────────
def main() -> None:
    st.title("📊 Wantedly Analytics ダッシュボード")
    st.caption(f"データ更新: {datetime.now().strftime('%Y-%m-%d %H:%M')} ／ 5分キャッシュ")

    # データ読み込み
    with st.spinner("データを読み込み中..."):
        summary_df = load_sheet("サマリー")
        pv_df      = load_sheet("PV推移")
        oubo_df    = load_sheet("応募推移")

    if summary_df.empty:
        st.error("データがありません。scraper.py を実行してください。")
        return

    # 数値列を変換
    for col in ["累計PV", "ブックマーク", "応募数", "応援数", "いいね"]:
        if col in summary_df.columns:
            summary_df[col] = pd.to_numeric(summary_df[col], errors="coerce").fillna(0).astype(int)

    boshu_df = summary_df[summary_df["種別"] == "募集"]
    story_df = summary_df[summary_df["種別"] == "ストーリー"]

    # ── サイドバー ────────────────────────
    with st.sidebar:
        st.header("🔍 フィルター")
        view = st.radio("種別", ["すべて", "募集", "ストーリー"])
        keyword = st.text_input("記事名で絞り込み")
        st.divider()
        st.caption("データは5分キャッシュされます。\n最新データを見るには下のボタンを押してください。")
        if st.button("🔄 キャッシュをクリア", use_container_width=True):
            st.cache_data.clear()
            st.rerun()

    # ── KPIカード ─────────────────────────
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1: metric_card("追跡記事数", str(len(summary_df)), f"募集 {len(boshu_df)} / ストーリー {len(story_df)}")
    with c2: metric_card("募集 累計PV合計", f"{boshu_df['累計PV'].sum():,}")
    with c3: metric_card("募集 累計応募合計", f"{boshu_df['応募数'].sum():,}")
    with c4: metric_card("ストーリー 累計PV合計", f"{story_df['累計PV'].sum():,}")
    with c5: metric_card("ストーリー いいね合計", f"{story_df['いいね'].sum():,}")

    st.divider()

    # ── タブ ──────────────────────────────
    tab1, tab2, tab3, tab4 = st.tabs(["📋 サマリー", "📈 PV推移", "📬 応募推移", "🔍 記事詳細"])

    # ──────────────────────────────────────
    # タブ1: サマリー
    # ──────────────────────────────────────
    with tab1:
        df = summary_df.copy()
        if view == "募集":       df = boshu_df.copy()
        elif view == "ストーリー": df = story_df.copy()
        if keyword: df = df[df["タイトル"].str.contains(keyword, na=False)]

        st.subheader(f"記事一覧（{len(df)}件）")

        # 棒グラフ: 上位20記事
        top20 = df.nlargest(20, "累計PV")
        fig = px.bar(
            top20, x="累計PV", y="タイトル", orientation="h",
            color="状態",
            color_discrete_map={"募集中": "#2e5899", "募集停止中": "#aaa", "": "#57a65e"},
            title="PV上位20記事",
            labels={"タイトル": "", "累計PV": "累計PV"},
        )
        fig.update_layout(height=500, yaxis={"categoryorder": "total ascending"},
                          showlegend=True)
        st.plotly_chart(fig, use_container_width=True)

        # テーブル
        disp_cols = [c for c in ["種別","タイトル","状態","累計PV","ブックマーク","応募数","応援数","いいね","初回取得日","最終取得日"]
                     if c in df.columns]
        st.dataframe(
            df[disp_cols].sort_values("累計PV", ascending=False),
            use_container_width=True,
            hide_index=True,
            column_config={
                "累計PV": st.column_config.ProgressColumn("累計PV", max_value=int(df["累計PV"].max()) or 1),
                "応募数": st.column_config.NumberColumn("応募数"),
            },
        )

    # ──────────────────────────────────────
    # タブ2: PV推移
    # ──────────────────────────────────────
    with tab2:
        st.subheader("PV推移（上位30記事）")

        if pv_df.empty or len(pv_df) < 2:
            st.info("2回以上スクレイプを実行するとグラフが表示されます。")
        else:
            date_col = pv_df.columns[0]
            pv_melt = pv_df.melt(id_vars=date_col, var_name="記事", value_name="PV")
            pv_melt["PV"] = pd.to_numeric(pv_melt["PV"], errors="coerce").fillna(0)
            pv_melt = pv_melt.rename(columns={date_col: "日付"})

            top_n = st.slider("表示記事数", 3, 30, 10, key="pv_n")
            top_titles = (
                pv_melt.groupby("記事")["PV"].max()
                .nlargest(top_n).index.tolist()
            )
            fig = px.line(
                pv_melt[pv_melt["記事"].isin(top_titles)],
                x="日付", y="PV", color="記事",
                markers=True, title=f"PV推移（上位{top_n}記事）",
            )
            fig.update_layout(height=450, hovermode="x unified",
                              legend=dict(font=dict(size=10)))
            st.plotly_chart(fig, use_container_width=True)

    # ──────────────────────────────────────
    # タブ3: 応募推移
    # ──────────────────────────────────────
    with tab3:
        st.subheader("応募数推移（上位30記事）")

        if oubo_df.empty or len(oubo_df) < 2:
            st.info("2回以上スクレイプを実行するとグラフが表示されます。")
        else:
            date_col = oubo_df.columns[0]
            oubo_melt = oubo_df.melt(id_vars=date_col, var_name="記事", value_name="応募数")
            oubo_melt["応募数"] = pd.to_numeric(oubo_melt["応募数"], errors="coerce").fillna(0)
            oubo_melt = oubo_melt.rename(columns={date_col: "日付"})

            top_n2 = st.slider("表示記事数", 3, 30, 10, key="oubo_n")
            top_titles2 = (
                oubo_melt.groupby("記事")["応募数"].max()
                .nlargest(top_n2).index.tolist()
            )
            fig2 = px.area(
                oubo_melt[oubo_melt["記事"].isin(top_titles2)],
                x="日付", y="応募数", color="記事",
                title=f"応募数推移（上位{top_n2}記事）",
            )
            fig2.update_layout(height=450, hovermode="x unified",
                               legend=dict(font=dict(size=10)))
            st.plotly_chart(fig2, use_container_width=True)

    # ──────────────────────────────────────
    # タブ4: 記事詳細
    # ──────────────────────────────────────
    with tab4:
        st.subheader("記事を選んで詳細確認")

        all_titles = summary_df["タイトル"].tolist()
        selected = st.selectbox("記事を選択", all_titles)

        if selected:
            row = summary_df[summary_df["タイトル"] == selected].iloc[0]

            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("状態",       row.get("状態", "-") or "ストーリー")
            c2.metric("初回取得日", row.get("初回取得日", "-"))
            c3.metric("最終取得日", row.get("最終取得日", "-"))
            c4.metric("累計PV",     f"{int(row['累計PV']):,}")
            if row["種別"] == "募集":
                c5.metric("応募数", f"{int(row['応募数']):,}")
            else:
                c5.metric("いいね", f"{int(row['いいね']):,}")

            st.divider()
            st.caption("※ 推移グラフは PV推移シートのデータを参照します。2回以上スクレイプ後に表示されます。")

            short_title = selected[:40]

            # PV推移グラフ（記事詳細）
            if not pv_df.empty and short_title in pv_df.columns:
                date_col = pv_df.columns[0]
                detail = pv_df[[date_col, short_title]].copy()
                detail.columns = ["日付", "累計PV"]
                detail["累計PV"] = pd.to_numeric(detail["累計PV"], errors="coerce")
                detail["日別PV"] = detail["累計PV"].diff().clip(lower=0)

                col_a, col_b = st.columns(2)
                with col_a:
                    fig_pv = px.line(detail, x="日付", y="累計PV", markers=True,
                                     title="累計PV推移")
                    st.plotly_chart(fig_pv, use_container_width=True)
                with col_b:
                    fig_daily = px.bar(detail, x="日付", y="日別PV",
                                       title="日別PV（前回比）",
                                       color_discrete_sequence=["#2e5899"])
                    st.plotly_chart(fig_daily, use_container_width=True)

            # 応募推移グラフ（記事詳細）
            if row["種別"] == "募集" and not oubo_df.empty and short_title in oubo_df.columns:
                date_col = oubo_df.columns[0]
                detail2 = oubo_df[[date_col, short_title]].copy()
                detail2.columns = ["日付", "応募数"]
                detail2["応募数"] = pd.to_numeric(detail2["応募数"], errors="coerce")
                fig_oubo = px.bar(detail2, x="日付", y="応募数",
                                  title="応募数推移",
                                  color_discrete_sequence=["#57a65e"])
                st.plotly_chart(fig_oubo, use_container_width=True)


if __name__ == "__main__":
    main()
