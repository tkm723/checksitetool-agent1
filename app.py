#!/usr/bin/env python3
# coding: utf-8
"""
医療クリニック Webサイト検品ツール — Streamlit Web UI

デプロイ: Streamlit Community Cloud (https://streamlit.io/cloud)
  1. このリポジトリをGitHubにpush
  2. Streamlit CloudでNew Appから接続
  3. Settings → Secrets に ANTHROPIC_API_KEY を設定（LLMモード使用時）
"""

import importlib.util
import io
import os
import sys
import tempfile
from collections import Counter
from datetime import datetime

import pandas as pd
import requests
import streamlit as st
from bs4 import BeautifulSoup

# ── checker.py を読み込む ──
_here = os.path.dirname(os.path.abspath(__file__))
_spec = importlib.util.spec_from_file_location(
    "checker", os.path.join(_here, "checker.py")
)
ck = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(ck)

# ── ページ設定 ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="医療クリニック Webサイト検品",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── スタイル ──────────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* 課題カード */
.card-ng {
    background: #fff0f0;
    border-left: 4px solid #e00;
    padding: 8px 14px;
    border-radius: 0 6px 6px 0;
    margin: 5px 0;
    font-size: 0.92em;
}
.card-warn {
    background: #fffbe6;
    border-left: 4px solid #e6a800;
    padding: 8px 14px;
    border-radius: 0 6px 6px 0;
    margin: 5px 0;
    font-size: 0.92em;
}
/* バッジ */
.badge-ng   { background:#e00;    color:#fff; padding:1px 8px; border-radius:10px; font-size:0.78em; font-weight:700; }
.badge-warn { background:#e6a800; color:#fff; padding:1px 8px; border-radius:10px; font-size:0.78em; font-weight:700; }
/* 箇所・補足 */
.loc { color:#666; font-size:0.83em; margin-top:3px; }
/* カテゴリ */
.cat { font-weight:700; margin-right:5px; }
/* エキスパンダーの数字バッジ */
.num-ng   { color:#e00; font-weight:700; }
.num-warn { color:#e6a800; font-weight:700; }
</style>
""", unsafe_allow_html=True)

# ── タイトル ──────────────────────────────────────────────────────────────────
st.title("🏥 医療クリニック Webサイト検品ツール")
st.caption("既存サイト・構築サイトを比較し、NG・要手動確認事項を全ページ自動チェックします")

# ── サイドバー：入力 ──────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ 検品設定")

    site_mode = st.radio(
        "案件タイプ",
        ["既存サイトあり", "新規案件（既存サイトなし）"],
        horizontal=True,
    )

    existing_url = ""
    manual_info  = {}

    if site_mode == "既存サイトあり":
        existing_url = st.text_input(
            "既存サイトURL（正解データ）",
            placeholder="https://www.example.com",
        )
    else:
        st.caption("📋 基本情報を入力してください")
        manual_info = {
            "tel":          st.text_input("電話番号",         placeholder="03-1234-5678"),
            "fax":          st.text_input("FAX番号",          placeholder="03-1234-5679"),
            "clinic_name":  st.text_input("院名",             placeholder="〇〇クリニック"),
            "director":     st.text_input("院長名",           placeholder="山田 太郎"),
            "hours":        st.text_area ("診療時間",         placeholder="月〜金 9:00〜18:00\n土 9:00〜13:00", height=80),
            "closed":       st.text_input("休診日",           placeholder="日・祝"),
            "address":      st.text_input("郵便番号・住所",   placeholder="〒123-4567 東京都〇〇区..."),
            "station":      st.text_input("最寄り駅・所要時間", placeholder="〇〇駅 徒歩5分"),
        }

    target_url = st.text_input(
        "構築サイトURL（チェック対象）",
        placeholder="https://staging.example.com",
    )
    wire_file = st.file_uploader(
        "ワイヤーフレーム PPTX（省略可）",
        type=["pptx"],
    )

    st.divider()

    use_agent = st.toggle(
        "🤖 LLMエージェント強化",
        value=False,
        help="意味的チェック・テキスト品質チェックを追加します（Anthropic API必須）",
    )

    llm_client = None
    model = ck.DEFAULT_MODEL

    if use_agent:
        api_key = (
            st.secrets.get("ANTHROPIC_API_KEY", "")
            if hasattr(st, "secrets")
            else ""
        ) or os.environ.get("ANTHROPIC_API_KEY", "")

        if not api_key:
            api_key = st.text_input(
                "ANTHROPIC_API_KEY",
                type="password",
                placeholder="sk-ant-...",
            )
        else:
            st.success("✓ APIキー設定済み")

        model = st.selectbox(
            "使用モデル",
            ["claude-sonnet-4-6", "claude-haiku-4-5-20251001"],
            index=0,
        )

    st.divider()

    run_btn = st.button(
        "▶ 検品実行",
        type="primary",
        use_container_width=True,
        disabled=not target_url.strip() or (
            site_mode == "既存サイトあり" and not existing_url.strip()
        ),
    )

    st.divider()
    with st.expander("チェック項目一覧", expanded=False):
        st.markdown("""
| # | 項目 | 対象 |
|---|---|---|
| 1 | ダミーテキスト（位置特定・×N件数） | 全ページ |
| 2 | 外部リンク target=_blank 未設定 | 全ページ |
| 3 | 空リンク・無効リンク | 全ページ |
| 4 | alt未設定・無意味なalt | 全ページ |
| 5 | h1（なし・重複・ダミー語） | 全ページ |
| 6 | 見出し階層スキップ | 全ページ |
| 7 | 電話番号ハイフン混在 | 全ページ |
| 8 | パンくずリスト | トップ以外 |
| 9 | reCAPTCHA 未設定 | フォームあり |
| 10 | 料金ダミー・税込未表記 | 全ページ |
| 11 | 自由診療 必要記載 | 料金記載ページ |
| 12 | 症例ページ 必要記載 | 症例ページ |
| 13 | 未承認医薬品 必要記載 | 該当ページ |
| 14 | 医療広告GL違反 | 全ページ |
| 15 | 基本情報照合 | トップページ |
| 16 | meta・OGP・canonical | トップページ |
| 17 | 投稿記事の移行確認 | トップページ |
| ✦ | ワイヤーフレーム照合 | オプション |
| 🤖 | 意味的チェック+テキスト品質 | LLMモード |
""")

# ── 検品実行 ──────────────────────────────────────────────────────────────────

def _fetch_existing(url: str):
    try:
        s = requests.Session()
        s.headers["User-Agent"] = "Mozilla/5.0 (compatible; WebInspector/1.0)"
        r = s.get(url, timeout=15)
        return BeautifulSoup(r.text, "lxml"), None
    except Exception as e:
        return None, str(e)


if run_btn:
    # バリデーション
    if site_mode == "既存サイトあり" and not existing_url.startswith("http"):
        st.error("既存サイトURLは http(s):// から始めてください")
        st.stop()
    if not target_url.startswith("http"):
        st.error("構築サイトURLは http(s):// から始めてください")
        st.stop()
    if use_agent and not (api_key if "api_key" in dir() else ""):
        st.error("LLMモードには ANTHROPIC_API_KEY が必要です")
        st.stop()

    # LLMクライアント初期化
    if use_agent:
        try:
            import anthropic
            llm_client = anthropic.Anthropic(api_key=api_key)
        except Exception as ex:
            st.error(f"Anthropic API初期化エラー: {ex}")
            st.stop()

    t_url = target_url.rstrip("/")
    e_url = existing_url.rstrip("/")
    bpath = ck.get_staging_base(t_url)

    # 結果をリセット
    st.session_state.pop("results", None)
    all_results: dict = {}
    start_ts = datetime.now()

    with st.status("検品実行中...", expanded=True) as status:
        # ── クロール ──────────────────────────────────────────────────────
        st.write(f"🔍 構築サイトを巡回中: **{t_url}**")
        if bpath:
            st.write(f"　 ステージングベースパス: `{bpath}`")
        crawler = ck.Crawler(t_url, bpath)
        crawler.crawl()
        st.write(f"✅ 巡回完了 — **{len(crawler.pages)} ページ**")

        # ── 既存サイト取得 ─────────────────────────────────────────────────
        existing_soup = None
        if site_mode == "既存サイトあり":
            st.write(f"🔍 既存サイト取得中: **{e_url}**")
            existing_soup, fetch_err = _fetch_existing(e_url)
            if fetch_err:
                st.warning(f"⚠️ 既存サイト取得エラー: {fetch_err}")
            else:
                st.write("✅ 既存サイト取得完了")
        else:
            st.write("📋 新規案件モード — 手動入力の基本情報で照合します")

        # ── ワイヤーフレーム保存 ───────────────────────────────────────────
        wire_path = None
        if wire_file:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
                tmp.write(wire_file.getvalue())
                wire_path = tmp.name
            st.write(f"📋 ワイヤーフレームファイル: `{wire_file.name}`")

        # ── 各ページチェック ───────────────────────────────────────────────
        st.write("🔎 チェック開始...")
        pages_list = list(crawler.pages.items())
        progress   = st.progress(0.0)
        log_area   = st.empty()

        for i, (url, page_data) in enumerate(pages_list):
            ppath = ck.rel_path(url, bpath)
            log_area.caption(f"チェック中: `{ppath}`")

            results = ck.run_checks(
                url, page_data, bpath, existing_soup, llm_client, model, manual_info
            )
            all_results[url] = results

            ng_n   = sum(1 for r in results if r["kind"] == "NG")
            warn_n = sum(1 for r in results if r["kind"] == "要手動確認")
            progress.progress((i + 1) / len(pages_list))

            icon = "🔴" if ng_n else ("🟡" if warn_n else "✅")
            st.write(f"　{icon} `{ppath}` — NG: {ng_n}件  要確認: {warn_n}件")

        log_area.empty()

        # ── ワイヤーフレーム照合 ───────────────────────────────────────────
        if wire_path:
            st.write("📋 ワイヤーフレーム照合中...")
            wire_res = ck.check_wireframe(wire_path, crawler.pages)
            top_url  = t_url + "/"
            key      = top_url if top_url in all_results else t_url
            all_results.setdefault(key, [])
            all_results[key] = wire_res + all_results[key]
            try:
                os.unlink(wire_path)
            except Exception:
                pass

        elapsed = (datetime.now() - start_ts).seconds
        status.update(
            label=f"✅ 検品完了（{elapsed}秒）",
            state="complete",
            expanded=False,
        )

    st.session_state["results"]     = all_results
    st.session_state["target_url"]  = t_url
    st.session_state["base_path"]   = bpath
    st.session_state["run_time"]    = start_ts.strftime("%Y%m%d_%H%M%S")
    st.session_state["llm_mode"]    = use_agent

# ── 結果表示 ──────────────────────────────────────────────────────────────────

def _render_issue(r: dict):
    """1件の課題をカードHTMLとして返す。"""
    is_ng  = r["kind"] == "NG"
    cls    = "card-ng" if is_ng else "card-warn"
    badge  = ('<span class="badge-ng">NG</span>'
              if is_ng else '<span class="badge-warn">要手動確認</span>')
    cat    = f'<span class="cat">{r["category"]}</span>'
    body   = r["content"]

    extras = []
    if r.get("location"):
        extras.append(f'📍 {r["location"]}')
    if r.get("existing"):
        v = r["existing"][:80]
        extras.append(f'既存値: {v}')
    if r.get("detected"):
        v = r["detected"][:80]
        extras.append(f'検出値: {v}')

    loc_html = ('<div class="loc">' + "　／　".join(extras) + "</div>") if extras else ""

    return f'<div class="{cls}">{badge} {cat}{body}{loc_html}</div>'


if "results" in st.session_state:
    all_results = st.session_state["results"]
    bpath       = st.session_state["base_path"]
    llm_mode    = st.session_state.get("llm_mode", False)

    total_ng    = sum(sum(1 for r in v if r["kind"] == "NG")        for v in all_results.values())
    total_warn  = sum(sum(1 for r in v if r["kind"] == "要手動確認") for v in all_results.values())
    total_pages = len(all_results)
    ok_pages    = sum(
        1 for v in all_results.values()
        if not any(r["kind"] in ("NG", "要手動確認") for r in v)
    )

    # ── サマリー指標 ──────────────────────────────────────────────────────────
    st.header("📊 検品サマリー")
    if llm_mode:
        st.caption("🤖 LLMエージェント強化モードで実行")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("🔴 NG件数",          total_ng)
    c2.metric("🟡 要手動確認件数",   total_warn)
    c3.metric("📄 チェックページ数", total_pages)
    c4.metric("✅ 問題なしページ",   ok_pages)

    # カテゴリ内訳
    cat_count = Counter(
        r["category"]
        for v in all_results.values()
        for r in v
        if r["kind"] == "NG"
    )
    if cat_count:
        st.subheader("NG カテゴリ内訳")
        cat_df = pd.DataFrame(
            cat_count.most_common(12), columns=["カテゴリ", "件数"]
        ).set_index("カテゴリ")
        st.bar_chart(cat_df)

    st.divider()

    # ── タブ：ページ別 / 全件テーブル ─────────────────────────────────────────
    tab_pages, tab_table = st.tabs(["📋 ページ別詳細", "📑 全件テーブル"])

    # ── ページ別詳細 ──────────────────────────────────────────────────────────
    with tab_pages:
        # NG件数の多い順に並べ替え
        sorted_pages = sorted(
            all_results.items(),
            key=lambda x: (
                -sum(1 for r in x[1] if r["kind"] == "NG"),
                -sum(1 for r in x[1] if r["kind"] == "要手動確認"),
            ),
        )

        for url, results in sorted_pages:
            ppath  = ck.rel_path(url, bpath)
            ng_n   = sum(1 for r in results if r["kind"] == "NG")
            warn_n = sum(1 for r in results if r["kind"] == "要手動確認")

            if ng_n == 0 and warn_n == 0:
                continue

            icon  = "🔴" if ng_n else "🟡"
            label = (
                f"{icon} {ppath}"
                f"　　"
                f"<span class='num-ng'>NG {ng_n}件</span>"
                f"　"
                f"<span class='num-warn'>要確認 {warn_n}件</span>"
            )

            with st.expander(f"{icon} {ppath}　NG:{ng_n}件  要確認:{warn_n}件",
                             expanded=(ng_n > 0)):
                # NG → 要手動確認 の順でソート
                sorted_r = sorted(
                    results,
                    key=lambda x: (0 if x["kind"] == "NG" else 1),
                )

                # NG
                ng_items = [r for r in sorted_r if r["kind"] == "NG"]
                if ng_items:
                    st.markdown(f"**🔴 NG — {len(ng_items)}件**")
                    cards = "".join(_render_issue(r) for r in ng_items)
                    st.markdown(cards, unsafe_allow_html=True)

                # 要手動確認
                warn_items = [r for r in sorted_r if r["kind"] == "要手動確認"]
                if warn_items:
                    st.markdown(f"**🟡 要手動確認 — {len(warn_items)}件**")
                    cards = "".join(_render_issue(r) for r in warn_items)
                    st.markdown(cards, unsafe_allow_html=True)

        if ok_pages:
            st.caption(f"✅ 問題なし: {ok_pages} ページ（詳細は全件テーブルを確認）")

    # ── 全件テーブル ──────────────────────────────────────────────────────────
    with tab_table:
        rows = []
        for url, results in all_results.items():
            ppath = ck.rel_path(url, bpath)
            for r in results:
                rows.append({
                    "ページパス": ppath,
                    "種別":       r["kind"],
                    "カテゴリ":   r["category"],
                    "内容":       r["content"],
                    "箇所":       r.get("location", ""),
                    "既存値":     r.get("existing", ""),
                    "検出値":     r.get("detected", ""),
                })

        df = pd.DataFrame(rows)

        # フィルター
        col_f1, col_f2 = st.columns([1, 3])
        with col_f1:
            kind_filter = st.multiselect(
                "種別フィルター",
                ["NG", "要手動確認", "OK"],
                default=["NG", "要手動確認"],
            )
        with col_f2:
            kw = st.text_input("キーワード検索（内容・カテゴリ）", "")

        df_filtered = df[df["種別"].isin(kind_filter)] if kind_filter else df
        if kw:
            mask = df_filtered["内容"].str.contains(kw, na=False) | \
                   df_filtered["カテゴリ"].str.contains(kw, na=False)
            df_filtered = df_filtered[mask]

        def _style_row(row):
            c = {"NG": "#ffcccc", "要手動確認": "#fffacc"}.get(row["種別"], "#ccffcc")
            return [f"background-color:{c}"] * len(row)

        st.dataframe(
            df_filtered.style.apply(_style_row, axis=1),
            use_container_width=True,
            height=550,
        )
        st.caption(f"表示: {len(df_filtered)} 件 / 全 {len(df)} 件")

    # ── Excel ダウンロード ─────────────────────────────────────────────────────
    st.divider()
    col_dl, col_info = st.columns([2, 5])
    with col_dl:
        buf = io.BytesIO()
        ck.write_excel(all_results, buf)
        buf.seek(0)
        st.download_button(
            label="📥 Excelダウンロード",
            data=buf,
            file_name=f"検品結果_{st.session_state['run_time']}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
    with col_info:
        st.caption(
            "サマリー・ページ別詳細の2シート構成。"
            "NG行は赤、要手動確認は黄、問題なしは緑で色分けされています。"
        )

else:
    # ── 初期画面 ──────────────────────────────────────────────────────────────
    st.info("👈 左のサイドバーでURLを入力し、「検品実行」を押してください")

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("""
#### 🔴 自動検出（NG）
- ダミーテキスト（位置特定つき）
- 外部リンク・空リンク
- alt属性未設定
- h1・見出し階層
- 電話番号ハイフン混在
- 料金ダミー・税込未表記
- 基本情報照合（TEL・院名）
""")
    with col2:
        st.markdown("""
#### 🟡 要手動確認
- パンくずリスト（独自実装）
- meta・OGP・canonical
- 自由診療 必要記載
- 症例ページ 必要記載
- 医療広告GL グレーゾーン
- 投稿記事の移行
- レイアウト・フォーム確認
""")
    with col3:
        st.markdown("""
#### 🤖 LLMモード追加
- 意味的な文脈チェック
  （「3ヶ月後」= 治療期間など）
- 体験談・口コミ形式の検出
- テキスト品質・誤字脱字
- 表記ゆれ
- 院長名・診療時間の比較
""")
