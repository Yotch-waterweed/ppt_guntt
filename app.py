"""
ガントチャート生成ツール - Streamlit Web GUI
起動: streamlit run app.py
"""
import streamlit as st
import tempfile
import os
import pandas as pd
from excel_reader import read_excel
from gantt_generator import GanttChartGenerator

st.set_page_config(page_title="ガントチャート生成ツール", page_icon="📊", layout="wide")

# --- ヘッダー ---
st.title("📊 ガントチャート PPTX 生成ツール")
st.caption("Excelファイルからガントチャートの PowerPoint スライドを自動生成します")

# --- サイドバー: 設定 ---
with st.sidebar:
    st.header("⚙️ 設定")
    
    calendar_mode = st.selectbox(
        "カレンダー表示モード",
        options=["auto", "bme", "monthly", "quarterly"],
        index=0,
        help="auto: 月数に応じて自動選択 / bme: 上旬・中旬・下旬 / monthly: 月単位 / quarterly: 四半期"
    )
    
    col1, col2 = st.columns(2)
    with col1:
        start_month = st.text_input("開始月 (任意)", placeholder="2025-04", help="YYYY-MM 形式。空欄でデータから自動判定")
    with col2:
        end_month = st.text_input("終了月 (任意)", placeholder="2027-03", help="YYYY-MM 形式。空欄でデータから自動判定")
    
    external_label = st.slider(
        "テキスト外出し閾値（月数）",
        min_value=1, max_value=24, value=3,
        help="カレンダーがこの月数以上の場合、矢羽のテキストを外側に配置"
    )

# --- Excelの書き方ガイド ---
with st.expander("📖 はじめての方へ：Excelの書き方ガイド（クリックして展開）", expanded=False):
    st.markdown("### 📝 Excelの基本フォーマット")
    st.markdown("1行目はヘッダーにし、**A列〜E列**に以下のルールでデータを入力してください。")
    
    # サンプルデータフレームを表示（Excel風）
    df_sample = pd.DataFrame({
        "大項目 (A列)": ["Phase1: 企画", "Phase1: 企画", "Phase1: 企画", "Phase2: 設計", "Phase2: 設計"],
        "種別 (B列)": ["milestone", "period", "task", "milestone", "task"],
        "名前 (C列)": ["PJ開始", "企画・構想", "現状分析", "設計完了", "基本設計"],
        "開始日 (D列)": ["2025-04-01", "2025-04-01", "2025-04-01", "2025-09-30", "2025-07-01"],
        "終了日 (E列)": ["", "2025-09-30", "2025-06-30", "", "2025-08-31"]
    })
    st.dataframe(df_sample, hide_index=True, use_container_width=True)
    
    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("#### 📌 B列の「種別」について")
        st.markdown("""
        - `period` 🟦 **(プロジェクト期間)**
          - 指定期間に青い四角いバーを描画します。
        - `milestone` 🔴 **(マイルストーン)**
          - 特定の日に赤いひし形を描画します。終了日は空欄でOKです。
        - `task` ⏭️ **(タスク / 矢羽)**
          - 指定期間にグレーの矢羽を描画します。プロジェクトの下にぶら下がって複数行並びます。
        """)
    with col_b:
        st.markdown("#### 📌 A列の「大項目」について")
        st.markdown("""
        大項目（A列）に**同じ名前**が連続している行は、**1つのグループ（レーン）**としてまとめられます。
        大項目が変わると、スライド上で新しい行（レーン）が作られます。
        """)
    
    # サンプルダウンロードボタン
    try:
        with open("sample_data.xlsx", "rb") as f:
            st.download_button("📥 入力用テンプレート（サンプルExcel）をダウンロード", data=f.read(), file_name="sample_data.xlsx", type="primary")
    except FileNotFoundError:
        pass

# --- 出力サンプル ---
with st.expander("✨ こんなパワポが作れます（出力サンプル）", expanded=False):
    st.markdown("カレンダーの長さ等に合わせて自動でレイアウトや文字サイズが調整されます。")
    
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**6ヶ月プロジェクト（BMEモード） 終了月2026-10 テキスト外だし閾値24ヶ月**")
        if os.path.exists("assets/sample_6month.png"):
            st.image("assets/sample_6month.png", use_container_width=True)
            
        st.markdown("**12ヶ月プロジェクト（monthlyモード） 終了月2026-04 テキスト外だし閾値5ヶ月**")
        if os.path.exists("assets/sample_12month.png"):
            st.image("assets/sample_12month.png", use_container_width=True)
            
    with col2:
        st.markdown("**32ヶ月プロジェクト（四半期モード）終了月2028-01 テキスト外だし閾値5ヶ月**")
        if os.path.exists("assets/sample_32month.png"):
            st.image("assets/sample_32month.png", use_container_width=True)
            
st.divider()

# --- メインエリア ---
uploaded_file = st.file_uploader(
    "📁 Excelファイルをアップロード",
    type=["xlsx", "xls"],
    help="A〜E列に「大項目, 種別, 名前, 開始日, 終了日」を記入したExcelファイル"
)

if uploaded_file is not None:
    # Excelを一時ファイルに保存して読み込み
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp.write(uploaded_file.getvalue())
        tmp_path = tmp.name
    
    try:
        data = read_excel(tmp_path)
    except Exception as e:
        st.error(f"❌ Excelの読み込みに失敗しました: {e}")
        os.unlink(tmp_path)
        st.stop()
    finally:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)
    
    if not data:
        st.warning("⚠️ データが見つかりませんでした。Excelのフォーマットを確認してください。")
        st.stop()
    
    # データプレビュー
    st.success(f"✅ {len(data)} プロジェクトを検出しました")
    
    with st.expander("📋 読み込みデータのプレビュー", expanded=False):
        for d in data:
            items_n = len(d.get("items", []))
            tasks_n = len(d.get("tasks", []))
            st.markdown(f"**{d['name']}** — {items_n} items, {tasks_n} tasks")
    
    # 生成ボタン
    if st.button("🚀 ガントチャートをPTTXで生成", type="primary", use_container_width=True):
        with st.spinner("生成中..."):
            try:
                gen = GanttChartGenerator(
                    data,
                    start_month=start_month if start_month else None,
                    end_month=end_month if end_month else None,
                    force_external_label_months=external_label,
                    calendar_mode=calendar_mode,
                )
                
                # 一時ファイルに出力
                with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp_out:
                    gen.generate(tmp_out.name)
                    tmp_out_path = tmp_out.name
                
                # ダウンロードボタン
                with open(tmp_out_path, "rb") as f:
                    pptx_bytes = f.read()
                os.unlink(tmp_out_path)
                
                st.balloons()
                st.download_button(
                    label="📥 PPTXをダウンロード",
                    data=pptx_bytes,
                    file_name="gantt_chart.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"❌ 生成エラー: {e}")
else:
    # アップロード前の説明
    st.info("👆 上のエリアにExcelファイルをドラッグ＆ドロップしてください")

