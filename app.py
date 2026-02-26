"""
ガントチャート生成ツール - Streamlit Web GUI
起動: streamlit run app.py
"""
import streamlit as st
import tempfile
import os
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
    
    st.divider()
    st.markdown("### 📝 Excelフォーマット")
    st.markdown("""
    | 列 | 内容 |
    |--|--|
    | A | 大項目 |
    | B | 種別 (`milestone`/`period`/`task`) |
    | C | 名前 |
    | D | 開始日 |
    | E | 終了日 |
    """)

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
    
    with st.expander("📋 読み込みデータの詳細", expanded=False):
        for d in data:
            items_n = len(d.get("items", []))
            tasks_n = len(d.get("tasks", []))
            st.markdown(f"**{d['name']}** — {items_n} items, {tasks_n} tasks")
    
    # 生成ボタン
    if st.button("🚀 ガントチャートを生成", type="primary", use_container_width=True):
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
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("### 1️⃣ Excelを用意")
        st.markdown("A〜E列にデータを入力")
    with col2:
        st.markdown("### 2️⃣ アップロード")
        st.markdown("ファイルをドラッグ＆ドロップ")
    with col3:
        st.markdown("### 3️⃣ ダウンロード")
        st.markdown("PPTXファイルを取得")
