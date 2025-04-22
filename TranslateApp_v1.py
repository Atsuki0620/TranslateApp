import streamlit as st
import pandas as pd
from io import BytesIO
from googletrans import Translator
import time

# ─── ユーティリティ関数 ─────────────────────────────────

@st.cache_data
def load_file(file) -> pd.DataFrame | None:
    """ExcelまたはCSVを読み込み。エラー時はNoneを返す。"""
    try:
        if file.name.endswith('.csv'):
            return pd.read_csv(file)
        else:
            return pd.read_excel(file)
    except Exception as e:
        st.error(f"📄 ファイル読み込みエラー: {e}")
        return None

def safe_translate(translator: Translator, text: str, dest: str = 'ja',
                   retries: int = 3, delay: float = 1.0) -> str:
    """翻訳を安全に行うための関数（エラー処理＋リトライ）"""
    if not text:
        return ""
    max_len = 4500
    translated = ""
    for start in range(0, len(text), max_len):
        segment = text[start:start + max_len]
        for attempt in range(retries):
            try:
                part = translator.translate(segment, dest=dest).text
                translated += part
                break
            except Exception as e:
                if attempt < retries - 1:
                    time.sleep(delay)
                else:
                    translated += f"[翻訳失敗: {e}]"
        time.sleep(delay)
    return translated

# ─── メイン処理 ─────────────────────────────────────────

def main():
    st.title("🤖 多言語対応 Excel・CSV自動翻訳アプリ")

    uploaded_file = st.file_uploader("📤 ExcelまたはCSVをアップロード", type=['xlsx', 'xls', 'csv'])
    if not uploaded_file:
        st.info("まずファイルをアップロードしてください。")
        return

    df = load_file(uploaded_file)
    if df is None:
        return

    st.subheader("📝 アップロードデータの確認（先頭10行）")
    st.dataframe(df.head(10))

    columns_to_translate = st.multiselect("▶ 翻訳対象の列を複数選択できます", df.columns)
    if not columns_to_translate:
        st.warning("翻訳する列を1つ以上選択してください。")
        return

    sleep_time = st.number_input(
        "⏱ 翻訳間隔(秒)（推奨:1.0〜3.0秒）",
        min_value=0.5, max_value=5.0, value=1.0, step=0.5, format="%.1f"
    )

    if st.button("🚀 翻訳開始"):
        translator = Translator()
        total_tasks = len(df) * len(columns_to_translate)
        current_task = 0
        progress = st.progress(0.0)
        error_count = 0

        with st.spinner("🔄 翻訳中…しばらくお待ちください"):
            for col in columns_to_translate:
                translations = []
                for cell in df[col].astype(str):
                    result = safe_translate(translator, cell, dest='ja')
                    if result.startswith("[翻訳失敗"):
                        error_count += 1
                    translations.append(result)
                    current_task += 1
                    progress.progress(current_task / total_tasks)
                    time.sleep(sleep_time)
                df[f"{col}_JP"] = translations

        st.success(f"✅ 翻訳完了（エラー件数: {error_count}）")
        st.dataframe(df)

        # Excelで出力
        output = BytesIO()
        try:
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Translated')
            output.seek(0)
            st.download_button(
                "📥 翻訳済みExcelをダウンロード",
                data=output,
                file_name="translated_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"🖨 Excel生成エラー: {e}")

if __name__ == "__main__":
    main()
