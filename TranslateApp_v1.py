import streamlit as st
import pandas as pd
from io import BytesIO
from googletrans import Translator
import time

# ─── ユーティリティ関数 ─────────────────────────────────

@st.cache_data
def load_excel(file) -> pd.DataFrame | None:
    """Excel読み込み。失敗時はNoneを返し、画面にエラー表示。"""
    try:
        return pd.read_excel(file)
    except Exception as e:
        st.error(f"📄 Excel読み込みエラー: {e}")
        return None

def safe_translate(translator: Translator, text: str, dest: str = 'ja',
                   retries: int = 3, delay: float = 1.0) -> str:
    """
    ・最大4500文字ずつにチャンク分割して翻訳  
    ・各チャンクでリトライを行い、最終的に失敗した場合はエラーを注記  
    """
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
        # 各チャンク後のインターバル（API負荷軽減）
        time.sleep(delay)
    return translated

# ─── メイン処理 ─────────────────────────────────────────

def main():
    st.title("多言語対応 Excel自動翻訳アプリ")

    uploaded_file = st.file_uploader("📤 Excelファイルをアップロード", type=['xlsx', 'xls'])
    if not uploaded_file:
        st.info("まずはファイルをアップロードしてください。")
        return

    df = load_excel(uploaded_file)
    if df is None:
        return

    st.write("### アップロードデータ例")
    st.dataframe(df.head(5))

    column_to_translate = st.selectbox("▶ 翻訳対象の列を選択", df.columns)
    sleep_time = st.number_input(
        "⏱ 翻訳間隔(秒)（推奨:1.0〜3.0秒）",
        min_value=0.5, max_value=5.0, value=1.0, step=0.5, format="%.1f"
    )

    if st.button("🚀 翻訳開始"):
        translator = Translator()
        translations: list[str] = []
        total = len(df)
        progress = st.progress(0.0)
        error_count = 0

        with st.spinner("🔄 翻訳中…しばらくお待ちください"):
            for idx, cell in enumerate(df[column_to_translate].astype(str), start=1):
                result = safe_translate(translator, cell, dest='ja')
                if result.startswith("[翻訳失敗"):
                    error_count += 1
                translations.append(result)
                progress.progress(idx / total)
                time.sleep(sleep_time)

        df[f"{column_to_translate}_JP"] = translations
        st.success(f"✅ 翻訳完了 （エラー件数: {error_count}）")
        st.dataframe(df)

        # ── Excel書き出し & ダウンロード ──
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
