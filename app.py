import streamlit as st
import pandas as pd
import openpyxl
import google.generativeai as genai
import io
import time

# --- 設定 ---
st.set_page_config(page_title="工場監査表 AIチェックツール", page_icon="📋", layout="wide")

# APIキーの入力
st.sidebar.header("初期設定")
api_key = st.sidebar.text_input("Gemini APIキーを入力してください", type="password")

model = None
if api_key:
    genai.configure(api_key=api_key)
    # 最新モデル gemini-3.8-flash を指定
    selected_model_name = 'gemini-3.8-flash'
    
    try:
        model = genai.GenerativeModel(selected_model_name)
        st.sidebar.success(f"接続完了: {selected_model_name}")
    except Exception as e:
        st.sidebar.error(f"モデル接続エラー: {e}")
else:
    st.sidebar.warning("APIキーを入力すると機能が有効になります。")

st.title("📋 工場監査表 AIチェックツール")
st.markdown("""
アップロードされた監査表の記述内容から誤字・脱字、体言止め、不自然な言い回しをAIが検出し、
公的・社外提出用の文章（だ・である調）へ自動校正します。
""")

# --- プロンプト（ルール）設定 ---
with st.expander("⚙️ AIチェックルール・プロンプト設定（クリックで展開・編集）", expanded=False):
    st.caption("AIに指示する校正ルールです。必要に応じて書き換えて調整できます。")
    default_prompt = """あなたは食品工場の品質保証・監査報告書の専門校正担当です。
以下の【校正基準】に基づき、入力されたテキストを厳格に校正してください。

【校正基準】
1. 誤字・脱字の修正（例：「貯水層」→「貯水槽」など）。
2. 文末の完全文化：社外・公的機関へ提出できる品質にするため、箇条書きの体言止め（名詞止め）は不可とし、「〜だ・である調」または「〜を確認した」等といった断定的な完全な文章に修正してください。
3. 専門用語の判定：HACCP、SSOP、ルクス、ATP、メッシュ等の専門用語は一般的なAIの食品衛生知識に基づいて判定してください。
4. 表記ゆれの統一：半角カタカナは全角カタカナに統一してください。
5. 単位と日付の補足：数値には適切な単位（ルクスなど）を補足し、年号は西暦（26年→2026年など）に修正してください。
6. 対象外：極端に短いテキストや「該当なし」「該当無し」「ー（ハイフン）」だけのテキストは修正の必要はありません。「問題なし」と判定してください。

入力テキスト: {text}

【出力形式】
修正が必要な場合は「修正後: <修正したテキスト>」のみを出力してください。
修正が不要な場合、または対象外の場合は「問題なし」と出力してください。
理由や解説は一切出力しないでください。"""
    system_prompt = st.text_area("プロンプト", value=default_prompt, height=250)

# --- ファイルアップロードと処理 ---
uploaded_file = st.file_uploader("監査表Excelファイル（.xlsx）をアップロード", type=["xlsx"])

if uploaded_file is not None and model is not None:
    if 'processed_data' not in st.session_state or st.session_state.uploaded_filename != uploaded_file.name:
        st.session_state.processed_data = None
        st.session_state.uploaded_filename = uploaded_file.name

    try:
        wb = openpyxl.load_workbook(uploaded_file)
        sheet = wb['監査表']
    except Exception as e:
        st.error(f"ファイルの読み込みに失敗しました。シート名が「監査表」であるか確認してください。エラー詳細: {e}")
        st.stop()

    if st.session_state.processed_data is None:
        if st.button("AIチェックを実行する", type="primary"):
            with st.spinner("AIが監査表をチェックしています..."):
                results = []
                progress_bar = st.progress(0)
                max_rows = sheet.max_row
                
                for row_idx in range(9, max_rows + 1):
                    progress_bar.progress(row_idx / max_rows)
                    
                    target_col = 9
                    cell_value = sheet.cell(row=row_idx, column=target_col).value
                    
                    if cell_value and isinstance(cell_value, str):
                        original_text = cell_value.strip()
                        if len(original_text) > 2 and original_text not in ["該当無し", "該当なし", "ー", "該当なし。"]:
                            prompt = system_prompt.replace("{text}", original_text)
                            try:
                                response = model.generate_content(prompt)
                                ai_result = response.text.strip()
                                
                                if "問題なし" not in ai_result:
                                    suggested_text = ai_result.replace("修正後:", "").replace("修正後：", "").strip()
                                    if suggested_text:
                                        results.append({
                                            "row": row_idx,
                                            "col": target_col,
                                            "original": original_text,
                                            "suggested": suggested_text,
                                            "apply": True
                                        })
                            except Exception as e:
                                st.warning(f"行 {row_idx} の処理中にエラーが発生しました: {e}")
                                time.sleep(1)

                st.session_state.processed_data = results
                st.success(f"チェック完了！ {len(results)}件の修正案が見つかりました。")
                st.rerun()

    if st.session_state.processed_data is not None:
        st.subheader("🔍 チェック結果一覧と適用")
        st.caption("チェック（✅）が入っている項目のみ、出力されるExcelファイルに修正が反映されます。")
        
        results = st.session_state.processed_data
        
        if len(results) == 0:
            st.info("AIによる修正提案はありませんでした。")
        else:
            for idx, res in enumerate(results):
                col1, col2 = st.columns([1, 11])
                with col1:
                    res['apply'] = st.checkbox("適用", value=res['apply'], key=f"check_{idx}")
                with col2:
                    st.markdown(f"**行 {res['row']}**")
                    st.markdown(f"❌ 元の記載: `{res['original']}`")
                    st.markdown(f"✨ 修正案 : **{res['suggested']}**")
                st.divider()

            st.subheader("💾 修正版Excelファイルの出力")
            
            for res in results:
                if res['apply']:
                    sheet.cell(row=res['row'], column=res['col']).value = res['suggested']
            
            virtual_workbook = io.BytesIO()
            wb.save(virtual_workbook)
            virtual_workbook.seek(0)
            
            new_filename = f"【修正版】_{uploaded_file.name}"
            
            st.download_button(
                label="📥 適用済みのExcelファイルをダウンロード",
                data=virtual_workbook,
                file_name=new_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
