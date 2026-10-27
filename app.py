import streamlit as st
import pandas as pd
import openpyxl
import google.generativeai as genai
import io
import json
from concurrent.futures import ThreadPoolExecutor, as_completed

# --- 基本設定 ---
st.set_page_config(page_title="工場監査表 AIチェックツール", page_icon="📋", layout="wide")

st.sidebar.header("初期設定")
api_key = st.sidebar.text_input("Gemini APIキーを入力してください", type="password")

model = None
if api_key:
    genai.configure(api_key=api_key)
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

with st.expander("⚙️ AIチェックルール（クリックで確認・編集）", expanded=False):
    rules = """【校正基準】
1. 誤字・脱字の修正（例：「貯水層」→「貯水槽」など）。
2. 文末の完全文化：箇条書きの体言止め（名詞止め）は不可とし、「〜だ・である調」または「〜を確認した」等の断定的な完全な文章に修正。
3. 専門用語の判定：HACCP、SSOP、ルクス、ATP、メッシュ等は適切に判定。
4. 表記ゆれの統一：半角カタカナは全角カタカナに統一。
5. 単位と日付の補足：数値には適切な単位（ルクスなど）を補足し、年号は西暦（26年→2026年など）に修正。
6. 対象外：極端に短いテキストや「該当なし」「該当無し」「ー（ハイフン）」だけのテキストは修正不要。"""
    system_rules = st.text_area("校正ルール", value=rules, height=180)

uploaded_file = st.file_uploader("監査表Excelファイル（.xlsx）をアップロード", type=["xlsx"])

def check_batch(batch, rules_text, model_instance):
    input_data = [{"row": item["row"], "text": item["original"]} for item in batch]
    input_json = json.dumps(input_data, ensure_ascii=False)
    prompt = f"""あなたは食品工場の品質保証・監査報告書の専門校正担当です。
以下の【校正基準】に基づき、入力された各行のテキストを厳格に校正してください。

{rules_text}

【入力データ】
{input_json}

【出力形式】
修正が必要な項目のみを抽出し、以下のJSON配列形式で出力してください。修正不要な行は含めないでください。
[
  {{"row": 行番号, "suggested": "修正後の文章"}}
]
"""
    generation_config = {"response_mime_type": "application/json", "temperature": 0.1}
    res = model_instance.generate_content(prompt, generation_config=generation_config)
    res_text = res.text.strip()
    if res_text.startswith("```json"):
        res_text = res_text.replace("```json", "").replace("```", "").strip()
    parsed = json.loads(res_text)
    return {p["row"]: p["suggested"] for p in parsed if isinstance(p, dict) and "row" in p and "suggested" in p}

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
            items_to_check = []
            for row_idx in range(9, sheet.max_row + 1):
                target_col = 9
                cell_value = sheet.cell(row=row_idx, column=target_col).value
                if cell_value and isinstance(cell_value, str):
                    orig = cell_value.strip()
                    if len(orig) > 2 and orig not in ["該当無し", "該当なし", "ー", "該当なし。"]:
                        items_to_check.append({
                            "row": row_idx,
                            "col": target_col,
                            "original": orig
                        })

            total_items = len(items_to_check)
            if total_items == 0:
                st.warning("チェック対象となる記述が見つかりませんでした。")
            else:
                progress_bar = st.progress(0, text="高速並列チェックを開始しています...")
                
                batch_size = 15
                batches = [items_to_check[i:i + batch_size] for i in range(0, total_items, batch_size)]
                total_batches = len(batches)
                
                all_suggestions = {}
                completed_batches = 0
                
                # 3並列で同時にAIへ通信
                with ThreadPoolExecutor(max_workers=3) as executor:
                    future_to_batch = {
                        executor.submit(check_batch, batch, system_rules, model): batch
                        for batch in batches
                    }
                    
                    for future in as_completed(future_to_batch):
                        try:
                            batch_result = future.result()
                            all_suggestions.update(batch_result)
                        except Exception as e:
                            st.warning(f"一部グループの処理で警告が出ました: {e}")
                        
                        completed_batches += 1
                        pct = int((completed_batches / total_batches) * 100)
                        progress_bar.progress(pct, text=f"進捗: {completed_batches}/{total_batches} グループ完了 ({pct}%)")

                all_results = []
                for item in items_to_check:
                    if item["row"] in all_suggestions:
                        all_results.append({
                            "row": item["row"],
                            "col": item["col"],
                            "original": item["original"],
                            "suggested": all_suggestions[item["row"]],
                            "apply": True
                        })

                progress_bar.progress(100, text="チェック完了！")
                st.session_state.processed_data = all_results
                st.success(f"チェック完了！ 合計 {len(all_results)}件の修正案が見つかりました。")
                st.rerun()

    # 結果一覧とダウンロード
    if st.session_state.processed_data is not None:
        st.subheader("🔍 チェック結果一覧と適用")
        st.caption("チェック（✅）が入っている項目のみ、出力されるExcelファイルに修正が反映されます。")
        
        results = st.session_state.processed_data
        
        if len(results) == 0:
            st.info("AIによる修正提案はありませんでした（すべての項目が基準を満たしています）。")
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
