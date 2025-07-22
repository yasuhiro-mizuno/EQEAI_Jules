import streamlit as st
import docx
import fitz  # PyMuPDF
import re
import openai
from dotenv import load_dotenv
import os
import json

# 環境変数の読み込み
load_dotenv()

# Azure OpenAIクライアントの初期化 (v1.0.0以降の作法)
client = openai.AzureOpenAI(
    api_key=os.getenv("AZURE_OPENAI_API_KEY"),
    azure_endpoint=os.getenv("AZURE_OPENAI_ENDPOINT"),
    api_version=os.getenv("AZURE_OPENAI_VERSION")
)
DEPLOYMENT_NAME = 'o4-mini'

def extract_text_from_docx(file):
    doc = docx.Document(file)
    text = ""
    for para in doc.paragraphs:
        text += para.text + "\n"
    return text

def extract_text_from_pdf(file):
    pdf = fitz.open(stream=file.read(), filetype="pdf")
    text = ""
    for page in pdf:
        text += page.get_text()
    return text

def evaluate_text_with_llm(text_content, selected_rules):
    rule_mapping = {
        "conciseness": "簡潔な文: 受動態、二重否定、部分否定、使役表現が使われていないか。",
        "missing_elements": "必要な語の欠落: 主語、述語、目的語が明確で、欠落していないか。",
        "ambiguity": "曖昧語の回避: 「適切に」「可能な限り」などの曖昧な表現が使われていないか。",
        "typos": "誤字脱字: 誤字や脱字がないか。",
        "dependency": "係り受け: 係り受けが一意に解釈できるか。"
    }
    active_rules_text = "\n".join([f"- {rule_mapping[key]}" for key in selected_rules])

    system_prompt = "あなたは、提供された文章を分析し、特定の品質基準に基づいて改善点を提案するAIアシスタントです。あなたの応答は、必ず指定されたJSON形式でなければなりません。他のテキストは一切含めないでください。"
    
    user_prompt = f"""
以下の「品質基準」に従って、「評価対象の文章」の各文を評価してください。
結果をJSON形式で返してください。

# 品質基準
{active_rules_text}

# 出力形式の例
{{
  "evaluations": [
    {{
      "original_sentence": "ユーザーによりデータが送信される。",
      "has_issue": true,
      "reason": "簡潔な文：受動態が使用されています。",
      "suggestion": "ユーザーがデータを送信する。"
    }},
    {{
      "original_sentence": "この文には問題ありません。",
      "has_issue": false,
      "reason": "",
      "suggestion": ""
    }}
  ]
}}

# 評価対象の文章
{text_content}
"""

    try:
        response = client.chat.completions.create(
            model=DEPLOYMENT_NAME,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            max_completion_tokens=4096 
        )
        content = response.choices[0].message.content
        
        # LLMの応答からJSONオブジェクトを抽出する
        match = re.search(r'\{.*\}', content, re.DOTALL)
        if match:
            return match.group(0)
        else:
            return f'{{"error": "LLM response did not contain a valid JSON object.", "raw_response": "{json.dumps(content)}"}}'
            
    except Exception as e:
        return f'{{"error": "An error occurred: {str(e)}"}}'

import json

def parse_llm_response_to_markdown_table(responses):
    header = "| 元の記述 | 改善点の詳細 | 改善案 |\n|-----------|-----------|---------|"
    issue_rows = []
    all_evaluations = [] # すべての評価結果を格納するリスト
    parsing_errors = [] # パースエラー情報を格納するリスト

    for res_str in responses:
        try:
            data = json.loads(res_str)
            if "error" in data:
                parsing_errors.append({"response": res_str, "error": data["error"]})
                continue

            evaluations = data.get("evaluations", [])
            all_evaluations.extend(evaluations) # UI表示用にすべての評価を保存
            
            for item in evaluations:
                if item.get("has_issue"):
                    row = f"| {item.get('original_sentence', '')} | {item.get('reason', '')} | {item.get('suggestion', '')} |"
                    issue_rows.append(row)

        except json.JSONDecodeError as e:
            parsing_errors.append({"response": res_str, "error": f"JSONDecodeError: {e}"})
            continue
        except AttributeError as e:
            parsing_errors.append({"response": res_str, "error": f"AttributeError: {e}"})
            continue
            
    if not issue_rows:
        table = "指摘事項は見つかりませんでした。"
    else:
        table = header + "\n" + "\n".join(issue_rows)
    
    return table, all_evaluations, parsing_errors


def main():
    st.title("表現品質評価AI")

    try:
        # ファイルアップローダー
        uploaded_file = st.file_uploader("WordまたはPDFファイルをアップロードしてください", type=["docx", "pdf"])

        # 表現品質評価ルール
        st.sidebar.title("表現品質評価ルール")
        rules = {
            "conciseness": st.sidebar.checkbox("簡潔な文"),
            "missing_elements": st.sidebar.checkbox("必要な語の欠落"),
            "ambiguity": st.sidebar.checkbox("曖昧語の回避"),
            "typos": st.sidebar.checkbox("誤字脱字"),
            "dependency": st.sidebar.checkbox("係り受け")
        }

        # 評価ボタン
        if st.button("評価"):
            if uploaded_file is not None:
                selected_rules = [rule for rule, checked in rules.items() if checked]
                if not selected_rules:
                    st.error("評価ルールを1つ以上選択してください。")
                    return

                with st.spinner("評価中..."):
                    text = ""
                    if uploaded_file.type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
                        text = extract_text_from_docx(uploaded_file)
                    elif uploaded_file.type == "application/pdf":
                        text = extract_text_from_pdf(uploaded_file)
                    
                    llm_response = evaluate_text_with_llm(text, selected_rules)
                    
                    # LLMの応答をパースしてMarkdownの表形式に変換
                    markdown_table, all_evaluations, parsing_errors = parse_llm_response_to_markdown_table([llm_response])

                    st.success("評価が完了しました。")

                    st.markdown("### 確認結果サマリー（改善点のみ）")
                    st.markdown(markdown_table)

                    # ダウンロードボタン
                    st.download_button(
                        label="確認結果サマリーをダウンロード",
                        data=markdown_table,
                        file_name="review_summary.md",
                        mime="text/markdown",
                    )

                    # パースエラーがあった場合に詳細を表示
                    if parsing_errors:
                        with st.expander("LLM応答の解析エラー"):
                            st.error("いくつかのLLMの応答を解析できませんでした。以下に詳細を示します。")
                            for error_info in parsing_errors:
                                st.markdown("**エラー内容:**")
                                st.code(error_info.get("error"))
                                st.markdown("**LLMからの生の応答:**")
                                st.code(error_info.get("response"))
                                st.divider()

                    # すべての評価結果を詳細表示
                    with st.expander("すべての確認プロセスを表示（分割された文章と確認結果）"):
                        if not all_evaluations and not parsing_errors:
                            st.info("確認対象の文章が見つかりませんでした。")
                        else:
                            for item in all_evaluations:
                                if item.get("has_issue"):
                                    st.error(f"**元の文:** {item.get('original_sentence')}")
                                    st.markdown(f"**改善点の詳細:** {item.get('reason')}")
                                    st.markdown(f"**改善案:** {item.get('suggestion')}")
                                else:
                                    st.success(f"**元の文:** {item.get('original_sentence')}")
                                    st.markdown("_改善点なし_")
                                st.divider()
            else:
                st.error("ファイルをアップロードしてください。")
    except Exception as e:
        st.error(f"予期せぬエラーが発生しました: {e}")

if __name__ == "__main__":
    main()
