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

# Azure OpenAIの設定
openai.api_type = "azure"
openai.api_key = os.getenv("AZURE_OPENAI_API_KEY")
openai.api_base = os.getenv("AZURE_OPENAI_ENDPOINT")
openai.api_version = os.getenv("AZURE_OPENAI_VERSION")
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

def extract_chapters(text):
    chapters = {}
    current_chapter = "不明"
    # 例：1. や 1.1 などの章番号を正規表現で探す
    for line in text.split('\n'):
        match = re.match(r'^(\d+(\.\d+)*)\s+(.*)', line)
        if match:
            current_chapter = f"{match.group(1)} {match.group(3)}"
        if current_chapter not in chapters:
            chapters[current_chapter] = ""
        chapters[current_chapter] += line + "\n"
    return chapters

def evaluate_text_with_llm(chapter_and_content, selected_rules):
    # selected_rulesの日本語マッピング
    rule_mapping = {
        "conciseness": "簡潔な文",
        "missing_elements": "必要な語の欠落",
        "ambiguity": "曖昧語の回避",
        "typos": "誤字脱字",
        "dependency": "係り受け"
    }
    active_rules = [rule_mapping[key] for key in selected_rules]

    # プロンプトの構築
    prompt = f"""
# 命令
あなたは、ソフトウェア要求仕様書の品質を評価する専門家です。
以下の「制約条件」と「評価対象の文章」に基づき、表現品質を評価してください。

# 制約条件
- まず、与えられた文章を1文ずつに分割してください。
- 分割した文の中から、ソフトウェアの要求や仕様に関連する文のみを抽出してください。要求仕様と関係ない文は無視してください。
- 抽出した各要求文に対して、以下の「表現品質評価ルール」を適用してください。
- 評価の結果、指摘事項があった文のみを結果として出力してください。指摘がない文は出力に含めないでください。
- 結果は必ずJSON形式で、以下の構造に従って出力してください。
  {{
    "results": [
      {{
        "chapter": "章番号・章名",
        "original_sentence": "元の記述",
        "reason": "指摘理由（どのルールに違反したか）",
        "suggestion": "改善案"
      }}
    ]
  }}

# 表現品質評価ルール
あなたがチェックすべきルールは以下の通りです： {', '.join(active_rules)}
- **簡潔な文**: 受動態、二重否定、部分否定、使役表現が使われていないか。
- **必要な語の欠落**: 主語、述語、目的語が明確で、欠落していないか。
- **曖昧語の回避**: 「適切に」「可能な限り」などの曖昧な表現が使われていないか。
- **誤字脱字**: 誤字や脱字がないか。
- **係り受け**: 係り受けが一意に解釈できるか。

# 評価対象の文章
{chapter_and_content}
"""

    try:
        response = openai.ChatCompletion.create(
            engine=DEPLOYMENT_NAME,
            messages=[
                {"role": "system", "content": "あなたは、ソフトウェア要求仕様書の品質を評価する専門家です。"},
                {"role": "user", "content": prompt}
            ],
            max_tokens=4096,
            temperature=0.0 # 再現性を高めるため
        )
        return response.choices[0].message['content']
    except Exception as e:
        return f'{{"error": "An error occurred: {str(e)}"}}'

import json

def parse_llm_response_to_markdown_table(responses):
    header = "| 章番号・章名 | 元の記述 | 指摘理由 | 改善案 |\n|--------------|-----------|-----------|---------|"
    all_rows = []

    for res_str in responses:
        try:
            # LLMの応答に含まれる可能性のあるマークダウンのコードブロック表記を削除
            if res_str.strip().startswith("```json"):
                res_str = res_str.strip()[7:-4]

            data = json.loads(res_str)
            if "error" in data:
                # エラーの場合はフッターに表示するためにNoneを返すなどの処理も可能
                continue

            results = data.get("results", [])
            for item in results:
                row = f"| {item.get('chapter', '')} | {item.get('original_sentence', '')} | {item.get('reason', '')} | {item.get('suggestion', '')} |"
                all_rows.append(row)
        except json.JSONDecodeError:
            # JSONパースエラーのハンドリング
            # ここではエラーを無視するが、ログに出力するなどの対応が望ましい
            continue
        except Exception:
            # その他の予期せぬエラー
            continue

    if not all_rows:
        return "指摘事項は見つかりませんでした。"

    return header + "\n" + "\n".join(all_rows)


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

                    chapters = extract_chapters(text)
                    all_results = []
                    progress_bar = st.progress(0)

                    for i, (chapter, content) in enumerate(chapters.items()):
                        # contentが空の場合はスキップ
                        if not content.strip():
                            continue

                        chapter_and_content = f"章: {chapter}\n内容:\n{content}"
                        llm_response = evaluate_text_with_llm(chapter_and_content, selected_rules)
                        all_results.append(llm_response)
                        progress_bar.progress((i + 1) / len(chapters))


                    # LLMの応答をパースしてMarkdownの表形式に変換
                    markdown_table = parse_llm_response_to_markdown_table(all_results)

                    st.success("評価が完了しました。")
                    st.markdown("### 評価結果")
                    st.markdown(markdown_table)

                    # ダウンロードボタン
                    st.download_button(
                        label="Markdown形式でダウンロード",
                        data=markdown_table,
                        file_name="evaluation_result.md",
                        mime="text/markdown",
                    )
            else:
                st.error("ファイルをアップロードしてください。")
    except Exception as e:
        st.error(f"予期せぬエラーが発生しました: {e}")

if __name__ == "__main__":
    main()
