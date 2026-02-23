# ===================================================
# 所見生成AI セットアップスクリプト
# PowerShellで実行してください
# ===================================================

Write-Host "セットアップを開始します..." -ForegroundColor Cyan

# --- requirements.txt ---
@"
anthropic>=0.40.0
xlrd>=2.0.1
openpyxl>=3.1.0
tkinterdnd2>=0.3.0
"@ | Set-Content -Encoding UTF8 "requirements.txt"

# --- excel_loader.py ---
@"
"""
Excelファイルから所見サンプルデータを読み込むモジュール
"""
import os


def load_shoken_examples(filepath: str, max_examples: int = 5) -> list[str]:
    if not os.path.exists(filepath):
        return []
    try:
        import xlrd
    except ImportError:
        return []
    examples = []
    try:
        wb = xlrd.open_workbook(filepath)
        for sheet in wb.sheets():
            for row_idx in range(1, sheet.nrows):
                if sheet.ncols > 5:
                    text = str(sheet.cell(row_idx, 5).value).strip()
                    if text and text not in ("", "nan", "0.0"):
                        examples.append(text)
                if len(examples) >= max_examples:
                    break
            if len(examples) >= max_examples:
                break
    except Exception:
        pass
    return examples


def format_examples_for_prompt(examples: list[str]) -> str:
    if not examples:
        return ""
    sanitized = [ex.replace("{", "").replace("}", "") for ex in examples]
    lines = ["\n\n【参考例文（この文体・構成・長さを参考にしてください）】"]
    for i, example in enumerate(sanitized, 1):
        lines.append(f"例{i}:{example}")
    return "\n".join(lines)
"@ | Set-Content -Encoding UTF8 "excel_loader.py"

# --- templates.py ---
@"
"""
テンプレート定義モジュール
"""
import os

from excel_loader import format_examples_for_prompt, load_shoken_examples

_EXCEL_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "コピー5-1所見.xls")
_shoken_examples = load_shoken_examples(_EXCEL_PATH, max_examples=5)
_examples_text = format_examples_for_prompt(_shoken_examples)

TEMPLATES = {
    "shoken": {
        "name": "通知表の所見",
        "description": "小学校の通知表向けの所見を生成します（参考例文を使用）",
        "fields": [
            {
                "name": "grade",
                "label": "学年",
                "description": "生徒の学年（例: 3年生、5年生）",
            },
            {
                "name": "activities",
                "label": "係・委員会・行事での様子",
                "description": "係や委員会・行事での具体的な活動内容（例: 給食係として毎日丁寧に配膳準備を行い他の係の子と協力して取り組んだ）",
            },
            {
                "name": "subject_learning",
                "label": "学習面での様子",
                "description": "教科での頑張りや具体的なエピソード（例: 算数の分数の授業では難しい問題も諦めずに取り組み友達に解き方を説明できた）",
            },
            {
                "name": "target_length",
                "label": "文字数目安",
                "description": "所見の目安文字数（例: 150字、170字、180字）",
            },
        ],
        "system_prompt": (
            "あなたは小学校の担任教師です。通知表の所見文を書いてください。\n\n"
            "【構成】\n"
            "前半（1〜3文）：係・委員会・行事・掃除など生活面の具体的な様子\n"
            "後半（1〜3文）：教科学習での具体的な活動・気づき・成長\n"
            "合計3〜5文が目安。\n\n"
            "【文体・表現ルール】\n"
            "- 書き出しは全角スペースなし。「前期〇〇として、」「後期〇〇として、」「〇〇係として」など役割から始める\n"
            "- 前半→後半の切り替えは「学習面では、」または「〇〇「単元名」では、」で始める\n"
            "- 教科の書き方：教科名「単元名」では、/ 教科名科の「単元名」では、/ 教科名の授業では、\n"
            "- 文末は「〜ことができました。」を基本とし、「〜しました。」「〜でした。」を混ぜて単調にしない\n"
            "- 称賛語：一生懸命・率先して・積極的に・丁寧に・立派でした・すばらしいです・頼もしかったです\n"
            "- 否定的表現・課題指摘は一切使わない\n"
            "- 抽象的な誉め言葉だけでなく、必ず具体的なエピソードを添える\n"
            "- 生徒名は使わない\n\n"
            "【学年別表記】\n"
            "- 1〜2年生：全文ひらがなで書く（漢字不使用）\n"
            "- 3年生以上：漢字仮名交じり文\n\n"
            "【文字数】{target_length}程度（前後15字以内）\n"
            + _examples_text
        ),
        "user_prompt_template": (
            "以下の情報をもとに{grade}の通知表の所見を書いてください。\n\n"
            "係・委員会・行事での様子:\n{activities}\n\n"
            "学習面での様子:\n{subject_learning}\n\n"
            "所見文のみを出力してください（説明や前置きは不要です）。"
        ),
    },
}
"@ | Set-Content -Encoding UTF8 "templates.py"

# --- main.py ---
@"
#!/usr/bin/env python3
"""
通知表 所見生成AI
"""
import os
import sys

import anthropic

from templates import TEMPLATES


def display_menu(options: list[str], title: str) -> int:
    print(f"\n{title}")
    print("=" * 40)
    for i, option in enumerate(options, 1):
        print(f"  {i}. {option}")
    print("  0. 終了")
    print()
    while True:
        try:
            num = int(input("選択してください (番号): ").strip())
            if num == 0:
                return 0
            if 1 <= num <= len(options):
                return num
            print(f"1〜{len(options)} または 0 を入力してください。")
        except ValueError:
            print("数字を入力してください。")


def collect_field_inputs(fields: list[dict]) -> dict[str, str]:
    inputs = {}
    print()
    for field in fields:
        print(f"【{field['label']}】")
        print(f"  ヒント: {field['description']}")
        while True:
            value = input("  入力: ").strip()
            if value:
                inputs[field["name"]] = value
                break
            print("  入力が必要です。もう一度入力してください。")
        print()
    return inputs


def run_generation(client: anthropic.Anthropic, template_key: str) -> str:
    template = TEMPLATES[template_key]
    print(f"\n選択: {template['name']}")
    inputs = collect_field_inputs(template["fields"])
    system_prompt = template["system_prompt"].format(**inputs)
    user_prompt = template["user_prompt_template"].format(**inputs)

    print("\n" + "=" * 60)
    print("【生成結果】")
    print("=" * 60 + "\n")

    generated_text = []
    with client.messages.stream(
        model="claude-opus-4-6",
        max_tokens=4096,
        thinking={"type": "adaptive"},
        system=system_prompt,
        messages=[{"role": "user", "content": user_prompt}],
    ) as stream:
        for event in stream:
            if event.type == "content_block_delta":
                if event.delta.type == "text_delta":
                    print(event.delta.text, end="", flush=True)
                    generated_text.append(event.delta.text)

    print("\n\n" + "=" * 60)
    return "".join(generated_text)


def save_result(content: str, template_name: str) -> None:
    import datetime
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"output_{template_name}_{timestamp}.txt"
    with open(filename, "w", encoding="utf-8") as f:
        f.write(content)
    print(f"ファイルに保存しました: {filename}")


def main() -> None:
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        print("エラー: 環境変数 ANTHROPIC_API_KEY が設定されていません。")
        sys.exit(1)

    client = anthropic.Anthropic(api_key=api_key)

    print("\n" + "=" * 60)
    print("  通知表 所見生成AI  (Powered by Claude)")
    print("=" * 60)

    while True:
        template_keys = list(TEMPLATES.keys())
        menu_options = [
            f"{TEMPLATES[key]['name']} - {TEMPLATES[key]['description']}"
            for key in template_keys
        ]
        choice = display_menu(menu_options, "テンプレートを選択してください")
        if choice == 0:
            print("\nご利用ありがとうございました。")
            break

        template_key = template_keys[choice - 1]
        try:
            generated_text = run_generation(client, template_key)
            print("\n生成した文章をファイルに保存しますか？")
            if input("  (y/n): ").strip().lower() == "y":
                save_result(generated_text, template_key)
        except anthropic.AuthenticationError:
            print("\nエラー: APIキーが無効です。")
            sys.exit(1)
        except KeyboardInterrupt:
            print("\n\nキャンセルされました。")

        print("\n別の文章を生成しますか？")
        if input("  (y/n): ").strip().lower() != "y":
            print("\nご利用ありがとうございました。")
            break


if __name__ == "__main__":
    main()
"@ | Set-Content -Encoding UTF8 "main.py"

Write-Host ""
Write-Host "ファイルの作成が完了しました！" -ForegroundColor Green
Write-Host ""
Write-Host "次のステップ:" -ForegroundColor Yellow
Write-Host "  1. このフォルダに「コピー5-1所見.xls」をコピーしてください"
Write-Host "  2. pip install -r requirements.txt"
Write-Host ""
Write-Host "【GUIツール（Excel一括生成）を使う場合】" -ForegroundColor Cyan
Write-Host "  3a. python shoken_app.py"
Write-Host "      → ウィンドウが開くのでExcelをドロップして所見を一括生成"
Write-Host "      → 入力Excel: A=学年 B=係・行事 C=学習面 D=文字数目安"
Write-Host "      → 出力Excel: 元ファイル名_生成済み.xlsx のE列に所見が書き込まれます"
Write-Host ""
Write-Host "【コマンドラインで使う場合】" -ForegroundColor Cyan
Write-Host "  3b. `$env:ANTHROPIC_API_KEY = 'sk-ant-（あなたのキー）'"
Write-Host "      python main.py"
