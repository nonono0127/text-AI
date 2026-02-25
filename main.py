#!/usr/bin/env python3
"""
テンプレートベース文章生成AI

Claude APIを使用して、選択したテンプレートに基づいて
各種文章（ブログ記事、ビジネスメール、レポートなど）を生成します。

使用方法:
    python main.py

必要な環境変数:
    ANTHROPIC_API_KEY: Anthropic APIキー
"""

import os
import re
import sys

import anthropic

from templates import TEMPLATES


def display_menu(options: list[str], title: str) -> int:
    """メニューを表示してユーザーの選択を受け取る"""
    print(f"\n{title}")
    print("=" * 40)
    for i, option in enumerate(options, 1):
        print(f"  {i}. {option}")
    print("  0. 終了")
    print()

    while True:
        try:
            choice = input("選択してください (番号): ").strip()
            num = int(choice)
            if num == 0:
                return 0
            if 1 <= num <= len(options):
                return num
            print(f"1〜{len(options)} または 0 を入力してください。")
        except ValueError:
            print("数字を入力してください。")


def collect_field_inputs(fields: list[dict]) -> dict[str, str]:
    """テンプレートの各フィールドに対してユーザー入力を収集する"""
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


def generate_text(
    client: anthropic.Anthropic,
    system_prompt: str,
    user_prompt: str,
) -> None:
    """Claude APIを使用してテキストをストリーミング生成する"""
    print("\n" + "=" * 60)
    print("【生成結果】")
    print("=" * 60 + "\n")

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

    print("\n\n" + "=" * 60)


def save_result(content: str, template_name: str) -> None:
    """生成した文章をファイルに保存する"""
    import datetime

    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"output_{template_name}_{timestamp}.txt"

    with open(filename, "w", encoding="utf-8") as f:
        f.write(content)

    print(f"ファイルに保存しました: {filename}")


def run_generation(client: anthropic.Anthropic, template_key: str) -> str:
    """テンプレートを使って文章を生成し、生成されたテキストを返す"""
    template = TEMPLATES[template_key]
    print(f"\n選択: {template['name']}")
    print(f"説明: {template['description']}")

    # フィールド入力を収集
    inputs = collect_field_inputs(template["fields"])

    # 文字数範囲を計算（target_length が含まれるテンプレート用）
    if "target_length" in inputs:
        m = re.search(r'\d+', inputs["target_length"])
        target = int(m.group()) if m else 170
        inputs["min_length"] = target - 15
        inputs["max_length"] = target + 15

    # プロンプトを構築
    system_prompt = template["system_prompt"].format(**inputs)
    user_prompt = template["user_prompt_template"].format(**inputs)

    # テキストを生成してストリーミング表示しながら収集
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


def main() -> None:
    """メイン実行関数"""
    # APIキーの確認
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        print("エラー: 環境変数 ANTHROPIC_API_KEY が設定されていません。")
        print("例: export ANTHROPIC_API_KEY='your-api-key-here'")
        sys.exit(1)

    client = anthropic.Anthropic(api_key=api_key)

    print("\n" + "=" * 60)
    print("  テンプレートベース文章生成AI")
    print("  Powered by Claude claude-opus-4-6")
    print("=" * 60)

    while True:
        # テンプレート選択メニュー
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

            # 保存確認
            print("\n生成した文章をファイルに保存しますか？")
            save_choice = input("  (y/n): ").strip().lower()
            if save_choice == "y":
                save_result(generated_text, template_key)

        except anthropic.AuthenticationError:
            print("\nエラー: APIキーが無効です。ANTHROPIC_API_KEY を確認してください。")
            sys.exit(1)
        except anthropic.RateLimitError:
            print("\nエラー: APIのレート制限に達しました。しばらく待ってから再試行してください。")
        except anthropic.APIConnectionError:
            print("\nエラー: APIへの接続に失敗しました。ネットワーク接続を確認してください。")
        except KeyboardInterrupt:
            print("\n\n操作がキャンセルされました。")

        # 続けるか確認
        print("\n別の文章を生成しますか？")
        continue_choice = input("  (y/n): ").strip().lower()
        if continue_choice != "y":
            print("\nご利用ありがとうございました。")
            break


if __name__ == "__main__":
    main()
