"""
Excelファイルから所見サンプルデータを読み込むモジュール
"""
import os


def load_shoken_examples(filepath: str, max_examples: int = 5) -> list[str]:
    """
    Excelファイルから所見テキストのサンプルを読み込む。

    Args:
        filepath: Excelファイルのパス
        max_examples: 読み込むサンプル数の上限

    Returns:
        所見テキストのリスト（ファイルが存在しない場合は空リスト）
    """
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
            for row_idx in range(1, sheet.nrows):  # ヘッダー行をスキップ
                if sheet.ncols > 5:
                    text = str(sheet.cell(row_idx, 5).value).strip()
                    # 有効なテキストのみ追加（空・nan・数値をスキップ）
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
    """サンプルリストをプロンプト用の文字列に変換する"""
    if not examples:
        return ""

    # format() で誤認識されないよう波括弧を除去
    sanitized = [ex.replace("{", "").replace("}", "") for ex in examples]

    lines = ["\n\n【参考例文（この文体・構成・長さを参考にしてください）】"]
    for i, example in enumerate(sanitized, 1):
        lines.append(f"例{i}:{example}")
    return "\n".join(lines)
