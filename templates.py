"""
テンプレート定義モジュール

各テンプレートは以下を持ちます:
- name: テンプレート名
- description: テンプレートの説明
- fields: 入力フィールドのリスト（name, label, description）
- system_prompt: Claudeへのシステムプロンプト
- user_prompt_template: ユーザープロンプトのテンプレート文字列（{field_name}形式）
"""
import os

from excel_loader import format_examples_for_prompt, load_shoken_examples

# Excelファイルから所見サンプルを読み込む（スクリプトと同じディレクトリを参照）
_EXCEL_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "コピー5-1所見.xls")
_shoken_examples = load_shoken_examples(_EXCEL_PATH, max_examples=5)
_examples_text = format_examples_for_prompt(_shoken_examples)

TEMPLATES = {
    "blog": {
        "name": "ブログ記事",
        "description": "SEOを意識した読みやすいブログ記事を生成します",
        "fields": [
            {
                "name": "topic",
                "label": "テーマ・トピック",
                "description": "記事のメインテーマ（例: Pythonでのデータ分析入門）",
            },
            {
                "name": "target_audience",
                "label": "ターゲット読者",
                "description": "想定する読者層（例: プログラミング初心者、エンジニア）",
            },
            {
                "name": "tone",
                "label": "文体・トーン",
                "description": "記事の雰囲気（例: 親しみやすい、専門的、カジュアル）",
            },
            {
                "name": "length",
                "label": "文字数目安",
                "description": "記事の長さの目安（例: 800字、1500字）",
            },
        ],
        "system_prompt": (
            "あなたはプロのブログライターです。"
            "SEOを意識しながら、読みやすく魅力的なブログ記事を書いてください。"
            "見出し（##）や箇条書きを適切に使用し、構造化された記事を作成してください。"
        ),
        "user_prompt_template": (
            "以下の条件でブログ記事を書いてください。\n\n"
            "テーマ: {topic}\n"
            "ターゲット読者: {target_audience}\n"
            "文体・トーン: {tone}\n"
            "文字数目安: {length}\n\n"
            "記事には導入、本文（複数の見出し付きセクション）、まとめを含めてください。"
        ),
    },
    "email": {
        "name": "ビジネスメール",
        "description": "目的に合ったビジネスメールを生成します",
        "fields": [
            {
                "name": "purpose",
                "label": "メールの目的",
                "description": "メールの目的（例: 会議の日程調整、製品の提案、お礼）",
            },
            {
                "name": "recipient",
                "label": "宛先・相手",
                "description": "送り先（例: 取引先の山田部長、社内の鈴木さん）",
            },
            {
                "name": "key_points",
                "label": "伝えたい要点",
                "description": "メールで伝えたい主要な内容や情報",
            },
            {
                "name": "formality",
                "label": "丁寧さのレベル",
                "description": "敬語の度合い（例: 非常に丁寧、普通、カジュアル）",
            },
        ],
        "system_prompt": (
            "あなたはビジネスコミュニケーションの専門家です。"
            "明確で礼儀正しく、目的を達成するための効果的なビジネスメールを作成してください。"
            "件名、本文、締めの言葉を適切に含めてください。"
        ),
        "user_prompt_template": (
            "以下の条件でビジネスメールを作成してください。\n\n"
            "メールの目的: {purpose}\n"
            "宛先・相手: {recipient}\n"
            "伝えたい要点: {key_points}\n"
            "丁寧さのレベル: {formality}\n\n"
            "件名から締めの言葉まで、すぐに送信できる完全なメールを作成してください。"
        ),
    },
    "report": {
        "name": "ビジネスレポート",
        "description": "データや事実をまとめたビジネスレポートを生成します",
        "fields": [
            {
                "name": "subject",
                "label": "レポートの主題",
                "description": "レポートで扱うテーマ（例: 第3四半期の売上分析、新製品市場調査）",
            },
            {
                "name": "data_points",
                "label": "含める情報・データ",
                "description": "レポートに含める主要な情報や数値",
            },
            {
                "name": "conclusion",
                "label": "結論・提言",
                "description": "レポートで伝えたい結論や提案（任意）",
            },
        ],
        "system_prompt": (
            "あなたはプロのビジネスアナリストです。"
            "データと事実に基づいた客観的で明確なビジネスレポートを作成してください。"
            "エグゼクティブサマリー、本文、結論・提言の構成で作成してください。"
        ),
        "user_prompt_template": (
            "以下の条件でビジネスレポートを作成してください。\n\n"
            "主題: {subject}\n"
            "含める情報・データ: {data_points}\n"
            "結論・提言: {conclusion}\n\n"
            "エグゼクティブサマリーから始め、各セクションを明確な見出しで区切ってください。"
        ),
    },
    "product_description": {
        "name": "商品説明文",
        "description": "ECサイト向けの魅力的な商品説明文を生成します",
        "fields": [
            {
                "name": "product_name",
                "label": "商品名",
                "description": "商品の名前",
            },
            {
                "name": "features",
                "label": "主な特徴・スペック",
                "description": "商品の主要な特徴や仕様",
            },
            {
                "name": "target_customer",
                "label": "ターゲット顧客",
                "description": "この商品を買ってほしい顧客像",
            },
            {
                "name": "selling_point",
                "label": "最大のセールスポイント",
                "description": "競合と差別化できる最大の強み",
            },
        ],
        "system_prompt": (
            "あなたはECサイトの商品コピーライターです。"
            "顧客が購買意欲を持つような、魅力的で説得力のある商品説明文を作成してください。"
            "感情に訴えながらも、具体的なメリットを伝えてください。"
        ),
        "user_prompt_template": (
            "以下の商品の説明文を作成してください。\n\n"
            "商品名: {product_name}\n"
            "主な特徴・スペック: {features}\n"
            "ターゲット顧客: {target_customer}\n"
            "最大のセールスポイント: {selling_point}\n\n"
            "キャッチコピー、商品説明（200〜300字）、主な特徴リスト（箇条書き）の順で作成してください。"
        ),
    },
    "custom": {
        "name": "カスタム（自由入力）",
        "description": "システムプロンプトとユーザープロンプトを自由に設定して文章を生成します",
        "fields": [
            {
                "name": "system_instruction",
                "label": "AIへの指示（システムプロンプト）",
                "description": "AIにどのような役割・スタイルで回答させるか",
            },
            {
                "name": "request",
                "label": "生成したい文章の内容",
                "description": "どのような文章を生成してほしいか詳しく記述",
            },
        ],
        "system_prompt": "{system_instruction}",
        "user_prompt_template": "{request}",
    },
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
                "description": (
                    "係や委員会・行事での具体的な活動内容"
                    "（例: 給食係として毎日丁寧に配膳準備を行い他の係の子と協力して取り組んだ）"
                ),
            },
            {
                "name": "subject_learning",
                "label": "学習面での様子",
                "description": (
                    "教科での頑張りや具体的なエピソード"
                    "（例: 算数の分数の授業では難しい問題も諦めずに取り組み友達に解き方を説明できた）"
                ),
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
            "- 教科の書き方：教科名科「単元名」では、（例：図画工作科「版画」では、/ 算数科「割合」では、）/ 教科名の授業では、\n"
            "- 文末は「〜ことができました。」を基本とし、「〜しました。」「〜でした。」「〜れました。」「〜いました。」を混ぜて単調にしない\n"
            "- 称賛語：一生懸命・率先して・積極的に・丁寧に・立派でした・すばらしいです・責任をもって・欠かさず・粘り強く・先を見通して\n"
            "- 否定的表現・課題指摘は一切使わない\n"
            "- 「素敵です」「感心しました」「頼もしかった」等の教師の主観的感想は使わない\n"
            "- 抽象的な誉め言葉だけでなく、必ず具体的なエピソードを添える\n"
            "- 生徒名は使わない\n\n"
            "【学年別表記】\n"
            "- 1〜2年生：全文ひらがなで書く（漢字不使用）\n"
            "- 3年生以上：漢字仮名交じり文\n\n"
            "【文字数】必ず{min_length}〜{max_length}字の範囲で書くこと\n"
            + _examples_text
        ),
        "user_prompt_template": (
            "以下の情報をもとに{grade}の通知表の所見を書いてください。\n\n"
            "係・委員会・行事での様子:\n{activities}\n\n"
            "学習面での様子:\n{subject_learning}\n\n"
            "文字数は必ず{min_length}〜{max_length}字の範囲に収めてください。\n"
            "所見文のみを出力してください（説明や前置きは不要です）。"
        ),
    },
}
