# --transformオプションで参照する関数のサンプル


# CSV分割
def csv_split(value):
    if not isinstance(value, str):
        return value
    import csv
    from io import StringIO

    reader = csv.reader(StringIO(value))
    return [row for row in reader if any(cell.strip() for cell in row)]


# 行分割（改行で分割）
def line_split(value):
    if not isinstance(value, str):
        return value
    return [line.strip() for line in value.split("\n") if line.strip()]


# 単語分割（空白で分割）
def word_split(value):
    if not isinstance(value, str):
        return value
    return [w for w in value.split() if w]


# JSONパース（文字列→オブジェクト）
def json_parse(value):
    import json

    try:
        return json.loads(value)
    except Exception:
        return value


# 文字列→文字列変換（例：全角→半角、トリム、置換など）
def string_transform(value):
    if not isinstance(value, str):
        return value
    # 例: 前後の空白を除去し、全角英数字を半角に変換
    import re
    import unicodedata

    # 前後の空白除去
    value = value.strip()
    # 全角→半角変換
    value = unicodedata.normalize("NFKC", value)
    # 例: 特定文字の置換（「株式会社」を「(株)」に）
    value = value.replace("株式会社", "(株)")
    # 例: 複数スペースを1つに
    value = re.sub(r"\s+", " ", value)
    # 例: 文字列追加
    value = f"変換後: {value}"
    return value
