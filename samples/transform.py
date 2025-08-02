# ユーザ定義関数のサンプル集


# =============================================================================
# 文字列変換
# =============================================================================


def csv(value):
    """CSV文字列を配列に分割"""
    if not isinstance(value, str):
        return value
    import csv
    from io import StringIO

    reader = csv.reader(StringIO(value))
    return [row for row in reader if any(cell.strip() for cell in row)]


def lines(value):
    """改行区切りの文字列を配列に分割"""
    if not isinstance(value, str):
        return value
    return [line.strip() for line in value.split("\n") if line.strip()]


def words(value):
    """空白区切りの文字列を配列に分割"""
    if not isinstance(value, str):
        return value
    return [w for w in value.split() if w]


# =============================================================================
# 配列・行列操作
# =============================================================================


def column(data, index=0):
    """指定列を抽出"""
    if not isinstance(data, list):
        return data

    try:
        return [
            row[index] if isinstance(row, list) and len(row) > index else None
            for row in data
        ]
    except (IndexError, TypeError):
        return []


def sum_col(data, index=0):
    """指定列の合計を計算"""
    if not isinstance(data, list):
        return 0

    try:
        total = 0
        for row in data:
            if isinstance(row, list) and len(row) > index:
                try:
                    total += float(row[index])
                except (ValueError, TypeError):
                    continue
        return total
    except (IndexError, TypeError):
        return 0


def flip(data):
    """行と列を入れ替え（転置）"""
    if not isinstance(data, list) or not data:
        return []

    # 最大列数を取得
    max_cols = max(len(row) if isinstance(row, list) else 0 for row in data)

    # 転置行列を作成
    result = []
    for col_index in range(max_cols):
        new_row = []
        for row in data:
            if isinstance(row, list):
                value = row[col_index] if col_index < len(row) else None
                new_row.append(value)
        result.append(new_row)

    return result


def clean(data):
    """空でない行のみを残す"""
    if not isinstance(data, list):
        return data

    result = []
    for row in data:
        if isinstance(row, list):
            # 行に空でない値が少なくとも1つあるかチェック
            has_data = any(cell is not None and str(cell).strip() for cell in row)
            if has_data:
                result.append(row)

    return result


# =============================================================================
# 数値計算
# =============================================================================


def total(data):
    """全要素の合計"""
    if not isinstance(data, list):
        return 0

    result = 0
    for row in data:
        if isinstance(row, list):
            for cell in row:
                try:
                    result += float(cell)
                except (ValueError, TypeError):
                    continue
    return result


def avg(data):
    """数値要素の平均"""
    if not isinstance(data, list):
        return 0

    total_val = 0
    count = 0
    for row in data:
        if isinstance(row, list):
            for cell in row:
                try:
                    total_val += float(cell)
                    count += 1
                except (ValueError, TypeError):
                    continue

    return total_val / count if count > 0 else 0


# =============================================================================
# 便利関数
# =============================================================================


def parse_json(value):
    """JSON文字列を解析"""
    import json

    try:
        return json.loads(value)
    except Exception:
        return value


def normalize(value):
    """文字列を正規化（トリム・全角半角変換・置換など）"""
    if not isinstance(value, str):
        return value

    import re
    import unicodedata

    # 前後の空白除去
    value = value.strip()
    # 全角→半角変換
    value = unicodedata.normalize("NFKC", value)
    # 「株式会社」を「(株)」に変換
    value = value.replace("株式会社", "(株)")
    # 複数スペースを1つに統一
    value = re.sub(r"\s+", " ", value)

    return value


def upper(value):
    """大文字に変換"""
    return str(value).upper() if value else value


def lower(value):
    """小文字に変換"""
    return str(value).lower() if value else value


# =============================================================================
# 後方互換性（旧関数名）
# =============================================================================

# 従来の関数名のエイリアス
csv_split = csv
line_split = lines
word_split = words
extract_column = column
sum_column = sum_col
transpose_matrix = flip
filter_non_empty_rows = clean
matrix_sum = total
calculate_average = avg
string_transform = normalize
json_parse = parse_json
