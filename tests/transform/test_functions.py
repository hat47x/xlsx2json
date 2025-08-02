"""
Test Transform Functions

Functions specifically for unit testing purposes.
"""

from typing import List, Any


def simple_multiply(data: List[List], factor: float = 2.0) -> List[List]:
    """
    テスト用: 各セルの値に係数を掛ける

    Args:
        data: 二次元配列データ
        factor: 掛ける係数

    Returns:
        各セルに係数を掛けた二次元配列
    """
    result = []
    for row in data:
        new_row = []
        for cell in row:
            try:
                new_row.append(float(cell) * factor)
            except (ValueError, TypeError):
                new_row.append(cell)
        result.append(new_row)
    return result


def string_transform(value: Any) -> str:
    """
    テスト用: 文字列変換

    Args:
        value: 変換対象の値

    Returns:
        "TEST_" プレフィックス付きの文字列
    """
    return f"TEST_{value}"


def count_cells(data: List[List]) -> int:
    """
    テスト用: セル数をカウント

    Args:
        data: 二次元配列データ

    Returns:
        全セル数
    """
    count = 0
    for row in data:
        count += len(row)
    return count


def extract_first_column(data: List[List]) -> List[Any]:
    """
    テスト用: 最初の列を抽出

    Args:
        data: 二次元配列データ

    Returns:
        最初の列の値のリスト
    """
    return [row[0] if row else None for row in data]


def validate_numeric_range(
    data: List[List], min_val: float = 0, max_val: float = 100
) -> bool:
    """
    テスト用: 数値範囲の検証

    Args:
        data: 二次元配列データ
        min_val: 最小値
        max_val: 最大値

    Returns:
        すべての数値が範囲内ならTrue
    """
    for row in data:
        for cell in row:
            try:
                value = float(cell)
                if not (min_val <= value <= max_val):
                    return False
            except (ValueError, TypeError):
                continue
    return True
