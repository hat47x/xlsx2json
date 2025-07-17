import argparse
import json
import os
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

from openpyxl import load_workbook
import jsonschema
from jsonschema import Draft7Validator


def load_schema(schema_path: Optional[Path]) -> Optional[Dict[str, Any]]:
    """
    指定されたパスから JSON スキーマを読み込む。
    スキーマが指定されていない場合は None を返す。
    """
    if not schema_path:
        return None

    with schema_path.open('r', encoding='utf-8') as f:
        return json.load(f)


def validate_and_log(
    data: Dict[str, Any],
    validator: Draft7Validator,
    log_dir: Path,
    base_name: str
) -> None:
    """
    JSON データをバリデートし、エラーがあれば .error.log ファイルに出力する。
    """
    errors = sorted(validator.iter_errors(data), key=lambda e: e.path)
    if not errors:
        return

    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"{base_name}.error.log"
    with log_file.open('w', encoding='utf-8') as f:
        for err in errors:
            path = '.'.join(str(p) for p in err.path)
            f.write(f"{path}: {err.message}\n")

    print(f"Validation errors logged: {log_file}")


def reorder_json(
    obj: Union[Dict[str, Any], List[Any], Any],
    schema: Dict[str, Any]
) -> Union[Dict[str, Any], List[Any], Any]:
    """
    スキーマの properties 順に dict のキーを再帰的に並べ替える。
    list の場合は項目ごとに再帰処理。
    その他はそのまま返す。
    """
    if isinstance(obj, dict) and isinstance(schema, dict):
        ordered: Dict[str, Any] = {}
        props = schema.get('properties', {})
        # スキーマ順に追加
        for key in props:
            if key in obj:
                ordered[key] = reorder_json(obj[key], props[key])
        # 追加キーはアルファベット順
        for key in sorted(k for k in obj if k not in props):
            ordered[key] = obj[key]
        return ordered

    if isinstance(obj, list) and isinstance(schema, dict) and 'items' in schema:
        return [reorder_json(item, schema['items']) for item in obj]

    return obj


def get_named_range_values(wb, defined_name) -> Any:
    """
    Excel の NamedRange からセル値を抽出し、単一セルは値、範囲はリストで返す。
    """
    values: List[Any] = []
    for sheet_name, coord in defined_name.destinations:
        cell_or_range = wb[sheet_name][coord]
        if isinstance(cell_or_range, tuple):  # 範囲
            for row in cell_or_range:
                for cell in row:
                    values.append(cell.value)
        else:
            values.append(cell_or_range.value)

    return values[0] if len(values) == 1 else values


def insert_json_path(root: Union[Dict[str, Any], List[Any]], keys: List[str], value: Any) -> None:
    """
    ドット区切りキーのリストから JSON 構造を構築し、値を挿入する。
    数字キーは list、文字列キーは dict として扱う。
    """
    key = keys[0]
    is_last = len(keys) == 1

    if re.fullmatch(r"\d+", key):
        idx = int(key) - 1
        if not isinstance(root, list):
            raise TypeError(f"Expected list at {keys}, got {type(root)}")
        while len(root) <= idx:
            root.append(None)
        if is_last:
            root[idx] = value
        else:
            if root[idx] is None:
                root[idx] = [] if re.fullmatch(r"\d+", keys[1]) else {}
            insert_json_path(root[idx], keys[1:], value)
    else:
        if not isinstance(root, dict):
            raise TypeError(f"Expected dict at {keys}, got {type(root)}")
        if is_last:
            root[key] = value
        else:
            if key not in root:
                root[key] = [] if re.fullmatch(r"\d+", keys[1]) else {}
            insert_json_path(root[key], keys[1:], value)


def parse_named_ranges(xlsx_path: Path) -> Dict[str, Any]:
    """
    Excel 名前付き範囲(json.*) を解析してネスト dict/list を返す。
    """
    wb = load_workbook(xlsx_path, data_only=True)
    result: Dict[str, Any] = {}

    for name, defined_name in wb.defined_names.items():
        if not name.startswith("json."):
            continue
        path_keys = name.removeprefix("json.").split('.')
        value = get_named_range_values(wb, defined_name)
        insert_json_path(result, path_keys, value)

    return result


def collect_xlsx_files(paths: List[str]) -> List[Path]:
    """
    ファイルまたはディレクトリのリストから、対象となる .xlsx ファイル一覧を取得。
    ディレクトリ指定時は直下のみ。
    """
    files: List[Path] = []
    for p in paths:
        p_path = Path(p)
        if p_path.is_dir():
            for entry in p_path.iterdir():
                if entry.suffix.lower() == '.xlsx':
                    files.append(entry)
        elif p_path.is_file() and p_path.suffix.lower() == '.xlsx':
            files.append(p_path)
        else:
            print(f"Warning: 未処理のパス: {p_path}")
    return files


def write_json(
    data: Dict[str, Any],
    output_path: Path,
    schema: Optional[Dict[str, Any]] = None,
    validator: Optional[Draft7Validator] = None,
) -> None:
    """
    JSON をファイルに書き出し。
    バリデーションとソートはオプション。
    """
    base_name = output_path.stem
    output_dir = output_path.parent

    # バリデーション → エラーログ
    if validator:
        validate_and_log(data, validator, output_dir, base_name)

    # ソート処理
    if schema:
        data = reorder_json(data, schema)

    # ファイル書き出し
    output_dir.mkdir(parents=True, exist_ok=True)
    with output_path.open('w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"Written: {output_path}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Excel 名前付き範囲(json.*) -> JSON ファイル出力"
    )
    parser.add_argument(
        'inputs', nargs='+',
        help='変換対象のファイルまたはフォルダ (.xlsx)'
    )
    parser.add_argument(
        '--output-dir', '-o', type=Path,
        help='一括出力先フォルダ (省略時は各入力ファイル隣の output-json)'
    )
    parser.add_argument(
        '--schema', '-s', type=Path,
        help='JSON Schema ファイル (バリデーションとソート用)'
    )
    args = parser.parse_args()

    xlsx_files = collect_xlsx_files(args.inputs)
    if not xlsx_files:
        print("Error: 対象の .xlsx が見つかりませんでした。")
        return

    schema = load_schema(args.schema) if args.schema else None
    validator = Draft7Validator(schema) if schema else None

    for xlsx_path in xlsx_files:
        data = parse_named_ranges(xlsx_path)

        # 出力先設定
        out_dir = args.output_dir if args.output_dir else xlsx_path.parent / 'output-json'
        out_file = out_dir / f"{xlsx_path.stem}.json"
        write_json(data, out_file, schema, validator)


if __name__ == '__main__':
    main()
