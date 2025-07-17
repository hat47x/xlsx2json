import argparse
import json
import os
import re
import logging

logger = logging.getLogger("xlsx2json")
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

from openpyxl import load_workbook
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

    logger.info(f"Validation errors logged: {log_file}")


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


def is_empty_value(value: Any) -> bool:
    """
    値が空かどうかを判定する。
    None、空文字列、空のリスト、空のdictを空値として扱う。
    """
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    if isinstance(value, list) and len(value) == 0:
        return True
    if isinstance(value, dict) and len(value) == 0:
        return True
    return False


def is_completely_empty(obj: Union[Dict[str, Any], List[Any], Any]) -> bool:
    """
    オブジェクトが完全に空かどうかを再帰的に判定する。
    """
    if obj is None:
        return True
    if isinstance(obj, str) and obj.strip() == "":
        return True
    if isinstance(obj, dict):
        if not obj:
            return True
        return all(is_completely_empty(value) for value in obj.values())
    if isinstance(obj, list):
        if not obj:
            return True
        return all(is_completely_empty(item) for item in obj)
    return False


def clean_empty_values(obj: Union[Dict[str, Any], List[Any], Any], suppress_empty: bool = True) -> Union[Dict[str, Any], List[Any], Any, None]:
    """
    空の値を再帰的に除去する。
    suppress_empty が False の場合は何もしない。
    配列の場合は、要素全体が完全に空の場合のみインデックスを詰める。
    """
    if not suppress_empty:
        return obj
    
    if isinstance(obj, dict):
        cleaned = {}
        for key, value in obj.items():
            cleaned_value = clean_empty_values(value, suppress_empty)
            if not is_empty_value(cleaned_value):
                cleaned[key] = cleaned_value
        return cleaned if cleaned else None
    
    elif isinstance(obj, list):
        # 配列の場合：まず全要素を再帰的に処理
        processed_items = []
        for item in obj:
            processed_item = clean_empty_values(item, suppress_empty)
            processed_items.append(processed_item)
        
        # 完全に空の要素のみを除去して詰める
        # 部分的に空の要素（例：{"name": null, "address": "value"}）は保持
        cleaned = []
        for item in processed_items:
            if not is_completely_empty(item):
                cleaned.append(item)
        
        return cleaned if cleaned else None
    
    else:
        return obj if not is_empty_value(obj) else None


def clean_empty_arrays_contextually(obj: Union[Dict[str, Any], List[Any], Any], suppress_empty: bool = True) -> Union[Dict[str, Any], List[Any], Any, None]:
    """
    配列要素の整合性を保ちながら空値を除去する。
    配列要素全体が空の場合のみインデックスを詰める。
    """
    if not suppress_empty:
        return obj
    
    if isinstance(obj, dict):
        cleaned = {}
        for key, value in obj.items():
            cleaned_value = clean_empty_arrays_contextually(value, suppress_empty)
            if cleaned_value is not None:
                cleaned[key] = cleaned_value
        return cleaned if cleaned else None
    
    elif isinstance(obj, list):
        # 配列の各要素を再帰的に処理
        processed_items = []
        for item in obj:
            processed_item = clean_empty_arrays_contextually(item, suppress_empty)
            processed_items.append(processed_item)
        
        # 配列要素の整合性チェック
        # 全てがNoneの場合のみ、その要素を配列から除去
        cleaned = []
        for item in processed_items:
            if item is not None:
                cleaned.append(item)
            # Noneでも、配列の整合性を保つため、一部の要素だけが空の場合は保持する必要がある
            # この処理は、完全に空の要素のみを除去する
        
        return cleaned if cleaned else None
    
    else:
        return obj if not is_empty_value(obj) else None


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
    アンダーバーをワイルドカード（1文字）としてJSON Schemaの項目名と照合し、ユニークにマッチする場合はその項目名に置換する。
    """
    return parse_named_ranges_with_prefix(xlsx_path, prefix="json.")


def parse_named_ranges_with_prefix(xlsx_path: Path, prefix: str = "json.") -> Dict[str, Any]:
    """
    Excel 名前付き範囲(prefix) を解析してネスト dict/list を返す。
    prefixはデフォルトで"json."。
    """
    wb = load_workbook(xlsx_path, data_only=True)
    result: Dict[str, Any] = {}

    # schemaファイルがあれば読み込む
    global _global_schema
    schema = None
    try:
        schema = _global_schema
    except Exception:
        schema = None

    def match_schema_key(key: str, schema_props: dict) -> str:
        if not schema_props:
            return key
        key = key.strip()
        pattern = '^' + re.escape(key).replace('_', '.') + '$'
        matches = [prop for prop in schema_props if re.fullmatch(pattern, prop, flags=re.UNICODE)]
        logger.debug(f"key={key}, pattern={pattern}, matches={matches}")
        if len(matches) == 1:
            return matches[0]
        elif len(matches) > 1:
            logger.warning(f"ワイルドカード照合で複数マッチ: '{key}' → {matches}。ユニークでないため置換しません。")
        return key

    for name, defined_name in wb.defined_names.items():
        if not name.startswith(prefix):
            continue
        path_keys = name.removeprefix(prefix).split('.')

        if schema is not None:
            props = schema.get('properties', {})
            items = schema.get('items', {})
            new_keys = []
            current_schema = schema
            for k in path_keys:
                if re.fullmatch(r"\d+", k):
                    new_keys.append(k)
                    if isinstance(current_schema, dict) and 'items' in current_schema:
                        current_schema = current_schema['items']
                        props = current_schema.get('properties', {}) if isinstance(current_schema, dict) else {}
                        items = current_schema.get('items', {}) if isinstance(current_schema, dict) else {}
                    else:
                        props = {}
                        items = {}
                else:
                    if not props or not isinstance(props, dict):
                        logger.debug(f"props is empty or not dict at key={k}, break")
                        break
                    logger.debug(f"props.keys() at key={k}: {list(props.keys())}")
                    new_k = match_schema_key(k, props)
                    new_keys.append(new_k)
                    next_schema = props.get(new_k, {}) if isinstance(props, dict) else {}
                    if isinstance(next_schema, dict) and 'properties' in next_schema:
                        current_schema = next_schema
                        props = next_schema['properties']
                        items = next_schema.get('items', {})
                    elif isinstance(next_schema, dict) and 'items' in next_schema:
                        current_schema = next_schema
                        props = next_schema.get('properties', {})
                        items = next_schema['items']
                    else:
                        props = {}
                        items = {}
            path_keys = new_keys

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
            logger.warning(f"未処理のパス: {p_path}")
    return files


def write_json(
    data: Dict[str, Any],
    output_path: Path,
    schema: Optional[Dict[str, Any]] = None,
    validator: Optional[Draft7Validator] = None,
    suppress_empty: bool = True,
) -> None:
    """
    JSON をファイルに書き出し。
    バリデーションとソートはオプション。
    """
    base_name = output_path.stem
    output_dir = output_path.parent

    # 空値の除去
    if suppress_empty:
        data = clean_empty_values(data, suppress_empty)
        if data is None:
            data = {}

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

    logger.info(f"ファイル出力成功: {output_path}")


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
    parser.add_argument(
        '--keep-empty', action='store_true',
        help='空のセル値も JSON に含める (デフォルトでは空値を除去)'
    )
    parser.add_argument(
        '--log-level', default='INFO', choices=['DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL'],
        help='ログレベルを指定 (デフォルト: INFO)'
    )
    parser.add_argument(
        '--prefix', type=str, default='json.',
        help='Excel 名前付き範囲のプレフィックス (デフォルト: json.)'
    )
    args = parser.parse_args()

    # ログ設定
    logging.basicConfig(level=getattr(logging, args.log_level), format='[%(levelname)s] %(message)s')
    logger.setLevel(getattr(logging, args.log_level))

    xlsx_files = collect_xlsx_files(args.inputs)
    if not xlsx_files:
        logger.error("対象の .xlsx が見つかりませんでした。")
        return

    schema = load_schema(args.schema) if args.schema else None
    validator = Draft7Validator(schema) if schema else None
    suppress_empty = not args.keep_empty

    # schemaをグローバル変数にセット（parse_named_rangesで参照）
    global _global_schema
    _global_schema = schema

    for xlsx_path in xlsx_files:
        data = parse_named_ranges_with_prefix(xlsx_path, prefix=args.prefix)

        # 出力先設定
        out_dir = args.output_dir if args.output_dir else xlsx_path.parent / 'output-json'
        out_file = out_dir / f"{xlsx_path.stem}.json"
        write_json(data, out_file, schema, validator, suppress_empty)


if __name__ == '__main__':
    main()