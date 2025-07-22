"""
xlsx2json - Excel の名前付き範囲を JSON に変換するツール
"""

import argparse
import datetime
import json
import re
import logging
import shlex
import subprocess
import importlib.util
from pathlib import Path
from typing import Any, Dict, List, Optional, Union, Callable

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
from jsonschema import Draft7Validator


# グローバル変数
_global_trim = False
_global_schema = None

logger = logging.getLogger("xlsx2json")


class ProcessingStats:
    """処理統計情報を管理するクラス"""
    
    def __init__(self):
        self.reset()
    
    def reset(self):
        self.containers_processed = 0
        self.cells_generated = 0
        self.cells_read = 0
        self.empty_cells_skipped = 0
        self.errors = []
        self.warnings = []
        self.start_time = None
        self.end_time = None
    
    def start_processing(self):
        import time
        self.start_time = time.time()
    
    def end_processing(self):
        import time
        self.end_time = time.time()
    
    def add_error(self, error_msg):
        self.errors.append(error_msg)
        logger.error(error_msg)
    
    def add_warning(self, warning_msg):
        self.warnings.append(warning_msg)
        logger.warning(warning_msg)
    
    def get_duration(self):
        if self.start_time and self.end_time:
            return self.end_time - self.start_time
        return 0
    
    def log_summary(self):
        """処理結果のサマリをログ出力"""
        duration = self.get_duration()
        logger.info("=" * 50)
        logger.info("処理統計サマリ")
        logger.info("=" * 50)
        logger.info(f"処理時間: {duration:.2f}秒")
        logger.info(f"処理されたコンテナ数: {self.containers_processed}")
        logger.info(f"生成されたセル名数: {self.cells_generated}")
        logger.info(f"読み取られたセル数: {self.cells_read}")
        logger.info(f"スキップされた空セル数: {self.empty_cells_skipped}")
        logger.info(f"エラー数: {len(self.errors)}")
        logger.info(f"警告数: {len(self.warnings)}")
        
        if self.errors:
            logger.info("\nエラー詳細:")
            for error in self.errors[-5:]:  # 最新5件のみ表示
                logger.info(f"  - {error}")
        
        if self.warnings:
            logger.info("\n警告詳細:")
            for warning in self.warnings[-5:]:  # 最新5件のみ表示
                logger.info(f"  - {warning}")
        
        logger.info("=" * 50)


# グローバル統計インスタンス
processing_stats = ProcessingStats()


# =============================================================================
# Core Utilities
# =============================================================================


def load_schema(schema_path: Optional[Path]) -> Optional[Dict[str, Any]]:
    """
    指定されたパスから JSON スキーマを読み込む。
    スキーマが指定されていない場合は None を返す。
    """
    if not schema_path:
        return None

    if not schema_path.exists():
        raise FileNotFoundError(f"スキーマファイルが見つかりません: {schema_path}")

    if not schema_path.is_file():
        raise ValueError(f"指定されたパスはファイルではありません: {schema_path}")

    try:
        with schema_path.open("r", encoding="utf-8") as f:
            return json.load(f)
    except json.JSONDecodeError as e:
        # 詳細なエラー情報をログに出力し、元の例外を再発生
        logger.error(f"無効なJSONフォーマットです: {schema_path} - {e}")
        raise  # 元のJSONDecodeErrorを再発生
    except Exception as e:
        raise IOError(f"スキーマファイルの読み込みに失敗しました: {schema_path} - {e}")


# =============================================================================
# Schema Operations
# =============================================================================


def validate_and_log(
    data: Dict[str, Any], validator: Draft7Validator, log_dir: Path, base_name: str
) -> None:
    """
    JSON データをバリデートし、エラーがあれば .error.log ファイルに出力する。
    """
    errors = sorted(validator.iter_errors(data), key=lambda e: e.path)
    if not errors:
        return

    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"{base_name}.error.log"
    with log_file.open("w", encoding="utf-8") as f:
        for err in errors:
            path = ".".join(str(p) for p in err.path)
            f.write(f"{path}: {err.message}\n")

    logger.debug(f"Validation errors written to: {log_file}")


# =============================================================================
# Data Validation and Cleaning
# =============================================================================


def reorder_json(
    obj: Union[Dict[str, Any], List[Any], Any], schema: Dict[str, Any]
) -> Union[Dict[str, Any], List[Any], Any]:
    """
    スキーマの properties 順に dict のキーを再帰的に並べ替える。
    list の場合は項目ごとに再帰処理。
    その他はそのまま返す。
    """
    if isinstance(obj, dict) and isinstance(schema, dict):
        ordered: Dict[str, Any] = {}
        props = schema.get("properties", {})
        # スキーマ順に追加
        for key in props:
            if key in obj:
                ordered[key] = reorder_json(obj[key], props[key])
        # 追加キーはアルファベット順
        for key in sorted(k for k in obj if k not in props):
            ordered[key] = obj[key]
        return ordered

    if isinstance(obj, list) and isinstance(schema, dict) and "items" in schema:
        return [reorder_json(item, schema["items"]) for item in obj]

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

    # 1セルなら値のみ返す（listでなく）
    if len(values) == 1:
        return values[0]
    return values


# =============================================================================
# Container Support Functions
# =============================================================================

def parse_range(range_str: str) -> tuple:
    """
    Excel範囲文字列を解析して開始座標と終了座標を返す
    例: "B2:D4" -> ((2, 2), (4, 4))
    例: "$B$2:$D$4" -> ((2, 2), (4, 4))
    """
    import re
    
    # $記号を削除してから解析
    cleaned_range = range_str.replace('$', '')
    
    match = re.match(r'^([A-Z]+)(\d+):([A-Z]+)(\d+)$', cleaned_range)
    if not match:
        raise ValueError(f"無効な範囲形式: {range_str}")
    
    start_col, start_row, end_col, end_row = match.groups()
    start_coord = (column_index_from_string(start_col), int(start_row))
    end_coord = (column_index_from_string(end_col), int(end_row))
    
    return start_coord, end_coord


def detect_instance_count(start_coord: tuple, end_coord: tuple, direction: str) -> int:
    """
    範囲とdirectionから、インスタンス数を検出
    """
    start_col, start_row = start_coord
    end_col, end_row = end_coord
    
    if direction == "vertical":
        return end_row - start_row + 1
    elif direction == "horizontal":
        return end_col - start_col + 1
    elif direction == "column":
        return end_col - start_col + 1
    else:
        raise ValueError(f"無効なdirection: {direction}")


def generate_cell_names(container_name: str, start_coord: tuple, end_coord: tuple, 
                       direction: str, items: list) -> list:
    """
    コンテナ用のセル名を生成
    1-base indexingを使用
    """
    start_col, start_row = start_coord
    end_col, end_row = end_coord
    cell_names = []
    
    instance_count = detect_instance_count(start_coord, end_coord, direction)
    
    for i in range(1, instance_count + 1):  # 1-base indexing
        for item in items:
            # フォーマット: コンテナ名_インデックス_アイテム名
            cell_name = f"{container_name}_{i}_{item}"
            cell_names.append(cell_name)
    
    return cell_names


def load_container_config(config_path: Path) -> Dict[str, Any]:
    """
    config.jsonからコンテナ設定を読み込む
    """
    if not config_path.exists():
        return {}
    
    try:
        with config_path.open("r", encoding="utf-8") as f:
            config = json.load(f)
            return config.get("containers", {})
    except (json.JSONDecodeError, FileNotFoundError):
        logger.warning(f"コンテナ設定の読み込みに失敗: {config_path}")
        return {}


def resolve_container_range(wb, range_spec: str) -> tuple:
    """
    範囲指定を解決して座標を返す
    range_specは名前付き範囲名または範囲文字列（"A1:C10"）
    """
    # 名前付き範囲として試行
    if range_spec in wb.defined_names:
        defined_name = wb.defined_names[range_spec]
        for sheet_name, coord in defined_name.destinations:
            return parse_range(coord)
    
    # 範囲文字列として解析
    try:
        return parse_range(range_spec)
    except ValueError:
        raise ValueError(f"無効な範囲指定: {range_spec}")


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


def clean_empty_values(
    obj: Union[Dict[str, Any], List[Any], Any], suppress_empty: bool = True
) -> Union[Dict[str, Any], List[Any], Any, None]:
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


def clean_empty_arrays_contextually(
    obj: Union[Dict[str, Any], List[Any], Any], suppress_empty: bool = True
) -> Union[Dict[str, Any], List[Any], Any, None]:
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


# =============================================================================
# JSON Path Operations
# =============================================================================


def insert_json_path(
    root: Union[Dict[str, Any], List[Any]],
    keys: List[str],
    value: Any,
    full_path: str = "",
) -> None:
    """
    ドット区切りキーのリストから JSON 構造を構築し、値を挿入する。
    数字キーは list、文字列キーは dict として扱う。
    """
    # 空のパスの場合のエラーハンドリング
    if not keys:
        raise ValueError(
            "JSONパスが空です。値を挿入するには少なくとも1つのキーが必要です。"
        )

    key = keys[0]
    is_last = len(keys) == 1
    current_path = f"{full_path}.{key}" if full_path else key

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
            insert_json_path(root[idx], keys[1:], value, current_path)
    else:
        # dict型でない場合は、呼び出し元の親dictの該当キーをdictに置き換える
        if not isinstance(root, dict):
            raise TypeError(f"insert_json_path: root must be dict, got {type(root)}")
        if is_last:
            root[key] = value
        else:
            # 既存値がstr型などの場合はdictに置き換え、元の値を'__value__'キーに退避
            if key in root and not isinstance(root[key], (dict, list)):
                prev_value = root[key]
                logger.warning(
                    f"パス重複が検出されました: '{current_path}' - "
                    f"既存値を '__value__' キーに退避します (値: {prev_value})"
                )
                root[key] = {"__value__": prev_value}
            
            if key not in root:
                root[key] = [] if re.fullmatch(r"\d+", keys[1]) else {}
            else:
                # 既存値の型チェックと変換
                next_key_is_numeric = re.fullmatch(r"\d+", keys[1]) if len(keys) > 1 else False
                
                if next_key_is_numeric and isinstance(root[key], dict) and not root[key]:
                    # 空辞書を配列に変換
                    root[key] = []
                    logger.debug(f"空辞書を配列に変換: {current_path}")
                elif not next_key_is_numeric and isinstance(root[key], list) and not root[key]:
                    # 空配列を辞書に変換
                    root[key] = {}
                    logger.debug(f"空配列を辞書に変換: {current_path}")
                    
            insert_json_path(root[key], keys[1:], value, current_path)


def parse_array_split_rules(
    array_split_rules: List[str], prefix: str
) -> Dict[str, List[str]]:
    """
    配列化設定のパース。
    形式: "json.parent.1=,|;" (複数の区切り文字を|で区切る)
    パイプ文字自体を区切り文字にする場合は \\| でエスケープする。
    プレフィックスは自動的に削除される。

    戻り値: {path: [delimiter1, delimiter2, ...]}
    """
    if not prefix or not isinstance(prefix, str):
        raise ValueError("prefixは空ではない文字列である必要があります。")

    rules = {}
    for rule in array_split_rules:
        if not rule or not isinstance(rule, str):
            logger.warning(f"無効なルール形式をスキップします: {rule}")
            continue

        if "=" not in rule:
            logger.warning(f"無効な配列化設定: {rule}")
            continue

        path, delimiter_str = rule.split("=", 1)
        # プレフィックスを削除
        if path.startswith(prefix):
            path = path.removeprefix(prefix)

        # エスケープされたパイプ文字を一時的に置換
        temp_placeholder = "___ESCAPED_PIPE___"
        delimiter_str = delimiter_str.replace("\\|", temp_placeholder)

        # 複数の区切り文字を|で区切る
        delimiter_parts = delimiter_str.split("|")
        delimiters = []
        for part in delimiter_parts:
            # エスケープ文字の処理
            processed_delimiter = (
                part.replace("\\n", "\n").replace("\\t", "\t").replace("\\r", "\r")
            )
            # エスケープされたパイプ文字を元に戻す
            processed_delimiter = processed_delimiter.replace(temp_placeholder, "|")
            delimiters.append(processed_delimiter)

        rules[path] = delimiters

    return rules


# =============================================================================
# Array Transform Rules
# =============================================================================


class ArrayTransformRule:
    """配列変換ルールを表すクラス"""

    def __init__(self, path: str, transform_type: str, transform_spec: str):
        # パラメータの基本検証
        if not path or not isinstance(path, str):
            raise ValueError("pathは空ではない文字列である必要があります。")
        if not transform_type or not isinstance(transform_type, str):
            raise ValueError("transform_typeは空ではない文字列である必要があります。")
        if not transform_spec or not isinstance(transform_spec, str):
            raise ValueError("transform_specは空ではない文字列である必要があります。")

        self.path = path
        self.transform_type = transform_type  # 'function', 'command', 'split'
        self.transform_spec = transform_spec
        self._transform_func: Optional[Callable] = None
        self._setup_transform()

    def _setup_transform(self):
        """変換関数をセットアップ"""
        if self.transform_type == "function":
            self._setup_python_function()
        elif self.transform_type == "command":
            self._setup_command()
        elif self.transform_type == "split":
            # splitは関数を外部でセットするので何もしない
            pass
        else:
            raise ValueError(f"Unknown transform type: {self.transform_type}")

    def _setup_python_function(self):
        """Python関数のセットアップ"""
        # 形式: module_path:function_name または file_path:function_name
        if ":" not in self.transform_spec:
            raise ValueError(
                f"Python function spec must be 'module:function' or 'file.py:function': {self.transform_spec}"
            )

        module_or_file, func_name = self.transform_spec.rsplit(":", 1)

        try:
            if module_or_file.endswith(".py"):
                # ファイルから関数を読み込み
                file_path = Path(module_or_file)
                if not file_path.exists():
                    raise FileNotFoundError(f"Transform file not found: {file_path}")

                spec = importlib.util.spec_from_file_location(
                    "transform_module", file_path
                )
                if spec is None or spec.loader is None:
                    raise ImportError(f"Cannot load module from {file_path}")

                module = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(module)
                self._transform_func = getattr(module, func_name)
            else:
                # モジュールから関数を読み込み
                module = importlib.import_module(module_or_file)
                self._transform_func = getattr(module, func_name)

            logger.info(f"Loaded transform function: {self.transform_spec}")
        except Exception as e:
            raise ValueError(
                f"Failed to load transform function '{self.transform_spec}': {e}"
            )

    def _setup_command(self):
        """外部コマンドのセットアップ"""
        # コマンドが実行可能かチェック
        try:
            result = subprocess.run(
                self.transform_spec.split()[0],
                capture_output=True,
                text=True,
                input="test",
                timeout=5,
            )
            logger.info(f"Command available: {self.transform_spec}")
        except (subprocess.TimeoutExpired, FileNotFoundError) as e:
            logger.warning(f"Command check failed for '{self.transform_spec}': {e}")

    def transform(self, value: Any) -> Any:
        """値を変換"""
        global _global_trim
        if self.transform_type == "function":
            result = self._transform_func(value)
            # --trim指定時は配列要素をstrip()
            if _global_trim and isinstance(result, list):
                return [v.strip() if isinstance(v, str) else v for v in result]
            return result
        elif self.transform_type == "command":
            return self._transform_with_command(value)
        elif self.transform_type == "split":
            if isinstance(value, list):
                return [self._transform_func(v) for v in value]
            else:
                return self._transform_func(value)
        else:
            return value

    def _transform_with_command(self, value: Any) -> Any:
        """外部コマンドで変換"""
        try:
            input_str = str(value) if value is not None else ""
            result = subprocess.run(
                shlex.split(self.transform_spec),
                input=input_str,
                capture_output=True,
                text=True,
                timeout=30,
            )

            if result.returncode != 0:
                logger.warning(
                    f"Command failed: {self.transform_spec}, stderr: {result.stderr}"
                )
                return value

            output = result.stdout.strip()

            # JSONとして解析を試行
            try:
                return json.loads(output)
            except json.JSONDecodeError:
                # JSONでない場合は行分割を試行
                if "\n" in output:
                    return [line.strip() for line in output.split("\n") if line.strip()]
                else:
                    return output

        except subprocess.TimeoutExpired:
            logger.error(f"Command timeout: {self.transform_spec}")
            return value
        except Exception as e:
            logger.error(f"Command execution error: {self.transform_spec}, error: {e}")
            return value


def parse_array_transform_rules(
    array_transform_rules: List[str], prefix: str, schema: dict = None
) -> Dict[str, ArrayTransformRule]:
    """
    配列変換ルールのパース。
    形式: "json.path=function:module:func_name" または "json.path=command:cat"
    """
    if not prefix or not isinstance(prefix, str):
        raise ValueError("prefixは空ではない文字列である必要があります。")

    rules = {}
    # prefixの末尾にドットがなければ自動で追加
    normalized_prefix = prefix if prefix.endswith(".") else prefix + "."

    def resolve_schema_path(path_keys, schema):
        # スキーマがなければそのまま
        if not schema:
            return None
        props = schema.get("properties", {})
        current_schema = schema
        resolved_keys = []
        for k in path_keys:
            if re.fullmatch(r"\d+", k):
                resolved_keys.append(k)
                if isinstance(current_schema, dict) and "items" in current_schema:
                    current_schema = current_schema["items"]
                    props = (
                        current_schema.get("properties", {})
                        if isinstance(current_schema, dict)
                        else {}
                    )
                else:
                    props = {}
            else:
                if not props or not isinstance(props, dict):
                    resolved_keys.append(k)
                    break
                # アンダースコアをワイルドカード1文字としてマッチ
                pattern = "^" + re.escape(k).replace("_", ".") + "$"
                matches = [
                    prop
                    for prop in props
                    if re.fullmatch(pattern, prop, flags=re.UNICODE)
                ]
                if len(matches) == 1:
                    resolved_keys.append(matches[0])
                    next_schema = props.get(matches[0], {})
                    if isinstance(next_schema, dict) and "properties" in next_schema:
                        current_schema = next_schema
                        props = next_schema["properties"]
                    elif isinstance(next_schema, dict) and "items" in next_schema:
                        current_schema = next_schema
                        props = next_schema.get("properties", {})
                    else:
                        props = {}
                else:
                    resolved_keys.append(k)
                    break
        return resolved_keys

    for rule in array_transform_rules:
        if "=" not in rule:
            logger.warning(f"無効な変換設定: {rule}")
            continue

        path, transform_spec = rule.split("=", 1)

        # プレフィックスを削除（normalized_prefixで統一）
        if path.startswith(normalized_prefix):
            path = path.removeprefix(normalized_prefix)

        # 変換タイプごとに必ず新しいインスタンスを生成
        if transform_spec.startswith("function:"):
            transform_type = "function"
            transform_spec = transform_spec[9:]
            rule_obj = ArrayTransformRule(path, transform_type, transform_spec)
        elif transform_spec.startswith("command:"):
            transform_type = "command"
            transform_spec = transform_spec[8:]
            rule_obj = ArrayTransformRule(path, transform_type, transform_spec)
        elif transform_spec.startswith("split:"):
            transform_type = "split"
            delimiter_str = transform_spec[6:]
            temp_placeholder = "___ESCAPED_PIPE___"
            delimiter_str = delimiter_str.replace("\\|", temp_placeholder)
            delimiter_parts = delimiter_str.split("|")
            delimiters = []
            for part in delimiter_parts:
                processed_delimiter = part.replace("\\n", "\n")
                processed_delimiter = processed_delimiter.replace("\\t", "\t")
                processed_delimiter = processed_delimiter.replace("\\r", "\r")
                processed_delimiter = processed_delimiter.replace(temp_placeholder, "|")
                delimiters.append(processed_delimiter)

            def split_transform(value, delims=delimiters):
                return convert_string_to_multidimensional_array(value, delims)

            rule_obj = ArrayTransformRule(path, transform_type, transform_spec)
            rule_obj._transform_func = split_transform
        else:
            logger.warning(
                f"不明な変換タイプ: {transform_spec} (function:、command:、split: で始まる必要があります)"
            )
            continue

        # ルール登録（元のパス）
        # ルール登録前に詳細デバッグ
        logger.debug(
            f"Rule register: path={path}, type={transform_type}, id={id(rule_obj)}"
        )
        if path in rules:
            logger.debug(
                f"Rule already exists: path={path}, type={rules[path].transform_type}, id={id(rules[path])}"
            )
        # function型が既に登録されている場合はsplit型で上書きしない
        # split型が既に登録されていてもfunction型が来たらfunction型で上書きする
        if path in rules:
            if transform_type == "split" and rules[path].transform_type == "function":
                logger.debug(
                    f"Skip split rule for {path} (function already registered)"
                )
            elif transform_type == "function" and rules[path].transform_type == "split":
                logger.debug(f"function型でsplit型を上書き: {path}")
                rules[path] = rule_obj
                logger.info(
                    f"Registered transform rule: {path} -> {transform_type}:{transform_spec}"
                )
            elif rules[path].transform_type == transform_type:
                logger.debug(f"同じ型のルールが既に登録済み: {path}")
            else:
                # その他の型の組み合わせは上書き
                rules[path] = rule_obj
                logger.info(
                    f"Registered transform rule: {path} -> {transform_type}:{transform_spec}"
                )
        else:
            rules[path] = rule_obj
            logger.info(
                f"Registered transform rule: {path} -> {transform_type}:{transform_spec}"
            )

        # スキーマで解決したパスも登録
        if schema:
            path_keys = path.split(".")
            resolved_keys = resolve_schema_path(path_keys, schema)
            resolved_path = ".".join(resolved_keys) if resolved_keys else None
            if resolved_path and resolved_path != path:
                # スキーマ解決後のパスにも新しいインスタンスを登録
                new_rule_obj = ArrayTransformRule(
                    resolved_path, transform_type, transform_spec
                )
                logger.debug(
                    f"Rule register: resolved_path={resolved_path}, type={transform_type}, id={id(new_rule_obj)}"
                )
                if resolved_path in rules:
                    logger.debug(
                        f"Rule already exists: resolved_path={resolved_path}, type={rules[resolved_path].transform_type}, id={id(rules[resolved_path])}"
                    )
                if resolved_path in rules:
                    if (
                        transform_type == "split"
                        and rules[resolved_path].transform_type == "function"
                    ):
                        logger.debug(
                            f"Skip split rule for {resolved_path} (function already registered)"
                        )
                    elif (
                        transform_type == "function"
                        and rules[resolved_path].transform_type == "split"
                    ):
                        logger.debug(f"function型でsplit型を上書き: {resolved_path}")
                        rules[resolved_path] = new_rule_obj
                        logger.info(
                            f"Registered transform rule (schema-resolved): {resolved_path} -> {transform_type}:{transform_spec}"
                        )
                    elif rules[resolved_path].transform_type == transform_type:
                        logger.debug(f"同じ型のルールが既に登録済み: {resolved_path}")
                    else:
                        # その他の型の組み合わせは上書き
                        rules[resolved_path] = new_rule_obj
                        logger.info(
                            f"Registered transform rule (schema-resolved): {resolved_path} -> {transform_type}:{transform_spec}"
                        )
                else:
                    rules[resolved_path] = new_rule_obj
                    logger.info(
                        f"Registered transform rule (schema-resolved): {resolved_path} -> {transform_type}:{transform_spec}"
                    )

    return rules


# =============================================================================
# Array Processing
# =============================================================================


def should_convert_to_array(
    path_keys: List[str], split_rules: Dict[str, List[str]]
) -> Optional[List[str]]:
    """
    指定されたパスが配列化対象かどうかを判定し、対応する区切り文字のリストを返す。
    """
    path_str = ".".join(path_keys)

    # 完全一致
    if path_str in split_rules:
        return split_rules[path_str]

    # 部分マッチング（前方一致）
    for rule_path, delimiters in split_rules.items():
        if path_str.startswith(rule_path + ".") or path_str == rule_path:
            return delimiters

    return None


def should_transform_to_array(
    path_keys: List[str], transform_rules: Dict[str, ArrayTransformRule]
) -> Optional[ArrayTransformRule]:
    """
    指定されたパスが配列変換対象かどうかを判定し、対応する変換ルールを返す。
    """
    path_str = ".".join(path_keys)

    # 完全一致のみ
    return transform_rules.get(path_str, None)


def is_string_array_schema(schema: Dict[str, Any]) -> bool:
    """
    スキーマが文字列配列かどうかを判定する。
    """
    if not isinstance(schema, dict):
        return False

    # type: "array" かつ items.type: "string" の場合
    if schema.get("type") == "array":
        items = schema.get("items", {})
        if isinstance(items, dict) and items.get("type") == "string":
            return True

    return False


def check_schema_for_array_conversion(
    path_keys: List[str], schema: Optional[Dict[str, Any]]
) -> bool:
    """
    スキーマを参照して、指定されたパスが文字列配列として定義されているかを判定する。
    """
    if not schema:
        return False

    current_schema = schema
    for key in path_keys:
        if re.fullmatch(r"\d+", key):
            # 数字キーの場合は items を参照
            if isinstance(current_schema, dict) and "items" in current_schema:
                current_schema = current_schema["items"]
            else:
                return False
        else:
            # 文字列キーの場合は properties を参照
            if isinstance(current_schema, dict) and "properties" in current_schema:
                props = current_schema["properties"]
                if key in props:
                    current_schema = props[key]
                else:
                    return False
            else:
                return False

    return is_string_array_schema(current_schema)


def convert_string_to_multidimensional_array(value: Any, delimiters: List[str]) -> Any:
    """
    文字列を指定された区切り文字のリストで多次元配列に変換する。

    Args:
        value: 変換対象の値
        delimiters: 区切り文字のリスト（1次元目、2次元目、3次元目...の順）

    Returns:
        多次元配列に変換された値
    """
    if not isinstance(value, str):
        return value

    if not value.strip():
        return []

    if not delimiters:
        return value

    # 再帰的に次元を分割
    def split_recursively(text: str, delimiter_list: List[str]) -> Any:
        if not delimiter_list:
            return text.strip()

        current_delimiter = delimiter_list[0]
        remaining_delimiters = delimiter_list[1:]

        # 現在の区切り文字で分割
        parts = text.split(current_delimiter)

        if not remaining_delimiters:
            # 最後の次元：文字列のリストを返す
            result = [part.strip() for part in parts if part.strip()]
            return result if result else []
        else:
            # 中間次元：再帰的に処理
            result = []
            for part in parts:
                part = part.strip()
                if part:
                    sub_result = split_recursively(part, remaining_delimiters)
                    result.append(sub_result)
            return result if result else []

    return split_recursively(value, delimiters)


def convert_string_to_array(value: Any, delimiter: str) -> Any:
    """
    文字列を指定された区切り文字で配列に変換する。
    （後方互換性のため残存）
    """
    if not isinstance(value, str):
        return value

    if not value.strip():
        return []

    # 区切り文字で分割
    parts = value.split(delimiter)
    # 前後の空白を削除
    result = [part.strip() for part in parts if part.strip()]

    return result if result else []


# =============================================================================
# Named Range Parsing
# =============================================================================


def parse_named_ranges_with_prefix(
    xlsx_path: Path,
    prefix: str,
    array_split_rules: Optional[Dict[str, List[str]]] = None,
    array_transform_rules: Optional[Dict[str, ArrayTransformRule]] = None,
    containers: Optional[Dict[str, Dict]] = None,
) -> Dict[str, Any]:
    """
    Excel 名前付き範囲(prefix) を解析してネスト dict/list を返す。
    prefixはデフォルトで"json"。
    array_split_rules: 配列化設定の辞書 {path: [delimiter1, delimiter2, ...]}
    array_transform_rules: 配列変換設定の辞書 {path: ArrayTransformRule}
    """
    # 文字列パスをPathオブジェクトに変換（互換性のため）
    if isinstance(xlsx_path, str):
        xlsx_path = Path(xlsx_path)

    if not xlsx_path or not isinstance(xlsx_path, Path):
        raise ValueError(
            "xlsx_pathは有効なPathオブジェクトまたは文字列パスである必要があります。"
        )

    if not xlsx_path.exists():
        raise FileNotFoundError(f"Excelファイルが見つかりません: {xlsx_path}")

    if not xlsx_path.is_file():
        raise ValueError(f"指定されたパスはファイルではありません: {xlsx_path}")

    if not prefix or not isinstance(prefix, str):
        raise ValueError("prefixは空ではない文字列である必要があります。")

    try:
        wb = load_workbook(xlsx_path, data_only=True)
    except Exception as e:
        raise ValueError(f"Excelファイルの読み込みに失敗しました: {xlsx_path} - {e}")

    # コンテナ処理：自動セル名生成
    if containers:
        logger.debug(f"コンテナ処理開始: {len(containers)}個のコンテナ")
        generated_names = generate_cell_names_from_containers(containers, wb)
        
        # 生成されたセル名を名前付き範囲として動的に追加
        for name, range_ref in generated_names.items():
            # プレフィックス付きのセル名を作成
            prefixed_name = f"{prefix}.{name}"
            if prefixed_name not in wb.defined_names:
                logger.debug(f"セル名追加: {prefixed_name} -> {range_ref}")
                # openpyxlでの動的セル名追加は複雑なため、
                # 内部辞書で直接管理してparse処理で参照
                wb._generated_names = getattr(wb, '_generated_names', {})
                wb._generated_names[prefixed_name] = range_ref
                logger.debug(f"セル名追加成功（内部管理）: {prefixed_name}")
            else:
                logger.debug(f"セル名は既存: {prefixed_name}")
        
        logger.info(f"コンテナ処理完了: {len(generated_names)}個のセル名を生成")

    result: Dict[str, Any] = {}

    # schemaファイルがあれば読み込む
    global _global_schema
    schema = None
    try:
        schema = _global_schema
    except Exception:
        schema = None

    if array_split_rules is None:
        array_split_rules = {}
    if array_transform_rules is None:
        array_transform_rules = {}

    def match_schema_key(key: str, schema_props: dict) -> str:
        if not schema_props:
            return key
        key = key.strip()
        pattern = "^" + re.escape(key).replace("_", ".") + "$"
        matches = [
            prop
            for prop in schema_props
            if re.fullmatch(pattern, prop, flags=re.UNICODE)
        ]
        logger.debug(f"key={key}, pattern={pattern}, matches={matches}")
        if len(matches) == 1:
            return matches[0]
        elif len(matches) > 1:
            logger.warning(
                f"ワイルドカード照合で複数マッチ: '{key}' → {matches}。ユニークでないため置換しません。"
            )
        return key

    # prefixの末尾にドットがなければ自動で追加
    normalized_prefix = prefix if prefix.endswith(".") else prefix + "."

    # 既存の名前付き範囲と生成されたセル名を統合
    all_names = dict(wb.defined_names.items())
    if hasattr(wb, '_generated_names'):
        logger.debug(f"生成されたセル名を処理対象に追加: {len(wb._generated_names)}個")
        for gen_name, gen_range in wb._generated_names.items():
            if gen_name not in all_names:
                # 簡易的なDefinedNameオブジェクト作成
                class GeneratedDefinedName:
                    def __init__(self, attr_text):
                        self.attr_text = attr_text
                        # destinationsを模擬（sheet_name, 範囲文字列のタプル）
                        if '!' in attr_text:
                            sheet_part, range_part = attr_text.split('!')
                            self.destinations = [(sheet_part, range_part)]
                        else:
                            # デフォルトでSheet1とする
                            self.destinations = [('Sheet1', attr_text)]
                        
                all_names[gen_name] = GeneratedDefinedName(gen_range)
                logger.debug(f"生成セル名追加: {gen_name} -> {gen_range}")

    for name, defined_name in all_names.items():
        if not name.startswith(normalized_prefix):
            continue
        # セル範囲名からjson.を除去し、'.'で分割して空文字列を除去
        original_path_keys = [
            k for k in name.removeprefix(normalized_prefix).split(".") if k
        ]
        path_keys = original_path_keys.copy()

        # スキーマ解決
        schema_path_keys = []
        schema_broken = False
        if schema is not None:
            props = schema.get("properties", {})
            items = schema.get("items", {})
            current_schema = schema
            for k in path_keys:
                if re.fullmatch(r"\d+", k):
                    schema_path_keys.append(k)
                    if isinstance(current_schema, dict) and "items" in current_schema:
                        current_schema = current_schema["items"]
                        props = (
                            current_schema.get("properties", {})
                            if isinstance(current_schema, dict)
                            else {}
                        )
                        items = (
                            current_schema.get("items", {})
                            if isinstance(current_schema, dict)
                            else {}
                        )
                    else:
                        props = {}
                        items = {}
                else:
                    if not props or not isinstance(props, dict):
                        logger.debug(f"props is empty or not dict at key={k}, break")
                        schema_broken = True
                        break
                    logger.debug(f"props.keys() at key={k}: {list(props.keys())}")
                    new_k = match_schema_key(k, props)
                    schema_path_keys.append(new_k)
                    next_schema = (
                        props.get(new_k, {}) if isinstance(props, dict) else {}
                    )
                    if isinstance(next_schema, dict) and "properties" in next_schema:
                        current_schema = next_schema
                        props = next_schema["properties"]
                        items = next_schema.get("items", {})
                    elif isinstance(next_schema, dict) and "items" in next_schema:
                        current_schema = next_schema
                        props = next_schema.get("properties", {})
                        items = next_schema["items"]
                    else:
                        props = {}
                        items = {}
            # スキーマで途中までしか解決できなかった場合は original_path_keys を使う
            if not schema_broken:
                path_keys = schema_path_keys

        # 値を取得
        # コンテナ生成されたセル名の場合は直接値を取得
        if hasattr(wb, '_generated_names') and name in wb._generated_names:
            value = wb._generated_names[name]
            logger.debug(f"コンテナ生成セル名の値を直接取得: {name} -> {value}")
        else:
            value = get_named_range_values(wb, defined_name)

        # 配列化処理の優先順位:
        # 1. 配列変換ルール（ArrayTransformRule）
        # 2. 明示的なルールによる配列化（多次元対応）
        # 3. スキーマによる配列化（デフォルト区切り文字: カンマ）

        # 1. 配列変換ルールをチェック（original_path_keys と schema-resolved path 両方）
        # 変換ルールの判定は path_keys/original_path_keys 両方で探す
        def get_transform_rule(array_transform_rules, path_keys, original_path_keys):
            key_path = ".".join(path_keys)
            orig_key_path = ".".join(original_path_keys)
            # 完全一致優先
            if key_path in array_transform_rules:
                return array_transform_rules[key_path]
            if orig_key_path in array_transform_rules:
                return array_transform_rules[orig_key_path]
            # 配列要素の場合は親キーでも判定
            if len(path_keys) > 1 and ".".join(path_keys[:-1]) in array_transform_rules:
                return array_transform_rules[".".join(path_keys[:-1])]
            if len(original_path_keys) > 1 and ".".join(original_path_keys[:-1]) in array_transform_rules:
                return array_transform_rules[".".join(original_path_keys[:-1])]

            # ワイルドカード（*）対応: ルール側に*が含まれる場合はパターンマッチ
            for rule_key, rule in array_transform_rules.items():
                if "*" in rule_key:
                    # *を正規表現の「任意の非ドット文字列」に変換
                    pattern = "^" + re.escape(rule_key).replace("\\*", "[^.]+") + "$"
                    if re.match(pattern, key_path):
                        return rule
                    if re.match(pattern, orig_key_path):
                        return rule
            return None

        transform_rule = get_transform_rule(
            array_transform_rules, path_keys, original_path_keys
        )
        logger.debug(
            f"original_path_keys={original_path_keys}, path_keys={path_keys}, transform_rule={transform_rule is not None}, value={value}"
        )

        if transform_rule is not None:
            insert_keys = (
                path_keys
                if ".".join(path_keys) in array_transform_rules
                else original_path_keys
            )
            logger.debug(
                f"変換ルールで変換: {insert_keys} -> rule={transform_rule.transform_type}:{transform_rule.transform_spec}"
            )
            if isinstance(value, list):
                value = [transform_rule.transform(v) for v in value]
            else:
                value = transform_rule.transform(value)

            # function型の場合は追加のsplit/配列化処理はスキップ
            if transform_rule.transform_type == "function":
                logger.debug(
                    f"function型変換後の値: {value} (追加配列化処理はスキップ)"
                )
            # どちらのキーで挿入するか判定
            insert_keys = (
                path_keys
                if ".".join(path_keys) in array_transform_rules
                else original_path_keys
            )
            insert_json_path(result, insert_keys, value, ".".join(insert_keys))
            continue

        logger.debug(f"配列化後の値: {value}")
        
        # シンプルなJSONパス挿入
        insert_json_path(result, path_keys, value, ".".join(path_keys))

    return result


# =============================================================================
# File Operations
# =============================================================================


def collect_xlsx_files(paths: List[str]) -> List[Path]:
    """
    ファイルまたはディレクトリのリストから、対象となる .xlsx ファイル一覧を取得。
    ディレクトリ指定時は直下のみ。
    """
    if not paths:
        raise ValueError("入力パスのリストが空です。少なくとも1つのパスが必要です。")

    files: List[Path] = []
    for p in paths:
        if not p or not isinstance(p, str):
            logger.warning(f"無効なパス形式をスキップします: {p}")
            continue

        p_path = Path(p)
        try:
            if p_path.is_dir():
                for entry in p_path.iterdir():
                    if entry.suffix.lower() == ".xlsx":
                        files.append(entry)
            elif p_path.is_file() and p_path.suffix.lower() == ".xlsx":
                files.append(p_path)
            else:
                logger.warning(f"未処理のパス: {p_path}")
        except (OSError, PermissionError) as e:
            logger.warning(f"パスへのアクセスに失敗しました: {p_path} - {e}")
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

    # datetime型を文字列に変換する関数
    def json_default(obj):
        import datetime

        if isinstance(obj, datetime.datetime):
            return obj.isoformat()
        if isinstance(obj, datetime.date):
            return obj.isoformat()
        return str(obj)

    # ファイル書き出し
    output_dir.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2, default=json_default)

    logger.info(f"ファイルの出力に成功しました: {output_path}")


# =============================================================================
# Container Functions  
# =============================================================================


def parse_container_args(container_args, config_containers=None):
    """CLIのcontainerオプションを解析し、設定ファイルとマージ"""
    combined_containers = config_containers.copy() if config_containers else {}
    
    if not container_args:
        return combined_containers
    
    for container_arg in container_args:
        container_def = json.loads(container_arg)
        combined_containers.update(container_def)
    return combined_containers


def validate_cli_containers(container_args):
    """CLIの--containerオプション専用検証"""
    for i, container_arg in enumerate(container_args):
        try:
            container_def = json.loads(container_arg)
        except json.JSONDecodeError:
            raise ValueError(f"無効なJSON形式: 引数{i+1}")
        
        for cell_name in container_def.keys():
            if not cell_name.startswith('json.'):
                raise ValueError(f"セル名は'json.'で始まる必要があります: {cell_name}")


def calculate_hierarchy_depth(cell_name):
    """数値インデックスを除外した階層深度を計算"""
    parts = cell_name.split('.')
    # 空の部分も除外し、数値でない部分のみを階層として扱う
    hierarchy_parts = [part for part in parts if part and not part.isdigit()]
    return len(hierarchy_parts) - 1  # 'json'を除く


def validate_container_config(containers):
    """コンテナ設定の妥当性を検証"""
    errors = []
    
    for container_name, container_def in containers.items():
        # セル名の形式チェック
        if not container_name.startswith('json.'):
            errors.append(f"コンテナ名は'json.'で始まる必要があります: {container_name}")
        
        # 必須項目のチェック
        has_range = "range" in container_def
        has_offset = "offset" in container_def
        
        if not has_range and not has_offset:
            errors.append(f"コンテナ{container_name}には'range'または'offset'が必要です")
        
        if has_range and has_offset:
            errors.append(f"コンテナ{container_name}には'range'と'offset'の両方を指定できません")
        
        # items の検証
        if "items" not in container_def:
            errors.append(f"コンテナ{container_name}には'items'が必要です")
        elif not isinstance(container_def["items"], list):
            errors.append(f"コンテナ{container_name}の'items'は配列である必要があります")
        elif len(container_def["items"]) == 0:
            errors.append(f"コンテナ{container_name}の'items'は空にできません")
        
        # direction の検証
        if "direction" in container_def:
            direction = container_def["direction"]
            if direction not in ["row", "column"]:
                errors.append(f"コンテナ{container_name}の'direction'は'row'または'column'である必要があります: {direction}")
        
        # increment の検証
        if "increment" in container_def:
            increment = container_def["increment"]
            if not isinstance(increment, int) or increment < 1:
                errors.append(f"コンテナ{container_name}の'increment'は1以上の整数である必要があります: {increment}")
        
        # offset の検証（子コンテナの場合）
        if has_offset:
            offset = container_def["offset"]
            if not isinstance(offset, int):
                errors.append(f"コンテナ{container_name}の'offset'は整数である必要があります: {offset}")
        
        # type の検証（明示的指定の場合）
        if "type" in container_def:
            container_type = container_def["type"]
            if container_type not in ["table", "card", "tree"]:
                errors.append(f"コンテナ{container_name}の'type'は'table'、'card'、または'tree'である必要があります: {container_type}")
    
    return errors


def validate_hierarchy_consistency(containers):
    """コンテナの階層構造の整合性を検証"""
    errors = []
    
    # 親子関係の検証
    for container_name, container_def in containers.items():
        if "offset" in container_def:  # 子コンテナ
            parent_name = get_parent_container_name(container_name)
            if parent_name and parent_name not in containers:
                errors.append(f"子コンテナ{container_name}の親コンテナ{parent_name}が見つかりません")
            elif parent_name:
                parent_def = containers[parent_name]
                if "offset" in parent_def:
                    errors.append(f"親コンテナ{parent_name}もoffset指定されています（range指定である必要があります）")
    
    # 循環参照の検証
    def has_circular_reference(container_name, visited=None):
        if visited is None:
            visited = set()
        
        if container_name in visited:
            return True
        
        visited.add(container_name)
        parent_name = get_parent_container_name(container_name)
        
        if parent_name and parent_name in containers:
            return has_circular_reference(parent_name, visited.copy())
        
        return False
    
    for container_name in containers:
        if has_circular_reference(container_name):
            errors.append(f"コンテナ{container_name}で循環参照が検出されました")
    
    return errors


def filter_cells_with_names(range_coords, cell_names):
    """json.*セル名を持つセルのみを抽出"""
    result = {}
    for cell_coord in range_coords:
        cell_name = cell_names.get(cell_coord)
        if cell_name and cell_name.startswith('json.'):
            result[cell_coord] = cell_name
    return result


def generate_cell_names_from_containers(containers, workbook):
    """
    コンテナ定義からセル名を自動生成し、実際のExcelデータから値を読み取る
    CONTAINER_SPEC.md準拠：Table、Card、Tree型に対応した階層構造処理
    """
    generated_names = {}
    
    # コンテナ設定の妥当性を検証
    config_errors = validate_container_config(containers)
    if config_errors:
        for error in config_errors:
            logger.error(f"コンテナ設定エラー: {error}")
        return generated_names
    
    # 階層構造の整合性を検証
    hierarchy_errors = validate_hierarchy_consistency(containers)
    if hierarchy_errors:
        for error in hierarchy_errors:
            logger.error(f"階層構造エラー: {error}")
        return generated_names
    
    # コンテナを階層順にソート（親→子の順で処理）
    sorted_containers = sort_containers_by_hierarchy(containers)
    
    for container_name, container_def in sorted_containers:
        logger.debug(f"コンテナ処理開始: {container_name}")
        processing_stats.containers_processed += 1
        
        # 階層レベルを計算し、ルート/子を判定
        has_range = "range" in container_def
        has_offset = "offset" in container_def
        
        if has_range:
            # ルートコンテナの処理
            process_root_container(container_name, container_def, workbook, generated_names)
        elif has_offset:
            # 子コンテナの処理
            process_child_container(container_name, container_def, workbook, generated_names)
        else:
            logger.warning(f"コンテナ{container_name}にrangeまたはoffsetが指定されていません")
    
    logger.debug(f"生成されたセル名と値: {generated_names}")
    return generated_names


def sort_containers_by_hierarchy(containers):
    """コンテナを階層の深さでソート（浅い順）"""
    container_items = list(containers.items())
    return sorted(container_items, key=lambda x: calculate_hierarchy_depth(x[0]))


def calculate_hierarchy_depth(cell_name):
    """セル名から階層の深さを計算（数値インデックスを除外）"""
    parts = cell_name.split('.')
    hierarchy_parts = [part for part in parts if not part.isdigit()]
    depth = len(hierarchy_parts) - 1  # 'json'を除く
    
    # ルートコンテナの特徴：rangeプロパティを持つ
    # 子コンテナの特徴：offsetプロパティを持つ
    logger.debug(f"階層深度計算: {cell_name} -> 部分={hierarchy_parts} -> 深度={depth}")
    return depth


def process_root_container(container_name, container_def, workbook, generated_names):
    """ルートコンテナ（range指定あり）の処理"""
    logger.debug(f"ルートコンテナ処理: {container_name}")
    
    # 方向とincrement設定
    direction = container_def.get("direction", "row")
    increment = container_def.get("increment", 1)
    items = container_def.get("items", [])
    labels = container_def.get("labels", [])
    
    logger.debug(f"コンテナ設定: direction={direction}, increment={increment}, items={items}, labels={labels}")
    
    # 範囲指定を解決
    range_spec = container_def.get("range")
    if not range_spec:
        logger.warning(f"ルートコンテナ{container_name}にrange指定がありません")
        return
        
    try:
        # 範囲解決
        actual_range = resolve_range_specification(range_spec, workbook)
        if not actual_range:
            logger.warning(f"範囲を解決できませんでした: {range_spec}")
            return
        
        # コンテナタイプの自動判定
        container_type = detect_container_type(actual_range, workbook, increment, container_def)
        logger.debug(f"検出されたコンテナタイプ: {container_type}")
        
        # タイプ別処理
        if container_type == "table":
            process_table_container(container_name, container_def, actual_range, workbook, generated_names)
        elif container_type == "card":
            process_card_container(container_name, container_def, actual_range, workbook, generated_names)
        elif container_type == "tree":
            process_tree_container(container_name, container_def, actual_range, workbook, generated_names)
        else:
            logger.warning(f"未対応のコンテナタイプ: {container_type}")
            
    except Exception as e:
        logger.error(f"ルートコンテナ処理エラー ({container_name}): {e}")


def process_child_container(container_name, container_def, workbook, generated_names):
    """子コンテナ（offset指定）の処理"""
    logger.debug(f"子コンテナ処理: {container_name}")
    
    offset = container_def.get("offset", 0)
    items = container_def.get("items", [])
    labels = container_def.get("labels", [])
    
    # 親コンテナを特定
    parent_container_name = get_parent_container_name(container_name)
    if not parent_container_name:
        logger.error(f"親コンテナが特定できません: {container_name}")
        return
    
    # 親コンテナの各インスタンスに対して子データを処理
    parent_instances = find_parent_instances(parent_container_name, generated_names)
    
    for parent_instance_idx in parent_instances:
        process_child_instance(container_name, container_def, parent_instance_idx, workbook, generated_names)


def resolve_range_specification(range_spec, workbook):
    """範囲指定を実際の座標に解決"""
    if range_spec in workbook.defined_names:
        defined_name = workbook.defined_names[range_spec]
        for sheet_name, coord in defined_name.destinations:
            logger.debug(f"名前付き範囲 {range_spec}: {coord}")
            return coord
    elif ":" in range_spec:
        logger.debug(f"直接範囲指定: {range_spec}")
        return range_spec
    else:
        if range_spec in workbook.defined_names:
            defined_name = workbook.defined_names[range_spec]
            cell_value = get_named_range_values(workbook, defined_name)
            if isinstance(cell_value, str) and ":" in cell_value:
                logger.debug(f"セル名 {range_spec} の値を範囲として使用: {cell_value}")
                return cell_value
    return None


def detect_container_type(range_spec, workbook, increment, container_def=None):
    """コンテナタイプを自動判定（拡張版）"""
    logger.debug(f"コンテナタイプ判定: range_spec={range_spec}, increment={increment}")
    
    # 明示的なタイプ指定がある場合はそれを使用
    if container_def and "type" in container_def:
        explicit_type = container_def["type"].lower()
        if explicit_type in ["table", "card", "tree"]:
            logger.debug(f"明示的タイプ指定: {explicit_type}")
            return explicit_type
    
    # 範囲名による判定
    range_spec_lower = range_spec.lower()
    if "card" in range_spec_lower:
        return "card"
    elif "tree" in range_spec_lower or "ツリー" in range_spec_lower:
        return "tree"
    elif "table" in range_spec_lower or "表" in range_spec_lower:
        return "table"
    elif "list" in range_spec_lower or "リスト" in range_spec_lower:
        return "table"  # リストはテーブル型として扱う
    
    # increment値による判定
    if increment > 1:
        return "card"  # incrementが大きい場合はカード型
    
    # 範囲のサイズと形状による判定
    try:
        if ":" in range_spec:
            start_coord, end_coord = parse_range(range_spec)
            width = end_coord[0] - start_coord[0] + 1
            height = end_coord[1] - start_coord[1] + 1
            
            # 正方形に近い場合はカード型
            aspect_ratio = max(width, height) / min(width, height)
            if aspect_ratio < 2.0 and min(width, height) > 3:
                logger.debug(f"範囲形状によりカード型判定: {width}x{height}, ratio={aspect_ratio:.2f}")
                return "card"
            
            # 縦長の場合はテーブル型、横長の場合もテーブル型
            logger.debug(f"範囲形状によりテーブル型判定: {width}x{height}")
    except Exception as e:
        logger.warning(f"範囲解析エラー: {e}")
    
    # デフォルトはテーブル型
    return "table"


def process_table_container(container_name, container_def, range_spec, workbook, generated_names):
    """テーブル型コンテナの処理（既存のロジック）"""
    direction = container_def.get("direction", "row")
    increment = container_def.get("increment", 1)
    items = container_def.get("items", [])
    
    # インスタンス数の検出
    start_coord, end_coord = parse_range(range_spec)
    direction_map = {"row": "vertical", "column": "horizontal"}
    internal_direction = direction_map.get(direction, direction)
    instance_count = detect_instance_count(start_coord, end_coord, internal_direction)
    logger.debug(f"検出されたインスタンス数: {instance_count}")
    
    # 基準セル名を検索
    base_container_name = container_name.replace("json.", "")
    base_positions = {}
    
    # 複数の方法で基準位置を取得
    for item in items:
        position = None
        
        # 方法1: インスタンス番号付きのセル名（既存の方法）
        base_cell_name = f"json.{base_container_name}.1.{item}"
        position = get_cell_position_from_name(base_cell_name, workbook)
        if position:
            logger.debug(f"基準位置取得（方法1）: {item} -> {position} (セル名: {base_cell_name})")
        else:
            # 方法2: 最初のインスタンスのセル名（配列指定）
            array_cell_name = f"json.{base_container_name}.{item}.1"
            position = get_cell_position_from_name(array_cell_name, workbook)
            if position:
                logger.debug(f"基準位置取得（方法2）: {item} -> {position} (セル名: {array_cell_name})")
            else:
                # 方法3: 直接のセル名（単一セル）
                direct_cell_name = f"json.{base_container_name}.{item}"
                position = get_cell_position_from_name(direct_cell_name, workbook)
                if position:
                    logger.debug(f"基準位置取得（方法3）: {item} -> {position} (セル名: {direct_cell_name})")
                else:
                    logger.warning(f"基準セル名が見つかりません: {item} (試行: {base_cell_name}, {array_cell_name}, {direct_cell_name})")
        
        if position:
            base_positions[item] = position
    
    if not base_positions:
        logger.error(f"基準位置が見つからないため、コンテナ {container_name} をスキップ")
        return
    
    # インスタンス分のセル名と値を生成
    ws = workbook.active
    
    for instance_idx in range(1, instance_count + 1):
        instance_values = {}
        all_empty = True
        
        for item in items:
            if item not in base_positions:
                continue
                
            # 基準位置からincrement分移動した位置を計算
            target_position = calculate_target_position(
                base_positions[item], direction, instance_idx, increment
            )
            
            # 生成するセル名
            cell_name = f"json.{base_container_name}.{instance_idx}.{item}"
            
            # Excelから値を読み取り
            cell_value = read_cell_value(target_position, ws)
            processing_stats.cells_read += 1
            
            logger.debug(f"セル値読み取り {cell_name}: row={target_position[1]}, col={target_position[0]}, value={cell_value}")
            
            if cell_value:
                all_empty = False
            else:
                processing_stats.empty_cells_skipped += 1
                
            instance_values[cell_name] = cell_value
        
        # 全項目が空でない場合のみ生成されたセル名に追加
        if not all_empty:
            generated_names.update(instance_values)
            processing_stats.cells_generated += len(instance_values)
            logger.debug(f"インスタンス{instance_idx}: 有効なデータとして追加")
        else:
            logger.debug(f"インスタンス{instance_idx}: 全項目が空のため除外")


def process_card_container(container_name, container_def, range_spec, workbook, generated_names):
    """カード型コンテナの処理"""
    logger.debug(f"カード型コンテナ処理: {container_name}")
    
    direction = container_def.get("direction", "row")
    increment = container_def.get("increment", 1)
    items = container_def.get("items", [])
    labels = container_def.get("labels", [])
    
    # 基準セル名を検索
    base_container_name = container_name.replace("json.", "")
    base_positions = {}
    
    # 複数の方法で基準位置を取得
    for item in items:
        position = None
        
        # 方法1: インスタンス番号付きのセル名（既存の方法）
        base_cell_name = f"json.{base_container_name}.1.{item}"
        position = get_cell_position_from_name(base_cell_name, workbook)
        if position:
            logger.debug(f"カード基準位置取得（方法1）: {item} -> {position}")
        else:
            # 方法2: 最初のインスタンスのセル名（配列指定）
            array_cell_name = f"json.{base_container_name}.{item}.1"
            position = get_cell_position_from_name(array_cell_name, workbook)
            if position:
                logger.debug(f"カード基準位置取得（方法2）: {item} -> {position}")
            else:
                # 方法3: 直接のセル名（単一セル）
                direct_cell_name = f"json.{base_container_name}.{item}"
                position = get_cell_position_from_name(direct_cell_name, workbook)
                if position:
                    logger.debug(f"カード基準位置取得（方法3）: {item} -> {position}")
                else:
                    logger.warning(f"カード基準セル名が見つかりません: {item}")
        
        if position:
            base_positions[item] = position
    
    if not base_positions:
        logger.error(f"カード型基準位置が見つかりません: {container_name}")
        return
    
    # カード数の検出
    ws = workbook.active
    
    # カード型では基準セル名から既存のカード数を検出
    card_count = detect_card_count_from_existing_names(base_container_name, workbook)
    if card_count == 0:
        # セル名がない場合は範囲から推定
        card_count = detect_card_count(base_positions, direction, increment, labels, ws)
    
    logger.debug(f"検出されたカード数: {card_count}")
    
    # 各カードのデータを生成
    for card_idx in range(1, card_count + 1):
        instance_values = {}
        all_empty = True
        
        for item in items:
            if item not in base_positions:
                continue
                
            # カード型では既存のセル名がある場合はそれを使用
            existing_cell_name = f"json.{base_container_name}.{card_idx}.{item}"
            if existing_cell_name in [name for name in workbook.defined_names.keys()]:
                # 既存のセル名から値を読み取り
                defined_name = workbook.defined_names[existing_cell_name]
                cell_value = get_named_range_values(workbook, defined_name)
                logger.debug(f"既存セル名から値取得: {existing_cell_name} -> {cell_value}")
            else:
                # 計算された位置から値を読み取り
                target_position = calculate_target_position(
                    base_positions[item], direction, card_idx, increment
                )
                cell_value = read_cell_value(target_position, ws)
                logger.debug(f"計算位置から値取得: {existing_cell_name} -> {cell_value}")
            
            if cell_value:
                all_empty = False
                
            instance_values[existing_cell_name] = cell_value
        
        if not all_empty:
            generated_names.update(instance_values)
            logger.debug(f"カード{card_idx}: 有効なデータとして追加")


def detect_card_count_from_existing_names(base_container_name, workbook):
    """既存のセル名からカード数を検出"""
    card_indices = set()
    prefix = f"json.{base_container_name}."
    
    for name in workbook.defined_names.keys():
        if name.startswith(prefix):
            # json.card.1.customer_name -> ['1', 'customer_name']
            parts = name[len(prefix):].split('.')
            if parts and parts[0].isdigit():
                card_indices.add(int(parts[0]))
    
    return max(card_indices) if card_indices else 0


def process_tree_container(container_name, container_def, range_spec, workbook, generated_names):
    """ツリー型コンテナの処理（階層構造対応）"""
    logger.debug(f"ツリー型コンテナ処理: {container_name}")
    
    # ツリー型は基本的にテーブル型と同じだが、階層構造を考慮
    process_table_container(container_name, container_def, range_spec, workbook, generated_names)


def get_cell_position_from_name(cell_name, workbook):
    """セル名から座標位置を取得"""
    if cell_name in workbook.defined_names:
        defined_name = workbook.defined_names[cell_name]
        for sheet_name, coord in defined_name.destinations:
            match = re.match(r'^\$?([A-Z]+)\$?(\d+)$', coord)
            if match:
                col_letter, row_num = match.groups()
                col_num = column_index_from_string(col_letter)
                return (col_num, int(row_num))
    return None


def calculate_target_position(base_position, direction, instance_idx, increment):
    """基準位置からターゲット位置を計算"""
    base_col, base_row = base_position
    
    if direction == "row":
        return (base_col, base_row + (instance_idx - 1) * increment)
    else:  # column
        return (base_col + (instance_idx - 1) * increment, base_row)


def read_cell_value(position, worksheet, normalize_options=None):
    """指定位置からセル値を読み取り"""
    try:
        col, row = position
        cell = worksheet.cell(row=row, column=col)
        cell_value = cell.value if cell.value is not None else ""
        
        # 正規化処理
        cell_value = normalize_cell_value(cell_value, normalize_options)
        
        return cell_value
    except Exception as e:
        logger.warning(f"セル値読み取りエラー: {e}")
        return ""


def auto_convert_data_type(value, auto_convert=True):
    """データ型を自動変換"""
    if not auto_convert or value is None:
        return value
    
    if isinstance(value, str):
        value_stripped = value.strip()
        
        # 空文字列の場合
        if not value_stripped:
            return ""
        
        # 数値の判定と変換
        try:
            # 整数として解釈可能か
            if value_stripped.isdigit() or (value_stripped.startswith('-') and value_stripped[1:].isdigit()):
                return int(value_stripped)
            
            # 浮動小数点数として解釈可能か
            float_value = float(value_stripped)
            # 整数と同じ値なら整数として返す
            if float_value.is_integer():
                return int(float_value)
            return float_value
        except ValueError:
            pass
        
        # 真偽値の判定
        if value_stripped.lower() in ['true', 'yes', 'on', '1', 'はい', '真', 'オン']:
            return True
        elif value_stripped.lower() in ['false', 'no', 'off', '0', 'いいえ', '偽', 'オフ']:
            return False
        
        # 日付の判定（簡易版）
        try:
            import datetime
            # ISO形式の日付
            if re.match(r'^\d{4}-\d{2}-\d{2}$', value_stripped):
                return datetime.datetime.strptime(value_stripped, '%Y-%m-%d').date()
            # ISO形式の日時
            if re.match(r'^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}', value_stripped):
                return datetime.datetime.fromisoformat(value_stripped.replace('Z', '+00:00'))
        except (ValueError, ImportError):
            pass
    
    # そのまま返す
    return value


def normalize_cell_value(cell_value, normalize_options=None):
    """セル値の正規化処理"""
    if normalize_options is None:
        normalize_options = {}
    
    # 型変換
    if normalize_options.get("auto_convert", False):
        cell_value = auto_convert_data_type(cell_value)
    
    # 文字列の場合の正規化
    if isinstance(cell_value, str):
        # 前後の空白除去
        if normalize_options.get("trim", True):
            cell_value = cell_value.strip()
        
        # 改行の正規化
        if normalize_options.get("normalize_newlines", False):
            cell_value = cell_value.replace('\r\n', '\n').replace('\r', '\n')
        
        # 全角・半角の正規化
        if normalize_options.get("normalize_width", False):
            import unicodedata
            cell_value = unicodedata.normalize('NFKC', cell_value)
    
    return cell_value


def detect_card_count(base_positions, direction, increment, labels, worksheet):
    """カード数を検出"""
    # 簡易実装：最初のアイテムの位置からincrement間隔でラベル確認
    max_cards = 10  # 最大検索数
    card_count = 0
    
    if not base_positions:
        return 0
    
    first_item = list(base_positions.keys())[0]
    base_position = base_positions[first_item]
    
    for card_idx in range(1, max_cards + 1):
        target_position = calculate_target_position(base_position, direction, card_idx, increment)
        cell_value = read_cell_value(target_position, worksheet)
        
        if cell_value:  # 値がある場合はカードが存在
            card_count = card_idx
        else:
            break  # 空の場合は終了
    
    return card_count


def get_parent_container_name(container_name):
    """子コンテナから親コンテナ名を取得"""
    parts = container_name.split('.')
    if len(parts) >= 3:
        # 数値インデックスを除去して親を特定
        parent_parts = []
        for part in parts[:-1]:  # 最後の部分を除く
            if not part.isdigit():
                parent_parts.append(part)
        return '.'.join(parent_parts)
    return None


def find_parent_instances(parent_container_name, generated_names):
    """生成済みセル名から親インスタンスのインデックスを取得"""
    parent_instances = set()
    prefix = parent_container_name + "."
    
    for cell_name in generated_names.keys():
        if cell_name.startswith(prefix):
            parts = cell_name.replace(prefix, "").split('.')
            if parts and parts[0].isdigit():
                parent_instances.add(int(parts[0]))
    
    return sorted(parent_instances)


def process_child_instance(container_name, container_def, parent_instance_idx, workbook, generated_names):
    """子コンテナの特定インスタンスを処理"""
    logger.debug(f"子インスタンス処理: {container_name}, 親インスタンス: {parent_instance_idx}")
    
    offset = container_def.get("offset", 0)
    items = container_def.get("items", [])
    direction = container_def.get("direction", "row")
    increment = container_def.get("increment", 1)
    
    # 親コンテナを特定
    parent_container_name = get_parent_container_name(container_name)
    if not parent_container_name:
        logger.error(f"親コンテナが特定できません: {container_name}")
        return
    
    # 親の基準位置を取得
    parent_base_name = parent_container_name.replace("json.", "")
    child_base_name = container_name.replace("json.", "")
    
    # 親の最初のアイテムから基準位置を取得
    parent_base_positions = {}
    for item in items:
        # 親コンテナの対応するアイテムから基準位置を取得
        parent_cell_name = f"json.{parent_base_name}.{parent_instance_idx}.{item}"
        if parent_cell_name in generated_names:
            # 親の生成済みセル名から位置を推定
            position = estimate_position_from_parent(parent_cell_name, workbook, generated_names)
            if position:
                parent_base_positions[item] = position
        else:
            # 既存のセル名から取得
            position = get_cell_position_from_name(parent_cell_name, workbook)
            if position:
                parent_base_positions[item] = position
    
    if not parent_base_positions:
        logger.warning(f"親コンテナの基準位置が見つかりません: {container_name}")
        return
    
    # 子コンテナのインスタンス数を検出
    child_instances = detect_child_instances_from_data(
        container_name, parent_instance_idx, parent_base_positions, 
        direction, increment, offset, workbook
    )
    
    logger.debug(f"検出された子インスタンス数: {len(child_instances)} (親{parent_instance_idx})")
    
    # 各子インスタンスのデータを生成
    ws = workbook.active
    
    for child_idx in child_instances:
        instance_values = {}
        all_empty = True
        
        for item in items:
            if item not in parent_base_positions:
                continue
            
            # 子の位置を計算（親の位置 + offset + 子のincrement）
            child_position = calculate_child_position(
                parent_base_positions[item], direction, child_idx, increment, offset
            )
            
            # 子のセル名を生成
            cell_name = f"json.{child_base_name}.{child_idx}.{item}"
            
            # セル値を読み取り
            cell_value = read_cell_value(child_position, ws)
            
            logger.debug(f"子セル値読み取り {cell_name}: row={child_position[1]}, col={child_position[0]}, value={cell_value}")
            
            if cell_value:
                all_empty = False
            
            instance_values[cell_name] = cell_value
        
        # 有効なデータがある場合のみ追加
        if not all_empty:
            generated_names.update(instance_values)
            logger.debug(f"子インスタンス{child_idx}: 有効なデータとして追加")


def detect_child_instances_from_data(container_name, parent_instance_idx, parent_base_positions, 
                                   direction, increment, offset, workbook):
    """実際のデータから子インスタンス数を検出"""
    if not parent_base_positions:
        return []
    
    # 最初のアイテムの位置から子の範囲をスキャン
    first_item = list(parent_base_positions.keys())[0]
    parent_position = parent_base_positions[first_item]
    
    ws = workbook.active
    child_instances = []
    max_children = 10  # 最大子数
    
    for child_idx in range(1, max_children + 1):
        child_position = calculate_child_position(
            parent_position, direction, child_idx, increment, offset
        )
        
        # 子の位置にデータがあるかチェック
        cell_value = read_cell_value(child_position, ws)
        
        if cell_value:
            child_instances.append(child_idx)
        else:
            # 連続する空セルが見つかったら終了
            break
    
    return child_instances


def estimate_position_from_parent(parent_cell_name, workbook, generated_names):
    """親の生成済みセル名から位置を推定"""
    # 簡易実装：既存のセル名から位置を取得
    return get_cell_position_from_name(parent_cell_name, workbook)


def calculate_child_position(parent_position, direction, child_idx, increment, offset):
    """親の位置から子の位置を計算"""
    parent_col, parent_row = parent_position
    
    if direction == "row":
        # 行方向：親の位置 + offset + (child_idx - 1) * increment
        return (parent_col, parent_row + offset + (child_idx - 1) * increment)
    else:  # column
        # 列方向：親の位置 + offset + (child_idx - 1) * increment  
        return (parent_col + offset + (child_idx - 1) * increment, parent_row)


# =============================================================================
# Main Function and CLI
# =============================================================================


def main():
    """メインエントリーポイント"""
    parser = argparse.ArgumentParser(description="Excel の名前付き範囲を JSON に変換")
    parser.add_argument("input_files", nargs="*", help="入力 Excel ファイル")
    parser.add_argument("--config", type=Path, help="設定ファイル")
    parser.add_argument("--output_dir", "-o", type=Path, help="出力ディレクトリ")
    parser.add_argument("--prefix", "-p", default="json", help="プレフィックス")
    parser.add_argument("--schema", "-s", type=Path, help="JSON スキーマファイル")
    parser.add_argument("--log_level", choices=["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"], help="ログレベル")
    parser.add_argument("--keep_empty", action="store_true", help="空の値を保持")
    parser.add_argument("--trim", action="store_true", help="文字列の前後の空白を削除")
    parser.add_argument("--container", action="append", help="コンテナ定義 (JSON)")
    
    args = parser.parse_args()
    
    # 設定ファイルの読み込み
    config = {}
    if args.config:
        with args.config.open("r", encoding="utf-8") as f:
            config = json.load(f)

    # コマンドライン引数で設定を上書き
    if args.input_files:
        config["inputs"] = args.input_files
    if args.output_dir:
        config["output_dir"] = str(args.output_dir)
    if args.prefix:
        config["prefix"] = args.prefix
    if args.schema:
        config["schema"] = str(args.schema)
    if args.keep_empty:
        config["keep_empty"] = True
    if args.log_level:
        config["log_level"] = args.log_level
    if args.container:
        validate_cli_containers(args.container)
        cli_containers = parse_container_args(args.container)
        config_containers = config.get("containers", {})
        config["containers"] = {**config_containers, **cli_containers}

    # ログレベル設定（config優先、なければINFO）
    log_level = config.get("log_level", "INFO")
    logging.basicConfig(
        level=getattr(logging, log_level),
        format="%(levelname)s: %(message)s"
    )
    
    global _global_trim, _global_schema
    _global_trim = args.trim

    try:
        # 必須パラメータのチェック
        if not config.get("inputs"):
            logger.error("入力ファイルが指定されていません")
            return 1

        # スキーマの読み込み
        schema_path = Path(config["schema"]) if config.get("schema") else None
        schema = load_schema(schema_path)
        _global_schema = schema

        # バリデーター作成
        validator = Draft7Validator(schema) if schema else None

        # ファイル処理
        input_files = collect_xlsx_files(config["inputs"])
        output_dir = Path(config.get("output_dir", "output"))
        prefix = config.get("prefix", "json")
        keep_empty = not config.get("keep_empty", True)
        containers = config.get("containers", {})

        # 変換ルール（transform）をparse_array_transform_rulesでパース
        array_transform_rules = None
        if "transform" in config:
            # transformはlistで渡される
            array_transform_rules = parse_array_transform_rules(config["transform"], prefix, schema)

        for xlsx_file in input_files:
            logger.info(f"処理開始: {xlsx_file}")
            processing_stats.reset()
            processing_stats.start_processing()

            try:
                # Excel解析
                result = parse_named_ranges_with_prefix(
                    xlsx_file, prefix, array_transform_rules=array_transform_rules, containers=containers
                )

                # 出力ファイル名
                output_file = output_dir / f"{xlsx_file.stem}.json"

                # JSON書き出し
                write_json(result, output_file, schema, validator, keep_empty)

                processing_stats.end_processing()

                # 統計情報をログ出力（DEBUGレベル以上の場合）
                if logger.isEnabledFor(logging.DEBUG):
                    processing_stats.log_summary()

            except Exception as e:
                processing_stats.add_error(f"ファイル処理エラー {xlsx_file}: {e}")
                processing_stats.end_processing()
                continue

        logger.info("処理完了")
        return 0

    except Exception as e:
        processing_stats.add_error(f"実行エラー: {e}")

        # 詳細なエラー情報をDEBUGレベルで出力
        if logger.isEnabledFor(logging.DEBUG):
            import traceback
            logger.debug(f"エラー詳細:\n{traceback.format_exc()}")

        return 1


if __name__ == "__main__":
    exit(main())  
