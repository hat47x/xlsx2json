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
import sys
import importlib.util
import time
import yaml
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Union, Callable, Tuple
from contextlib import contextmanager

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
from openpyxl.styles import Border
from jsonschema import Draft7Validator

logger = logging.getLogger("xlsx2json")


class Xlsx2JsonError(Exception):
    """xlsx2json関連のベース例外クラス"""

    pass


class ConfigurationError(Xlsx2JsonError):
    """設定関連のエラー"""

    pass


class FileProcessingError(Xlsx2JsonError):
    """ファイル処理関連のエラー"""

    pass


class Xlsx2JsonValidationError(Xlsx2JsonError):
    """バリデーション関連のエラー"""

    pass


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


@dataclass
class ProcessingConfig:
    """処理設定を管理するデータクラス"""

    input_files: List[Union[str, Path]] = field(default_factory=list)
    prefix: str = "json"
    trim: bool = False
    keep_empty: bool = False
    output_dir: Optional[Path] = None
    output_format: str = "json"
    schema: Optional[Dict[str, Any]] = None
    containers: Dict[str, Any] = field(default_factory=dict)
    transform_rules: List[str] = field(default_factory=list)

    def __post_init__(self):
        if self.output_dir and isinstance(self.output_dir, str):
            self.output_dir = Path(self.output_dir)


class Xlsx2JsonConverter:
    """Excel から JSON への変換を行うメインクラス"""

    def __init__(self, config: ProcessingConfig):
        self.config = config
        self.processing_stats = ProcessingStats()
        self.validator = None
        if config.schema:
            self.validator = Draft7Validator(config.schema)

    def process_files(self, input_files: List[Union[str, Path]]) -> int:
        """ファイルリストを処理する"""
        self.processing_stats.start_processing()

        try:
            xlsx_files = self._collect_xlsx_files(input_files)
            for xlsx_file in xlsx_files:
                try:
                    self._process_single_file(xlsx_file)
                except Exception as e:
                    # 個別ファイルのエラーはログに記録するが処理は継続
                    self.processing_stats.add_error(
                        f"ファイル処理エラー {xlsx_file}: {e}"
                    )
                    logger.error(f"ファイル処理を継続します: {xlsx_file}")
        except Exception as e:
            self.processing_stats.add_error(f"処理中にエラーが発生: {e}")
            return 1
        finally:
            self.processing_stats.end_processing()
            self.processing_stats.log_summary()

        # エラーがあっても処理完了の場合は0を返す（従来の動作を維持）
        return 0

    def _collect_xlsx_files(self, inputs: List[Union[str, Path]]) -> List[Path]:
        """入力からXLSXファイルを収集"""
        files = []
        for input_item in inputs:
            path = Path(input_item)
            if path.is_file() and path.suffix.lower() == ".xlsx":
                files.append(path)
            elif path.is_dir():
                files.extend(path.glob("*.xlsx"))
        return files

    def _process_single_file(self, xlsx_file: Path) -> None:
        """単一ファイルの処理"""
        try:
            logger.info(f"Processing: {xlsx_file}")

            # 変換ルールの処理
            array_transform_rules = None
            if self.config.transform_rules:
                array_transform_rules = parse_array_transform_rules(
                    self.config.transform_rules,
                    self.config.prefix,
                    self.config.schema,
                    self.config.trim,
                )

            # 既存の処理ロジックを呼び出し
            result = parse_named_ranges_with_prefix(
                xlsx_file,
                self.config.prefix,
                array_transform_rules=array_transform_rules,
                containers=self.config.containers,
                schema=self.config.schema,
            )

            # 出力処理
            default_output = xlsx_file.parent / "output"
            output_dir = self.config.output_dir or default_output
            base_name = xlsx_file.stem

            # ファイル書き込み（JSON/YAML対応）
            self._write_output(result, output_dir, base_name)

            logger.info(f"処理完了: {xlsx_file}")
            self.processing_stats.containers_processed += 1

        except Exception as e:
            self.processing_stats.add_error(f"ファイル処理エラー {xlsx_file}: {e}")
            # 例外を再発生しない（処理を継続するため）

    def _write_output(self, data: dict, output_dir: Path, base_name: str) -> None:
        """データ出力を書き込み（JSON/YAML対応）"""
        try:
            # 出力フォーマットに応じて拡張子を決定
            if self.config.output_format == "yaml":
                extension = ".yaml"
            else:
                extension = ".json"

            # write_data関数が存在するかチェック
            if "write_data" in globals():
                output_path = output_dir / f"{base_name}{extension}"
                suppress_empty = not self.config.keep_empty
                write_data(
                    data,
                    output_path,
                    self.config.output_format,
                    self.config.schema,
                    self.validator,
                    suppress_empty,
                )
            else:
                # 簡易的な出力（JSONのみ）
                output_file = output_dir / f"{base_name}.json"
                with output_file.open("w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                logger.info(f"JSONファイル出力: {output_file}")
        except Exception as e:
            logger.error(f"JSON出力エラー: {e}")
            raise


# グローバル統計インスタンス
processing_stats = ProcessingStats()


# =============================================================================
# Core Utilities
# =============================================================================


class SchemaLoader:
    """JSONスキーマの読み込みと管理を行うクラス"""

    @staticmethod
    def load_schema(schema_path: Optional[Path]) -> Optional[Dict[str, Any]]:
        """指定されたパスからJSONスキーマを読み込む"""
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
            logger.error(f"無効なJSONフォーマットです: {schema_path} - {e}")
            raise  # 元のJSONDecodeErrorを再発生
        except Exception as e:
            raise IOError(
                f"スキーマファイルの読み込みに失敗: {schema_path} - {e}"
            ) from e

    @staticmethod
    def validate_and_log(
        data: Dict[str, Any], validator: Draft7Validator, log_dir: Path, base_name: str
    ) -> None:
        """JSONデータをバリデートし、エラーがあればファイルに出力"""
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
    範囲の場合は行列構造を保持する。
    """
    all_values: List[Any] = []
    for sheet_name, coord in defined_name.destinations:
        cell_or_range = wb[sheet_name][coord]
        if isinstance(cell_or_range, tuple):  # 範囲
            # 2次元の場合 (複数行複数列)
            if hasattr(cell_or_range[0], "__iter__") and not isinstance(
                cell_or_range[0], str
            ):
                # 行列構造を保持
                for row in cell_or_range:
                    row_values = [cell.value for cell in row]
                    if len(row_values) == 1:
                        # 単一列の場合は値そのものを追加
                        all_values.append(row_values[0])
                    else:
                        # 複数列の場合は行として追加
                        all_values.append(row_values)
            else:
                # 1次元の場合 (単一行または単一列)
                all_values.extend([cell.value for cell in cell_or_range])
        else:
            # 単一セル
            all_values.append(cell_or_range.value)

    # 1セルなら値のみ返す（listでなく）
    if len(all_values) == 1:
        return all_values[0]
    return all_values


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
    cleaned_range = range_str.replace("$", "")

    match = re.match(r"^([A-Z]+)(\d+):([A-Z]+)(\d+)$", cleaned_range)
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


def generate_cell_names(
    container_name: str,
    start_coord: tuple,
    end_coord: tuple,
    direction: str,
    items: list,
) -> list:
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
    config.yaml/config.jsonからコンテナ設定を読み込む
    YAMLフォーマットを優先し、JSONもサポートする（JSONはYAMLのサブセット）
    注意：JSON Schemaファイルは従来通りJSON形式のみサポート
    """
    if not config_path.exists():
        return {}

    try:
        with config_path.open("r", encoding="utf-8") as f:
            # YAMLとして読み込み（JSONはYAMLのサブセットなので自動対応）
            config = yaml.safe_load(f)
            if config is None:
                return {}
            return config.get("containers", {})
    except yaml.YAMLError as e:
        logger.warning(
            f"設定ファイルの読み込みに失敗（YAML解析エラー）: {config_path} - {e}"
        )
        return {}
    except FileNotFoundError:
        logger.warning(f"設定ファイルが見つかりません: {config_path}")
        return {}
    except Exception as e:
        logger.warning(f"設定ファイルの読み込みに失敗: {config_path} - {e}")
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


class DataCleaner:
    """データのクリーニングと検証を行うクラス"""

    @staticmethod
    def is_empty_value(value: Any) -> bool:
        """値が空かどうかを判定する（None, 空白のみの文字列, 空リスト, 空dictはTrue。0やFalseはFalse）"""
        if value is None:
            return True
        if isinstance(value, str):
            return value.strip() == ""
        if isinstance(value, list):
            return len(value) == 0
        if isinstance(value, dict):
            return len(value) == 0
        return False

    @staticmethod
    def is_completely_empty(obj: Any) -> bool:
        """完全に空かどうか（None, 空白のみの文字列, 空リスト/空dict/全要素が完全に空ならTrue。他はFalse）"""
        if obj is None:
            return True
        if isinstance(obj, str):
            return obj.strip() == ""
        if isinstance(obj, list):
            return all(DataCleaner.is_completely_empty(v) for v in obj)
        if isinstance(obj, dict):
            return all(DataCleaner.is_completely_empty(v) for v in obj.values())
        return False

    @staticmethod
    def clean_empty_values(
        obj: Union[Dict[str, Any], List[Any], Any], suppress_empty: bool = True
    ) -> Union[Dict[str, Any], List[Any], Any, None]:
        """空の値を再帰的に除去する"""
        if not suppress_empty:
            return obj

        if isinstance(obj, dict):
            cleaned = {}
            for key, value in obj.items():
                cleaned_value = DataCleaner.clean_empty_values(value, suppress_empty)
                if not DataCleaner.is_empty_value(cleaned_value):
                    cleaned[key] = cleaned_value
            return cleaned if cleaned else None

        elif isinstance(obj, list):
            processed_items = []
            for item in obj:
                processed_item = DataCleaner.clean_empty_values(item, suppress_empty)
                processed_items.append(processed_item)

            cleaned = []
            for item in processed_items:
                if not DataCleaner.is_completely_empty(item):
                    cleaned.append(item)

            return cleaned if cleaned else None

        else:
            return obj if not DataCleaner.is_empty_value(obj) else None

    @staticmethod
    def clean_empty_arrays_contextually(
        obj: Union[Dict[str, Any], List[Any], Any], suppress_empty: bool = True
    ) -> Union[Dict[str, Any], List[Any], Any, None]:
        """配列要素の整合性を保ちながら空値を除去する"""
        if not suppress_empty:
            return obj

        if isinstance(obj, dict):
            cleaned = {}
            for key, value in obj.items():
                cleaned_value = DataCleaner.clean_empty_arrays_contextually(
                    value, suppress_empty
                )
                if cleaned_value is not None:
                    cleaned[key] = cleaned_value
            return cleaned if cleaned else None

        elif isinstance(obj, list):
            processed_items = []
            for item in obj:
                processed_item = DataCleaner.clean_empty_arrays_contextually(
                    item, suppress_empty
                )
                processed_items.append(processed_item)

            cleaned = []
            for item in processed_items:
                if item is not None:
                    cleaned.append(item)

            return cleaned if cleaned else None

        else:
            return obj if not DataCleaner.is_empty_value(obj) else None


# =============================================================================
# JSON Path Operations
# =============================================================================


def parse_json_path(path: str) -> List[str]:
    """
    JSON パス文字列をキーのリストに解析する

    例:
        "data.items[0].value" -> ["data", "items", "0", "value"]
        "users[1].profile.name" -> ["users", "1", "profile", "name"]
    """
    if not path:
        return []

    # 配列アクセス記法 [n] を .n に変換
    path = re.sub(r"\[(\d+)\]", r".\1", path)

    # ドットで分割してキーのリストを作成
    keys = [key for key in path.split(".") if key]

    return keys


def insert_json_path(
    root: Union[Dict[str, Any], List[Any]],
    keys: Union[List[str], str],
    value: Any,
    full_path: str = "",
) -> None:
    """
    ドット区切りキーのリストまたは文字列から JSON 構造を構築し、値を挿入する。
    数字キーは list、文字列キーは dict として扱う。
    配列要素の構築時には適切に辞書から配列への変換も行う。
    """
    # 文字列パスの場合はリストに変換
    if isinstance(keys, str):
        keys = parse_json_path(keys)

    # 空のパスの場合のエラーハンドリング
    if not keys:
        raise ValueError(
            "JSONパスが空です。値を挿入するには少なくとも1つのキーが必要です。"
        )

    key = keys[0]
    is_last = len(keys) == 1
    current_path = f"{full_path}.{key}" if full_path else key

    if re.fullmatch(r"\d+", key):
        idx = int(key) - 1  # 1-basedインデックスを0-basedに変換
        if idx < 0:
            raise ValueError(f"配列インデックスは1以上である必要があります: {key}")

        # rootが辞書の場合は配列に変換する必要がある
        if isinstance(root, dict):
            # 空の辞書の場合は単純に配列に置き換え
            if not root:
                root_ref = []
                # 呼び出し元の参照を更新するため、rootを変更
                root.clear()
                while len(root_ref) <= idx:
                    root_ref.append(None)

                if is_last:
                    root_ref[idx] = value
                else:
                    if root_ref[idx] is None:
                        root_ref[idx] = [] if re.fullmatch(r"\d+", keys[1]) else {}
                    insert_json_path(root_ref[idx], keys[1:], value, current_path)
                return
            else:
                # 辞書に既存データがある場合はエラー
                raise ValueError(
                    f"Cannot convert dict to array at path '{current_path}' - "
                    f"existing dict has non-numeric keys: {list(root.keys())}"
                )

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
        # dict型でない場合は、str型ならdictに置き換えて再帰
        if not isinstance(root, dict):
            if isinstance(root, str):
                # 既存値を__value__に退避
                new_dict = {"__value__": root}
                root = new_dict
            else:
                raise TypeError(
                    f"insert_json_path: root must be dict, got {type(root)}"
                )
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
                next_key_is_numeric = (
                    re.fullmatch(r"\d+", keys[1]) if len(keys) > 1 else False
                )

                if (
                    next_key_is_numeric
                    and isinstance(root[key], dict)
                    and not root[key]
                ):
                    # 空辞書を配列に変換
                    root[key] = []
                    logger.debug(f"空辞書を配列に変換: {current_path}")
                elif (
                    not next_key_is_numeric
                    and isinstance(root[key], list)
                    and not root[key]
                ):
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

    def __init__(
        self,
        path: str,
        transform_type: str,
        transform_spec: str,
        trim_enabled: bool = False,
    ):
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
        self.trim_enabled = trim_enabled
        self._transform_func: Optional[Callable] = None
        self._setup_transform()

    def _setup_transform(self):
        """変換関数をセットアップ"""
        if self.transform_type == "function":
            self._setup_python_function()
        elif self.transform_type == "command":
            self._setup_command()
        elif self.transform_type == "split":
            self._setup_split()
        else:
            raise ValueError(f"Unknown transform type: {self.transform_type}")

    def _setup_split(self):
        """split変換のセットアップ"""
        # transform_specは区切り文字（複数の場合は|で区切り）
        delimiter_str = self.transform_spec

        # パイプ文字のエスケープ処理
        if r"\|" in delimiter_str:
            delimiter_str = delimiter_str.replace(r"\|", "PIPE_ESCAPE_TEMP")

        # 区切り文字を分割
        delimiters = delimiter_str.split("|")

        # エスケープを元に戻す
        delimiters = [d.replace("PIPE_ESCAPE_TEMP", "|") for d in delimiters]

        # 特殊文字の置換
        for i, delimiter in enumerate(delimiters):
            delimiter = delimiter.replace("\\n", "\n")
            delimiter = delimiter.replace("\\t", "\t")
            delimiter = delimiter.replace("\\r", "\r")
            delimiters[i] = delimiter

        # 変換関数を設定
        def split_func(value):
            # 1区切りのみの場合は一次元配列として分割
            if len(delimiters) == 1:
                return convert_string_to_array(value, delimiters[0])
            else:
                return convert_string_to_multidimensional_array(value, delimiters)

        self._transform_func = split_func

    def _setup_python_function(self):
        """Python関数のセットアップ"""
        # 形式: module_path:function_name または file_path:function_name
        if ":" not in self.transform_spec:
            raise ValueError(
                f"Python function spec must be 'module:function' or 'file.py:function': {self.transform_spec}"
            )

        module_or_file, func_name = self.transform_spec.rsplit(":", 1)

        try:
            # 外部ファイルまたはモジュールの処理
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

    def transform(self, value: Any, workbook=None) -> Any:
        """値を変換

        valueがExcelの名前付き範囲の値である場合、そのデータ形式
        （値、1次元配列、2次元配列、さらに高次元配列）に応じて適切に変換関数に渡す

        変換関数が辞書を返した場合、動的セル名構築として処理する
        """
        if self.transform_type == "function":
            result = self._transform_with_function(value)
            # trim指定時は配列要素をstrip()
            if self.trim_enabled and isinstance(result, list):
                return self._apply_trim_recursively(result)
            return result
        elif self.transform_type == "command":
            return self._transform_with_command(value)
        elif self.transform_type == "split":
            return self._apply_split_recursively(value)
        else:
            return value

    def _apply_trim_recursively(self, data: Any) -> Any:
        """多次元配列に対して再帰的にstripを適用"""
        if isinstance(data, list):
            return [self._apply_trim_recursively(item) for item in data]
        elif isinstance(data, str):
            return data.strip()
        else:
            return data

    def _apply_split_recursively(self, value: Any, depth: int = None) -> Any:
        """多次元split変換。リスト要素も個別に処理する。"""
        if isinstance(value, list):
            # リストの各要素を個別に変換
            result = []
            for item in value:
                processed_item = self._apply_split_recursively(item, depth)
                # 各要素を独立して処理（extendではなくappend）
                result.append(processed_item)
            return result
        elif isinstance(value, str):
            # 文字列を変換関数で処理
            return self._transform_func(value)
        else:
            return value

    def _transform_with_command(self, value: Any) -> Any:
        """外部コマンドで変換"""
        try:
            input_str = str(value) if value is not None else ""

            # 標準出力の順序を保つため、stderrのみキャプチャし、stdoutは直接表示
            result = subprocess.run(
                shlex.split(self.transform_spec),
                input=input_str,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                timeout=30,
            )

            if result.returncode != 0:
                # stderrがある場合はログに出力（順序保持のため即座に出力）
                if result.stderr.strip():
                    print(
                        f"Command stderr: {result.stderr.strip()}",
                        file=sys.stderr,
                        flush=True,
                    )
                logger.warning(
                    f"Command failed: {self.transform_spec}, "
                    f"returncode: {result.returncode}"
                )
                return value

            output = result.stdout.strip()

            # コマンドの標準出力があれば即座に表示（順序保持）
            if output and output != str(value):
                print(f"Command output: {output}", flush=True)

            # JSONとして解析を試行
            try:
                return json.loads(output)
            except json.JSONDecodeError:
                # JSONでない場合は行分割を試行
                if "\n" in output:
                    lines = [
                        line.strip() for line in output.split("\n") if line.strip()
                    ]
                    return lines
                else:
                    return output

        except subprocess.TimeoutExpired:
            logger.error(f"Command timeout: {self.transform_spec}")
            return value
        except Exception as e:
            logger.error(f"Command execution error: {self.transform_spec}, error: {e}")
            return value

    def _transform_with_function(self, value: Any) -> Any:
        """関数による変換（標準出力制御付き）"""
        if self._transform_func is None:
            logger.warning(f"Transform function not initialized: {self.transform_spec}")
            return value

        try:
            # 標準出力をキャプチャして順序を制御
            import io
            from contextlib import redirect_stdout, redirect_stderr

            stdout_capture = io.StringIO()
            stderr_capture = io.StringIO()

            with redirect_stdout(stdout_capture), redirect_stderr(stderr_capture):
                result = self._transform_func(value)

            # キャプチャした出力があれば即座に表示（順序保持）
            stdout_content = stdout_capture.getvalue()
            stderr_content = stderr_capture.getvalue()

            if stdout_content.strip():
                print(f"Function output: {stdout_content.strip()}", flush=True)

            if stderr_content.strip():
                print(
                    f"Function stderr: {stderr_content.strip()}",
                    file=sys.stderr,
                    flush=True,
                )

            return result

        except Exception as e:
            logger.error(f"Function execution error: {self.transform_spec}, error: {e}")
            return value


def parse_array_transform_rules(
    array_transform_rules: List[str],
    prefix: str,
    schema: dict = None,
    trim_enabled: bool = False,
) -> Dict[str, List[ArrayTransformRule]]:
    """
    配列変換ルールのパース。
    形式: "json.path=function:module:func_name" または "json.path=command:cat"
    ワイルドカード対応: "json.arr.*.name=split:," または "json.arr.*=range:A1:B2:function:builtins:len"
    連続適用対応: 同一セル名に対する複数の--transform指定を順次適用
    """
    if not prefix or not isinstance(prefix, str):
        raise ValueError("prefixは空ではない文字列である必要があります。")

    rules = {}  # Dict[str, List[ArrayTransformRule]]
    wildcard_rules = {}  # ワイルドカードルールを別途管理

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

        # ワイルドカード（*）を含むかチェック
        has_wildcard = "*" in path

        # プレフィックス処理
        if path.startswith(normalized_prefix):
            path = path[len(normalized_prefix) :]
        elif path.startswith(prefix):
            path = path[len(prefix) :]
            if not path.startswith("."):
                path = "." + path if path else ""
            if path.startswith("."):
                path = path[1:]

        # 変換ルール作成
        try:
            transform_type = "function"
            actual_spec = transform_spec

            if transform_spec.startswith("function:"):
                transform_type = "function"
                actual_spec = transform_spec[9:]  # "function:"を除去
            elif transform_spec.startswith("command:"):
                transform_type = "command"
                actual_spec = transform_spec[8:]  # "command:"を除去
            elif transform_spec.startswith("split:"):
                transform_type = "split"
                actual_spec = transform_spec[6:]  # "split:"を除去

            rule_obj = ArrayTransformRule(
                path, transform_type, actual_spec, trim_enabled
            )

            if has_wildcard:
                # ワイルドカードルールとして保存（リスト形式）
                if path not in wildcard_rules:
                    wildcard_rules[path] = []
                wildcard_rules[path].append(rule_obj)
            else:
                # 通常のルール（リスト形式）
                if path not in rules:
                    rules[path] = []
                rules[path].append(rule_obj)

        except Exception as e:
            logger.error(f"変換ルール作成エラー: {rule}, エラー: {e}")
            continue

    # ワイルドカードルールも統合（後で適用）
    if wildcard_rules:
        rules.update(wildcard_rules)

    return rules


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
    path_keys: List[str], transform_rules: Dict[str, List[ArrayTransformRule]]
) -> Optional[List[ArrayTransformRule]]:
    """
    指定されたパスが配列変換対象かどうかを判定し、対応する変換ルールのリストを返す。
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
    最終次元は個別の文字列要素として展開される。

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

    def split_recursively(text: str, delimiter_list: List[str], depth: int = 0) -> Any:
        if not delimiter_list:
            return text.strip()

        current_delimiter = delimiter_list[0]
        remaining_delimiters = delimiter_list[1:]

        # 現在の区切り文字で分割
        parts = text.split(current_delimiter)

        result = []
        for part in parts:
            part = part.strip()
            if part:
                if remaining_delimiters:
                    # まだ区切り文字が残っている場合は再帰処理
                    sub_result = split_recursively(
                        part, remaining_delimiters, depth + 1
                    )
                    result.append(sub_result)
                else:
                    # 最後の区切り文字の場合：個別の文字列として追加
                    result.append(part)

        # 最終次元の場合、配列を個別要素として展開
        if len(delimiter_list) == 1:
            return result  # 最終次元なので配列のまま返す

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
    schema: Optional[Dict[str, Any]] = None,
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
                wb._generated_names = getattr(wb, "_generated_names", {})
                wb._generated_names[prefixed_name] = range_ref
                logger.debug(f"セル名追加成功（内部管理）: {prefixed_name}")
            else:
                logger.debug(f"セル名は既存: {prefixed_name}")

        logger.info(f"コンテナ処理完了: {len(generated_names)}個のセル名を生成")

    result: Dict[str, Any] = {}

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
    if hasattr(wb, "_generated_names"):
        logger.debug(f"生成されたセル名を処理対象に追加: {len(wb._generated_names)}個")
        for gen_name, gen_range in wb._generated_names.items():
            if gen_name not in all_names:
                # 簡易的なDefinedNameオブジェクト作成
                class GeneratedDefinedName:
                    def __init__(self, attr_text):
                        self.attr_text = attr_text
                        # destinationsを模擬（sheet_name, 範囲文字列のタプル）
                        if "!" in attr_text:
                            sheet_part, range_part = attr_text.split("!")
                            self.destinations = [(sheet_part, range_part)]
                        else:
                            # デフォルトでSheet1とする
                            self.destinations = [("Sheet1", attr_text)]

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
        if hasattr(wb, "_generated_names") and name in wb._generated_names:
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
        def get_transform_rules(array_transform_rules, path_keys, original_path_keys):
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
            if (
                len(original_path_keys) > 1
                and ".".join(original_path_keys[:-1]) in array_transform_rules
            ):
                return array_transform_rules[".".join(original_path_keys[:-1])]

            # ワイルドカード（*）対応: ルール側に*が含まれる場合はパターンマッチ
            for rule_key, rule_list in array_transform_rules.items():
                if "*" in rule_key:
                    # *を正規表現の「任意の非ドット文字列」に変換
                    pattern = "^" + re.escape(rule_key).replace("\\*", "[^.]+") + "$"
                    if re.match(pattern, key_path):
                        return rule_list
                    if re.match(pattern, orig_key_path):
                        return rule_list
            return None

        transform_rules = get_transform_rules(
            array_transform_rules, path_keys, original_path_keys
        )
        logger.debug(
            f"original_path_keys={original_path_keys}, path_keys={path_keys}, transform_rules={transform_rules is not None and len(transform_rules)}, value={value}"
        )

        if transform_rules is not None:
            insert_keys = (
                path_keys
                if ".".join(path_keys) in array_transform_rules
                else original_path_keys
            )

            # 連続適用: 複数の変換ルールを順次適用
            current_value = value
            for i, transform_rule in enumerate(transform_rules):
                logger.debug(
                    f"変換ルール{i+1}/{len(transform_rules)}で変換: {insert_keys} -> rule={transform_rule.transform_type}:{transform_rule.transform_spec}"
                )
                if isinstance(current_value, list):
                    current_value = [
                        transform_rule.transform(v, wb) for v in current_value
                    ]
                else:
                    current_value = transform_rule.transform(current_value, wb)

                # 辞書戻り値の場合は動的セル名構築を処理
                if isinstance(current_value, dict):
                    logger.debug(f"辞書戻り値による動的セル名構築: {current_value}")
                    for key, val in current_value.items():
                        if key.startswith(prefix):
                            # 絶対指定
                            abs_path = (
                                key[len(prefix) :]
                                if key.startswith(prefix + ".")
                                else key[len(prefix) :]
                            )
                            if abs_path:
                                abs_parts = abs_path.split(".")
                                # 動的に生成されたセル名に対しても変換ルールを再帰的に適用
                                dynamic_rules = get_transform_rules(
                                    array_transform_rules, abs_parts, abs_parts
                                )
                                if dynamic_rules:
                                    logger.debug(
                                        f"動的セル名 {abs_path} に対してさらに変換ルールを適用"
                                    )
                                    for dynamic_rule in dynamic_rules:
                                        val = dynamic_rule.transform(val, wb)
                                abs_current = result
                                for abs_part in abs_parts[:-1]:
                                    if abs_part not in abs_current:
                                        abs_current[abs_part] = {}
                                    abs_current = abs_current[abs_part]
                                abs_current[abs_parts[-1]] = val
                        else:
                            # 相対指定
                            rel_path = f"{'.'.join(insert_keys)}.{key}"
                            rel_parts = rel_path.split(".")
                            # 相対指定でも変換ルールを適用
                            dynamic_rules = get_transform_rules(
                                array_transform_rules, rel_parts, rel_parts
                            )
                            if dynamic_rules:
                                logger.debug(
                                    f"相対セル名 {rel_path} に対してさらに変換ルールを適用"
                                )
                                for dynamic_rule in dynamic_rules:
                                    val = dynamic_rule.transform(val, wb)
                            rel_current = result
                            for rel_part in rel_parts[:-1]:
                                if rel_part not in rel_current:
                                    rel_current[rel_part] = {}
                                rel_current = rel_current[rel_part]
                            rel_current[rel_parts[-1]] = val
                    continue

                logger.debug(f"変換ルール{i+1}後の値: {current_value}")

            # function型の場合は追加のsplit/配列化処理はスキップ
            if (
                len(transform_rules) > 0
                and transform_rules[-1].transform_type == "function"
            ):
                logger.debug(
                    f"function型変換後の値: {current_value} (追加配列化処理はスキップ)"
                )

            # 最終結果を挿入
            insert_json_path(result, insert_keys, current_value, ".".join(insert_keys))
            continue

        logger.debug(f"配列化後の値: {value}")

        # 配列要素のパスの特別処理
        if len(path_keys) >= 2 and re.fullmatch(r"\d+", path_keys[1]):
            # departments.1.name のような配列要素パス、または parent.1.1 のような2次元配列パス
            array_name = path_keys[0]
            array_index = int(path_keys[1]) - 1  # 1-based to 0-based

            logger.debug(
                f"配列要素処理: array_name={array_name}, array_index={array_index}, path_keys={path_keys}"
            )

            # 配列が存在しない場合は作成
            if array_name not in result:
                result[array_name] = []
                logger.debug(f"新しい配列を作成: {array_name}")

            # 配列の参照を取得
            array_ref = result[array_name]

            # まだ辞書の場合は空の配列に変換
            if isinstance(array_ref, dict):
                logger.debug(f"辞書を空の配列に変換: {array_name}")
                result[array_name] = []
                array_ref = result[array_name]

            # 配列のサイズを必要に応じて拡張
            logger.debug(
                f"配列拡張前: len={len(array_ref)}, 必要インデックス={array_index}"
            )
            while len(array_ref) <= array_index:
                array_ref.append({})
                logger.debug(f"配列要素追加: len={len(array_ref)}")

            logger.debug(f"配列拡張後: len={len(array_ref)}")

            # 配列要素が存在することを確認
            if array_index >= len(array_ref):
                logger.error(
                    f"配列インデックス範囲外: array_index={array_index}, len={len(array_ref)}"
                )
                raise IndexError(f"配列インデックス {array_index} が範囲外です")

            # 配列要素への安全なアクセス
            try:
                current_element = array_ref[array_index]
                logger.debug(
                    f"配列要素取得成功: index={array_index}, type={type(current_element)}"
                )
            except (KeyError, IndexError) as e:
                logger.error(
                    f"配列要素アクセスエラー: {e}, array_type={type(array_ref)}, len={len(array_ref)}"
                )
                # デバッグのために配列の内容をログ出力
                logger.error(f"配列内容: {array_ref}")
                raise

            # 2次元配列パターンの確認（parent.1.1 のような場合）
            if len(path_keys) >= 3 and re.fullmatch(r"\d+", path_keys[2]):
                # 2次元配列のインデックス
                second_index = int(path_keys[2]) - 1  # 1-based to 0-based
                logger.debug(f"2次元配列パターン検出: second_index={second_index}")

                # 1次元目の配列要素が配列でない場合は配列に変換
                if not isinstance(current_element, list):
                    array_ref[array_index] = []
                    current_element = array_ref[array_index]
                    logger.debug(f"配列要素を配列に変換: index={array_index}")

                # 2次元目の配列のサイズを必要に応じて拡張
                while len(current_element) <= second_index:
                    current_element.append(None)
                    logger.debug(f"2次元配列要素追加: len={len(current_element)}")

                # 値を設定（3次元以降のパスがあれば再帰的に処理）
                if len(path_keys) > 3:
                    # さらに深い階層がある場合
                    if current_element[second_index] is None:
                        current_element[second_index] = {}
                    remaining_keys = path_keys[3:]
                    insert_json_path(
                        current_element[second_index],
                        remaining_keys,
                        value,
                        ".".join(remaining_keys),
                    )
                else:
                    # 2次元配列の最終要素に値を設定
                    current_element[second_index] = value
                    logger.debug(
                        f"2次元配列要素に値設定: "
                        f"[{array_index}][{second_index}] = {value}"
                    )
            else:
                # 1次元配列パターン（従来の処理）
                # 配列要素が辞書でない場合は辞書に変換
                if not isinstance(current_element, dict):
                    array_ref[array_index] = {}
                    current_element = array_ref[array_index]
                    logger.debug(f"配列要素を辞書に変換: index={array_index}")

                # 残りのパスを配列要素内に挿入
                if len(path_keys) > 2:
                    remaining_keys = path_keys[2:]
                    remaining_path = ".".join(remaining_keys)
                    insert_json_path(
                        current_element, remaining_keys, value, remaining_path
                    )
                else:
                    array_ref[array_index] = value
        else:
            # シンプルなJSONパス挿入
            insert_json_path(result, path_keys, value, ".".join(path_keys))

    # ワイルドカード変換の適用
    if array_transform_rules:
        logger.debug("ワイルドカード変換ルール適用開始")
        result = apply_wildcard_transforms(
            result, array_transform_rules, normalized_prefix
        )
        logger.debug("ワイルドカード変換ルール適用完了")

    # 二次元配列変換の適用
    if array_transform_rules:
        logger.debug("二次元配列変換ルール適用開始")
        result = apply_2d_array_transforms(
            result, array_transform_rules, wb, normalized_prefix
        )
        logger.debug("二次元配列変換ルール適用完了")

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


def _has_non_empty_content(obj: Any) -> bool:
    """オブジェクト全体に空でないコンテンツが含まれているかチェック"""
    if obj is None:
        return False

    if isinstance(obj, dict):
        return any(_has_non_empty_content(value) for value in obj.values())

    if isinstance(obj, list):
        return any(_has_non_empty_content(item) for item in obj)

    # スカラー値は空でないと判定
    return not DataCleaner.is_empty_value(obj)


def _is_empty_container(obj: Any) -> bool:
    """空のコンテナ（辞書・リスト）かどうかをチェック"""
    return isinstance(obj, (list, dict)) and len(obj) == 0


def prune_empty_elements(obj: Any, *, _has_sibling_data: Optional[bool] = None) -> Any:
    """
    再帰的に dict/list から空要素を除去する

    Args:
        obj: 処理対象のオブジェクト
        _has_sibling_data: 内部用パラメータ（外部使用禁止）

    Returns:
        空要素を除去した結果:
        - dict: 空要素のみの場合は {} または None
        - list: 空要素のみの場合は []
        - その他: そのまま返す
    """
    # 最初の呼び出し時のみ、全体的な非空データの存在をチェック
    if _has_sibling_data is None:
        _has_sibling_data = _has_non_empty_content(obj)

    # 型に応じた処理
    if isinstance(obj, dict):
        return _prune_dict(obj, _has_sibling_data)
    elif isinstance(obj, list):
        return _prune_list(obj, _has_sibling_data)
    else:
        return obj


def _prune_dict(obj: dict, has_sibling_data: bool) -> Optional[dict]:
    """辞書の空要素を除去"""
    if not obj:
        return {}

    # 子要素をプルーニング（Walrus演算子とdict内包表記を使用）
    pruned_items = {
        key: pruned_value
        for key, value in obj.items()
        if (
            pruned_value := prune_empty_elements(
                value, _has_sibling_data=has_sibling_data
            )
        )
        is not None
    }

    if not pruned_items:
        return None

    # 非空要素の存在をチェック（ジェネレータ式を使用）
    has_non_empty = any(
        not _is_empty_container(value) for value in pruned_items.values()
    )

    # 三項演算子を使用（読みやすさのため分割）
    if has_non_empty:
        return pruned_items
    else:
        return {} if has_sibling_data else None


def _prune_list(obj: list, has_sibling_data: bool) -> list:
    """リストの空要素を除去"""
    # 空でない要素のみを残す（Walrus演算子とリスト内包表記を使用）
    pruned_items = [
        pruned_item
        for item in obj
        if (
            pruned_item := prune_empty_elements(
                item, _has_sibling_data=has_sibling_data
            )
        )
        is not None
    ]

    if not pruned_items:
        return []

    # 非空要素の存在をチェック（ジェネレータ式を使用）
    has_non_empty = any(not _is_empty_container(item) for item in pruned_items)

    # 三項演算子を使用
    return pruned_items if has_non_empty else []


def write_data(
    data: Dict[str, Any],
    output_path: Path,
    output_format: str = "json",
    schema: Optional[Dict[str, Any]] = None,
    validator: Optional[Draft7Validator] = None,
    suppress_empty: bool = True,
) -> None:
    """
    データをファイルに書き出し（JSON/YAML対応）。
    バリデーションとソートはオプション。
    """
    base_name = output_path.stem
    output_dir = output_path.parent

    # 全フィールドが未設定の要素を除去
    data = prune_empty_elements(data)

    # 空値の除去
    if suppress_empty:
        data = DataCleaner.clean_empty_values(data, suppress_empty)
        if data is None:
            data = {}

    # バリデーション → エラーログ
    if validator:
        errors = list(validator.iter_errors(data))
        if errors:
            # エラーログファイルの作成
            log_file = output_dir / f"{base_name}.error.log"
            log_file.parent.mkdir(parents=True, exist_ok=True)
            with open(log_file, "w", encoding="utf-8") as f:
                for error in errors:
                    path_str = ".".join(str(p) for p in error.absolute_path)
                    msg = f"Validation error at {path_str}: {error.message}\n"
                    f.write(msg)
            # 最初のエラーをログに出力
            first_error = errors[0]
            logger.error(f"Validation error: {first_error.message}")

    # ソート処理
    if schema:
        data = reorder_json(data, schema)

    # datetime型を文字列に変換する関数
    def json_default(obj):
        if isinstance(obj, datetime.datetime):
            return obj.isoformat()
        if isinstance(obj, datetime.date):
            return obj.isoformat()
        return str(obj)

    # ファイル書き出し
    output_dir.mkdir(parents=True, exist_ok=True)

    if output_format == "yaml":
        # YAML形式で出力
        with output_path.open("w", encoding="utf-8") as f:
            # datetime オブジェクトを文字列に変換
            yaml_data = json.loads(json.dumps(data, default=json_default))
            yaml.dump(
                yaml_data, f, default_flow_style=False, allow_unicode=True, indent=2
            )
    else:
        # JSON形式で出力（デフォルト）
        with output_path.open("w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2, default=json_default)

    logger.info(f"ファイルの出力に成功しました: {output_path}")


# =============================================================================
# Border Rectangle Detection
# =============================================================================


def has_border(worksheet, row, col, side):
    """
    指定セルの指定方向に罫線があるかチェック
    隣接セルの境界線も考慮して、人間の目で見たときの連続性を判定
    """
    try:
        # 自セルの罫線をチェック
        cell = worksheet.cell(row=row, column=col)
        border = getattr(cell.border, side, None)
        if border is not None and border.style is not None:
            return True

        # 隣接セルの境界線もチェック
        adjacent_row, adjacent_col = row, col
        adjacent_side = side

        if side == "top":
            adjacent_row = row - 1
            adjacent_side = "bottom"
        elif side == "bottom":
            adjacent_row = row + 1
            adjacent_side = "top"
        elif side == "left":
            adjacent_col = col - 1
            adjacent_side = "right"
        elif side == "right":
            adjacent_col = col + 1
            adjacent_side = "left"

        # 隣接セルが存在する場合、その境界線をチェック
        if adjacent_row > 0 and adjacent_col > 0:
            try:
                adjacent_cell = worksheet.cell(row=adjacent_row, column=adjacent_col)
                adjacent_border = getattr(adjacent_cell.border, adjacent_side, None)
                if adjacent_border is not None and adjacent_border.style is not None:
                    return True
            except:
                pass

        return False
    except:
        return False


def detect_rectangular_regions(worksheet, cell_names_map=None):
    """
    罫線で囲まれた四角形領域を検出
    左上から大きい順にソートして返す
    """
    regions = []

    # セル名マップがある場合はその範囲、ない場合は制限された範囲を対象
    if cell_names_map:
        min_row = min(row for row, col in cell_names_map.keys())
        max_row = max(row for row, col in cell_names_map.keys())
        min_col = min(col for row, col in cell_names_map.keys())
        max_col = max(col for row, col in cell_names_map.keys())
    else:
        # セル名がない場合は実際のワークシートの有効範囲内に制限
        min_row, min_col = 1, 1
        # 実際にデータがある範囲を上限とする
        actual_max_row = worksheet.max_row if worksheet.max_row else 30
        actual_max_col = worksheet.max_column if worksheet.max_column else 30
        max_row, max_col = min(actual_max_row, 30), min(actual_max_col, 30)

    logger.debug(f"四角形検出範囲: 行{min_row}-{max_row}, 列{min_col}-{max_col}")

    # 各セルを起点として四角形を検出（大きい領域から小さい領域へ）
    for top in range(min_row, max_row + 1):
        for left in range(min_col, max_col + 1):
            # 幅と高さを独立して変化させて長方形も検出
            # セル名領域のサイズを上限とする
            max_width = min(max_col - left + 1, max_col - min_col + 1)
            max_height = min(max_row - top + 1, max_row - min_row + 1)

            # 大きい面積から小さい面積へソートして検出
            size_combinations = []
            for width in range(1, max_width + 1):
                for height in range(1, max_height + 1):
                    area = width * height
                    size_combinations.append((area, width, height))

            # 面積の大きい順にソート
            size_combinations.sort(reverse=True)

            for area, width, height in size_combinations:
                right = left + width - 1
                bottom = top + height - 1

                if bottom > max_row or right > max_col:
                    continue

                # 四角形の罫線完成度を計算
                completeness = calculate_border_completeness(
                    worksheet, top, left, bottom, right
                )

                # 人の目で見て完全に囲まれた領域のみを対象（100%完成度）
                if completeness >= 1.0:
                    # セル名がある場合はその中に含まれるかチェック
                    if cell_names_map:
                        cell_names_in_region = get_cell_names_in_region(
                            cell_names_map, top, left, bottom, right
                        )
                        if not cell_names_in_region:
                            continue

                    region_tuple = (top, left, bottom, right, completeness)
                    regions.append(region_tuple)

    # 大きい順、完成度順、左上位置順でソート
    regions.sort(
        key=lambda r: (-(r[2] - r[0] + 1) * (r[3] - r[1] + 1), -r[4], r[0], r[1])
    )

    # 重複する領域を除去（より意味のある大きな領域を優先）
    filtered_regions = []
    for region in regions:
        top, left, bottom, right, completeness = region
        region_area = (bottom - top + 1) * (right - left + 1)

        # 既存の領域と重複していないかチェック
        is_redundant = False
        for existing_region in filtered_regions:
            ex_top, ex_left, ex_bottom, ex_right, ex_completeness = existing_region
            ex_area = (ex_bottom - ex_top + 1) * (ex_right - ex_left + 1)

            # 既存の大きな領域に完全に包含される小さな領域は除外
            if (
                top >= ex_top
                and left >= ex_left
                and bottom <= ex_bottom
                and right <= ex_right
                and region_area < ex_area
            ):
                is_redundant = True
                break

            # 新しい領域が既存の小さな領域を包含する場合は既存を削除
            elif (
                ex_top >= top
                and ex_left >= left
                and ex_bottom <= bottom
                and ex_right <= right
                and ex_area < region_area
            ):
                filtered_regions.remove(existing_region)

        if not is_redundant:
            filtered_regions.append(region)

    logger.debug(f"検出された四角形領域数: {len(filtered_regions)}")
    return filtered_regions


def is_complete_rectangle(worksheet, top, left, bottom, right):
    """指定範囲が完全な四角形の罫線で囲まれているかチェック"""
    try:
        # 上辺をチェック
        for col in range(left, right + 1):
            if not has_border(worksheet, top, col, "top"):
                return False

        # 下辺をチェック
        for col in range(left, right + 1):
            if not has_border(worksheet, bottom, col, "bottom"):
                return False

        # 左辺をチェック
        for row in range(top, bottom + 1):
            if not has_border(worksheet, row, left, "left"):
                return False

        # 右辺をチェック
        for row in range(top, bottom + 1):
            if not has_border(worksheet, row, right, "right"):
                return False

        return True
    except:
        return False


def get_cell_names_in_region(cell_names_map, top, left, bottom, right):
    """指定領域内のセル名を取得"""
    cell_names = []
    for row in range(top, bottom + 1):
        for col in range(left, right + 1):
            if (row, col) in cell_names_map:
                cell_names.append(cell_names_map[(row, col)])
    return cell_names


def detect_hierarchical_regions(worksheet, cell_names_map=None):
    """
    ツリー型階層構造を検出
    左から右へのドリルダウン構造を前提として、各階層にセル名が存在する構成を想定
    """
    logger.info("ツリー型階層構造の検出を開始")

    # すべての完全な四角形領域を検出
    all_regions = detect_all_complete_rectangles(worksheet, cell_names_map)

    if not all_regions:
        logger.warning("完全な四角形領域が検出されませんでした")
        return []

    logger.info(f"検出された全領域数: {len(all_regions)}")
    for i, region in enumerate(all_regions[:5]):  # 最初の5個を表示
        bounds = region["bounds"]
        logger.info(
            f"  領域{i}: ({bounds[0]},{bounds[1]},{bounds[2]},{bounds[3]}) 面積:{region['area']}"
        )

    # ツリー構造に適した領域選択
    tree_regions = select_tree_structure_regions(all_regions, cell_names_map)

    logger.info(f"ツリー構造用選択領域数: {len(tree_regions)}")

    # セル名ベースでの重複除去
    tree_regions = select_largest_regions_by_cell_names(tree_regions)

    # 階層関係を構築（左右方向の階層展開を考慮）
    hierarchy = build_tree_hierarchy(tree_regions, cell_names_map)

    return hierarchy


def select_tree_structure_regions(all_regions, cell_names_map=None):
    """
    ツリー型階層構造に適した領域を選択
    左から右への展開パターンと各階層でのセル名存在を重視
    """
    if not all_regions:
        return []

    # セル名を持つ領域を基軸として分析
    regions_with_names = [r for r in all_regions if r.get("cell_names")]
    regions_without_names = [r for r in all_regions if not r.get("cell_names")]

    logger.info(
        f"セル名あり領域: {len(regions_with_names)}, セル名なし領域: {len(regions_without_names)}"
    )

    selected_regions = []

    # 1. セル名を持つ領域は階層の各レベルを表すため必ず含める
    selected_regions.extend(regions_with_names)

    # 2. セル名なし領域から構造的に重要なものを選択
    for region in regions_without_names:
        area = region["area"]
        bounds = region["bounds"]
        top, left, bottom, right = bounds

        # 大きな全体コンテナ（ルート領域候補）
        if area >= 200:
            selected_regions.append(region)
            logger.debug(f"ルート領域候補を追加: {bounds} (面積: {area})")
            continue

        # 中程度の領域：セル名領域を適切に包含する階層コンテナ
        if area >= 20:
            # この領域がセル名領域を包含しているかチェック
            contains_named_regions = []
            for named_region in regions_with_names:
                if is_region_contained(named_region, region):
                    contains_named_regions.append(named_region)

            if len(contains_named_regions) >= 1:
                # 1つ以上のセル名領域を包含する場合は階層コンテナとして採用
                selected_regions.append(region)
                logger.debug(
                    f"階層コンテナを追加: {bounds} (面積: {area}, 包含: {len(contains_named_regions)}個)"
                )

        # 小さな領域でも形状的に意味があるもの
        elif area >= 8:
            # 適度な矩形形状で、ツリー構造の一部となり得るもの
            width = right - left + 1
            height = bottom - top + 1
            ratio = max(width, height) / min(width, height)

            if ratio <= 3.0:  # アスペクト比が3以下
                # セル名領域との位置関係をチェック
                has_structural_meaning = False
                for named_region in regions_with_names:
                    named_bounds = named_region["bounds"]
                    # 隣接または近接している場合
                    if (
                        abs(bounds[0] - named_bounds[0]) <= 2
                        or abs(bounds[1] - named_bounds[1]) <= 2
                    ):
                        has_structural_meaning = True
                        break

                if has_structural_meaning:
                    selected_regions.append(region)
                    logger.debug(f"構造的小領域を追加: {bounds} (面積: {area})")

    logger.info(f"選択された構造的領域数: {len(selected_regions)}")

    return selected_regions


def build_tree_hierarchy(regions, cell_names_map=None):
    """
    ツリー型階層構造を構築
    左から右への展開パターンを重視した親子関係の判定
    """
    # 領域をIDで管理
    for i, region in enumerate(regions):
        region["id"] = i
        region["parent"] = None
        region["children"] = []
        region["level"] = 0
        region["tree_position"] = analyze_tree_position(region, cell_names_map)

    # 左から右、上から下の順序でソート（ツリー構造の自然な順序）
    sorted_regions = sorted(
        enumerate(regions),
        key=lambda x: (
            x[1]["bounds"][1],  # 左端の列位置
            x[1]["bounds"][0],  # 上端の行位置
            -x[1]["area"],  # 面積（大きい順）
        ),
    )

    # 階層関係を構築
    for i, (idx_i, region) in enumerate(sorted_regions):
        potential_parents = []

        for j, (idx_j, other_region) in enumerate(sorted_regions):
            if idx_i != idx_j and is_region_contained(region, other_region):
                # 面積差による階層判定
                area_ratio = other_region["area"] / region["area"]

                # セル名による階層判定の調整
                child_has_names = bool(region.get("cell_names", []))
                parent_has_names = bool(other_region.get("cell_names", []))

                min_ratio = 1.2  # 基本的な最小比率

                if child_has_names and parent_has_names:
                    min_ratio = 1.1  # セル名同士は緩い条件
                elif parent_has_names and not child_has_names:
                    min_ratio = 1.3
                elif not parent_has_names and child_has_names:
                    min_ratio = 1.5  # 子にセル名がある場合は厳格に
                else:
                    min_ratio = 2.0  # 両方ともセル名なしは厳格に

                if area_ratio >= min_ratio:
                    # ツリー展開パターンをチェック
                    if is_valid_tree_relationship(region, other_region):
                        potential_parents.append((idx_j, other_region, area_ratio))

        # 最適な親を選択
        if potential_parents:
            best_parent_idx, best_parent, best_ratio = min(
                potential_parents, key=lambda x: x[2]
            )

            # 既存の親関係をチェック
            current_parent_id = region["parent"]
            if (
                current_parent_id is None
                or regions[current_parent_id]["area"] > best_parent["area"]
            ):
                # 既存の親関係を解除
                if current_parent_id is not None:
                    regions[current_parent_id]["children"].remove(region["id"])

                # 新しい親子関係を設定
                region["parent"] = best_parent["id"]
                best_parent["children"].append(region["id"])

                logger.debug(
                    f"ツリー階層設定: 領域{region['id']} → 親領域{best_parent['id']} "
                    f"(面積比: {best_ratio:.2f})"
                )

    # 階層レベルを計算
    calculate_hierarchy_levels(regions)

    # 階層構造を整理
    hierarchy = organize_hierarchy(regions)

    return hierarchy


def analyze_tree_position(region, cell_names_map=None):
    """
    領域のツリー内での位置を分析
    """
    bounds = region["bounds"]
    top, left, bottom, right = bounds

    return {
        "left_position": left,
        "top_position": top,
        "width": right - left + 1,
        "height": bottom - top + 1,
        "has_cell_names": bool(region.get("cell_names", [])),
    }


def is_valid_tree_relationship(child_region, parent_region):
    """
    ツリー構造として有効な親子関係かチェック
    左から右への展開パターンを考慮
    """
    child_bounds = child_region["bounds"]
    parent_bounds = parent_region["bounds"]

    child_top, child_left, child_bottom, child_right = child_bounds
    parent_top, parent_left, parent_bottom, parent_right = parent_bounds

    # 基本的な包含関係
    if not is_region_contained(child_region, parent_region):
        return False

    # 線状領域のチェーン関係を除外
    child_height = child_bottom - child_top + 1
    child_width = child_right - child_left + 1
    parent_height = parent_bottom - parent_top + 1
    parent_width = parent_right - parent_left + 1

    # 1行または1列の連続は階層として意味がない
    if child_height == 1 and parent_height <= 2:
        return False
    if child_width == 1 and parent_width <= 2:
        return False

    # 非常に小さな領域同士の関係は除外（ただしセル名がある場合は例外）
    child_has_names = bool(child_region.get("cell_names", []))
    parent_has_names = bool(parent_region.get("cell_names", []))

    if child_region["area"] < 10 and parent_region["area"] < 20:
        if not (child_has_names or parent_has_names):
            return False

    return True


def detect_all_complete_rectangles(worksheet, cell_names_map=None):
    """
    重複除去なしで全ての完全な四角形領域を検出
    ただし、意味のある階層構造を形成する領域のみを対象とする
    """
    regions = []

    # セル名マップがある場合はその範囲全体、ない場合はワークシート全域を対象
    if cell_names_map:
        min_row = min(row for row, col in cell_names_map.keys())
        max_row = max(row for row, col in cell_names_map.keys())
        min_col = min(col for row, col in cell_names_map.keys())
        max_col = max(col for row, col in cell_names_map.keys())
        # セル名範囲を若干拡張して境界領域も検出対象とする
        min_row = max(1, min_row - 2)
        max_row = min(worksheet.max_row or 100, max_row + 2)
        min_col = max(1, min_col - 2)
        max_col = min(worksheet.max_column or 100, max_col + 2)
    else:
        # セル名がない場合はワークシートの実際の使用範囲全域を対象
        min_row, min_col = 1, 1
        max_row = worksheet.max_row if worksheet.max_row else 100
        max_col = worksheet.max_column if worksheet.max_column else 100

    logger.debug(f"四角形検出範囲: 行{min_row}-{max_row}, 列{min_col}-{max_col}")

    # ツリー構造に適した階層を形成する領域のみを検出
    # 最小サイズ制限を緩和（面積2以上）し、指定範囲全域を検査対象とする
    min_area = 2
    detected_regions_set = set()  # 重複検出防止用

    # デバッグ用カウンタ
    total_checked = 0
    border_found = 0

    # 各セルを起点として四角形を検出
    for top in range(min_row, max_row + 1):
        for left in range(min_col, max_col + 1):
            # 幅と高さの組み合わせを効率的に探索
            max_width = min(max_col - left + 1, max_col - min_col + 1)
            max_height = min(max_row - top + 1, max_row - min_row + 1)

            # 意味のある階層を形成するサイズのみを対象
            # 小さいものから大きいものへと順次検出（階層構造の基礎から構築）
            size_combinations = []
            for width in range(1, max_width + 1):
                for height in range(1, max_height + 1):
                    area = width * height
                    if area >= min_area:  # 最小面積制限
                        size_combinations.append((area, width, height))

            # 面積の小さい順にソート（階層の基礎から構築）
            size_combinations.sort()

            for area, width, height in size_combinations:
                right = left + width - 1
                bottom = top + height - 1

                if bottom > max_row or right > max_col:
                    continue

                # 重複チェック
                bounds_key = (top, left, bottom, right)
                if bounds_key in detected_regions_set:
                    continue

                # 検出状況をカウント
                total_checked += 1

                # 四角形の罫線完成度を計算
                completeness = calculate_border_completeness(
                    worksheet, top, left, bottom, right
                )

                # 人の目で見て完全に囲まれた領域のみを対象（100%完成度）
                if completeness >= 1.0:
                    border_found += 1
                    # セル名がある場合はその中に含まれるかチェック
                    cell_names_in_region = []
                    if cell_names_map:
                        cell_names_in_region = get_cell_names_in_region(
                            cell_names_map, top, left, bottom, right
                        )

                    # より厳格なフィルタリング：ツリー構造に適した領域のみを採用
                    should_include = False

                    # 1. セル名がある領域は常に採用（ツリーの各ノード）
                    if cell_names_in_region:
                        should_include = True
                    # 2. 大きな領域（50セル以上）は構造的コンテナとして採用
                    elif area >= 50:
                        should_include = True
                    # 3. 中程度の領域（10-49セル）はバランスの取れた形状なら採用
                    elif area >= 10:
                        ratio = max(width, height) / min(width, height)
                        if ratio <= 5.0:  # アスペクト比が5以下
                            should_include = True
                    # 4. 小さな領域（4-9セル）は正方形に近い形状のみ
                    elif area >= 4:
                        ratio = max(width, height) / min(width, height)
                        if ratio <= 3.0:  # アスペクト比が3以下
                            should_include = True
                    # 5. 最小領域（2-3セル）は正方形のみ
                    elif area >= 2:
                        ratio = max(width, height) / min(width, height)
                        if ratio <= 1.5:  # アスペクト比が1.5以下（ほぼ正方形）
                            should_include = True

                    if should_include:
                        region = {
                            "bounds": (top, left, bottom, right),
                            "area": area,
                            "completeness": completeness,
                            "cell_names": cell_names_in_region,
                        }
                        regions.append(region)
                        detected_regions_set.add(bounds_key)

    # 同一セル名を含む領域から最大のものを選択
    regions = select_largest_regions_by_cell_names(regions)

    # 面積の大きい順、左上位置順でソート
    regions.sort(key=lambda r: (-r["area"], r["bounds"][0], r["bounds"][1]))

    logger.info(
        f"検出統計: チェック総数{total_checked}, 完全境界{border_found}, 最終採用{len(regions)}"
    )
    return regions


def select_largest_regions_by_cell_names(regions):
    """
    同一のセル名を内包する領域の中で一番大きい領域のみを採用
    繰り返しコンテナとしての識別を改善する
    """
    logger.info(f"セル名ベース選択開始: {len(regions)}個の領域を処理")

    if not regions:
        return regions

    # セル名ごとに領域をグループ化
    cell_name_groups = {}
    regions_without_names = []

    for region in regions:
        cell_names = region.get("cell_names", [])
        if cell_names:
            # cell_namesが正しくリストであることを確認
            if not isinstance(cell_names, list):
                cell_names = [cell_names]

            # 複数のセル名がある場合は、それぞれのセル名でグループ化
            for cell_name in cell_names:
                # cell_nameが文字列であることを確認
                if isinstance(cell_name, (list, tuple)):
                    # もしリストまたはタプルなら、その中身を展開
                    for name in cell_name:
                        if name not in cell_name_groups:
                            cell_name_groups[name] = []
                        cell_name_groups[name].append(region)
                else:
                    if cell_name not in cell_name_groups:
                        cell_name_groups[cell_name] = []
                    cell_name_groups[cell_name].append(region)
        else:
            # セル名がない領域はそのまま保持
            regions_without_names.append(region)

    logger.debug(f"セル名グループ分析:")
    logger.debug(f"  セル名なし領域: {len(regions_without_names)}個")
    for cell_name, group_regions in cell_name_groups.items():
        logger.debug(f"  セル名'{cell_name}': {len(group_regions)}個の領域")
        for region in group_regions:
            logger.debug(f"    {region['bounds']} (面積: {region['area']})")

    # 各セル名グループから最大の領域を選択
    selected_regions = []
    total_excluded = 0

    for cell_name, group_regions in cell_name_groups.items():
        if len(group_regions) == 1:
            # 1つしかない場合はそのまま採用
            selected_regions.append(group_regions[0])
            logger.debug(
                f"セル名'{cell_name}': 単一領域のため採用 {group_regions[0]['bounds']}"
            )
        else:
            # 複数ある場合は最大面積の領域を選択
            largest_region = max(group_regions, key=lambda r: r["area"])
            selected_regions.append(largest_region)
            excluded_count = len(group_regions) - 1
            total_excluded += excluded_count

            logger.info(
                f"セル名'{cell_name}': {len(group_regions)}個から最大領域を選択"
            )
            logger.info(
                f"  採用: {largest_region['bounds']} (面積: {largest_region['area']})"
            )

            # 選択されなかった領域をログ出力
            for region in group_regions:
                if region != largest_region:
                    logger.info(f"  除外: {region['bounds']} (面積: {region['area']})")

    # セル名がない領域も追加
    selected_regions.extend(regions_without_names)

    logger.info(
        f"セル名ベース選択結果: {len(regions)}個 → {len(selected_regions)}個 (除外: {total_excluded}個)"
    )

    return selected_regions


def build_region_hierarchy(regions, cell_names_map=None):
    """
    領域リストから階層構造を構築
    セル名の意味を重視し、境界共有を考慮した包含関係で親子関係を判定
    """
    # 領域をIDで管理
    for i, region in enumerate(regions):
        region["id"] = i
        region["parent"] = None
        region["children"] = []
        region["level"] = 0

    # 面積の大きい順でソート（大きな領域が親になりやすくする）
    sorted_regions = sorted(
        enumerate(regions),
        key=lambda x: (-x[1]["area"], x[1]["bounds"][0], x[1]["bounds"][1]),
    )

    # 包含関係を分析して親子関係を設定
    # セル名の意味を重視した階層構築
    for i, (idx_i, region) in enumerate(sorted_regions):
        potential_parents = []

        # より大きな領域から親候補を探す
        for j, (idx_j, other_region) in enumerate(sorted_regions):
            if idx_i != idx_j and is_region_contained(region, other_region):
                # 面積差が意味のある差であることを確認
                area_ratio = other_region["area"] / region["area"]

                # セル名がある領域同士の場合は、より緩い条件で親子関係を認める
                child_has_names = bool(region.get("cell_names", []))
                parent_has_names = bool(other_region.get("cell_names", []))

                min_ratio = 1.2  # デフォルトの最小比率

                if child_has_names and parent_has_names:
                    # 両方にセル名がある場合は意味のある階層の可能性が高い
                    min_ratio = 1.2
                elif parent_has_names and not child_has_names:
                    # 親にセル名があり子にない場合も意味のある階層
                    min_ratio = 1.3
                elif not parent_has_names and child_has_names:
                    # 子にセル名があり親にない場合は厳格に
                    min_ratio = 2.0
                else:
                    # 両方ともセル名がない場合は最も厳格に
                    min_ratio = 2.0

                if area_ratio >= min_ratio:
                    # チェーン的な関係を除外（線状の領域の連続は階層として意味がない）
                    region_bounds = region["bounds"]
                    other_bounds = other_region["bounds"]

                    # 子領域の形状をチェック
                    child_height = region_bounds[2] - region_bounds[0] + 1
                    child_width = region_bounds[3] - region_bounds[1] + 1

                    # 親領域の形状をチェック
                    parent_height = other_bounds[2] - other_bounds[0] + 1
                    parent_width = other_bounds[3] - other_bounds[1] + 1

                    # 線状領域（1行または1列）のチェーン関係を除外
                    is_chain_relationship = False

                    # 子が1行で親も近い高さの場合（行のチェーン）
                    if child_height == 1 and parent_height <= 2:
                        is_chain_relationship = True

                    # 子が1列で親も近い幅の場合（列のチェーン）
                    if child_width == 1 and parent_width <= 2:
                        is_chain_relationship = True

                    # 小さな領域同士の微細な包含関係を除外（ただしセル名がある場合は例外）
                    if region["area"] < 10 and other_region["area"] < 20:
                        if not (child_has_names or parent_has_names):
                            is_chain_relationship = True

                    if not is_chain_relationship:
                        potential_parents.append((idx_j, other_region, area_ratio))

        # 最も適切な親を選択（面積比が最小で、かつ直接的な包含関係）
        if potential_parents:
            # 面積比が最小の包含領域を親として選択（最も直接的な親）
            best_parent_idx, best_parent, best_ratio = min(
                potential_parents, key=lambda x: x[2]
            )

            # 包含関係のタイプを分析
            containment_type = get_containment_type(region, best_parent)
            shared_boundaries = analyze_boundary_sharing(region, best_parent)

            # 既存の親関係をチェック
            current_parent_id = region["parent"]
            if (
                current_parent_id is None
                or regions[current_parent_id]["area"] > best_parent["area"]
            ):
                # 既存の親関係を解除
                if current_parent_id is not None:
                    regions[current_parent_id]["children"].remove(region["id"])

                # 新しい親子関係を設定
                region["parent"] = best_parent["id"]
                region["containment_type"] = containment_type
                region["shared_boundaries"] = shared_boundaries
                region["area_ratio"] = best_ratio
                best_parent["children"].append(region["id"])

                logger.debug(
                    f"親子関係設定: 領域{region['id']} → 親領域{best_parent['id']} "
                    f"(面積比: {best_ratio:.2f}, タイプ: {containment_type})"
                )

    # 重複した子関係を整理
    for region in regions:
        region["children"] = list(set(region["children"]))  # 重複を除去

    # 階層レベルを計算
    calculate_hierarchy_levels(regions)

    # 階層構造を整理
    hierarchy = organize_hierarchy(regions)

    return hierarchy


def is_region_contained(inner_region, outer_region):
    """
    inner_regionがouter_regionに包含されているかチェック
    一部の境界線を共有している場合も包含関係として認識する
    例: A1:Z14とB2:Z14は右辺と下辺を共有していても包含関係とする
    """
    inner_top, inner_left, inner_bottom, inner_right = inner_region["bounds"]
    outer_top, outer_left, outer_bottom, outer_right = outer_region["bounds"]

    # 面積チェック（内側領域が外側領域より小さい必要がある）
    if inner_region["area"] >= outer_region["area"]:
        return False

    # 包含関係のチェック（境界共有を許可）
    # 内側領域の境界が外側領域の境界と同じか内側にある
    is_contained = (
        inner_top >= outer_top
        and inner_left >= outer_left
        and inner_bottom <= outer_bottom
        and inner_right <= outer_right
    )

    # 完全に同じ領域は包含関係ではない
    if (
        inner_top == outer_top
        and inner_left == outer_left
        and inner_bottom == outer_bottom
        and inner_right == outer_right
    ):
        return False

    return is_contained


def is_region_overlapping(region_a, region_b):
    """
    2つの領域が重複しているかチェック
    包含関係でない場合の重複を検出
    """
    a_top, a_left, a_bottom, a_right = region_a["bounds"]
    b_top, b_left, b_bottom, b_right = region_b["bounds"]

    # 重複なしの条件
    if a_bottom < b_top or b_bottom < a_top or a_right < b_left or b_right < a_left:
        return False

    return True


def analyze_boundary_sharing(inner_region, outer_region):
    """
    2つの領域間での境界共有を詳細分析
    """
    inner_top, inner_left, inner_bottom, inner_right = inner_region["bounds"]
    outer_top, outer_left, outer_bottom, outer_right = outer_region["bounds"]

    shared_boundaries = []

    # 上辺の共有チェック
    if (
        inner_top == outer_top
        and inner_left >= outer_left
        and inner_right <= outer_right
    ):
        shared_boundaries.append("top")

    # 下辺の共有チェック
    if (
        inner_bottom == outer_bottom
        and inner_left >= outer_left
        and inner_right <= outer_right
    ):
        shared_boundaries.append("bottom")

    # 左辺の共有チェック
    if (
        inner_left == outer_left
        and inner_top >= outer_top
        and inner_bottom <= outer_bottom
    ):
        shared_boundaries.append("left")

    # 右辺の共有チェック
    if (
        inner_right == outer_right
        and inner_top >= outer_top
        and inner_bottom <= outer_bottom
    ):
        shared_boundaries.append("right")

    return shared_boundaries


def get_containment_type(inner_region, outer_region):
    """
    包含関係のタイプを判定
    """
    shared_boundaries = analyze_boundary_sharing(inner_region, outer_region)

    if not shared_boundaries:
        return "strict_containment"  # 厳密な包含（境界共有なし）
    elif len(shared_boundaries) == 1:
        return f"boundary_shared_{shared_boundaries[0]}"  # 1辺共有
    elif len(shared_boundaries) == 2:
        return f"boundary_shared_{'+'.join(shared_boundaries)}"  # 2辺共有
    elif len(shared_boundaries) == 3:
        return f"boundary_shared_{'+'.join(shared_boundaries)}"  # 3辺共有
    else:
        return "boundary_shared_all"  # 全辺共有（同一領域、通常は発生しない）


def calculate_overlap_area(region_a, region_b):
    """
    2つの領域の重複面積を計算
    """
    a_top, a_left, a_bottom, a_right = region_a["bounds"]
    b_top, b_left, b_bottom, b_right = region_b["bounds"]

    # 重複領域の境界を計算
    overlap_top = max(a_top, b_top)
    overlap_left = max(a_left, b_left)
    overlap_bottom = min(a_bottom, b_bottom)
    overlap_right = min(a_right, b_right)

    # 重複がない場合
    if overlap_top > overlap_bottom or overlap_left > overlap_right:
        return 0

    # 重複面積を計算
    overlap_width = overlap_right - overlap_left + 1
    overlap_height = overlap_bottom - overlap_top + 1

    return overlap_width * overlap_height


def calculate_hierarchy_levels(regions):
    """
    各領域の階層レベルを計算
    """

    def calc_level(region_id, regions):
        region = regions[region_id]
        if region["parent"] is None:
            region["level"] = 0
        else:
            parent_region = regions[region["parent"]]
            if parent_region["level"] == 0 and parent_region["parent"] is not None:
                calc_level(region["parent"], regions)
            region["level"] = parent_region["level"] + 1

    for region in regions:
        if region["level"] == 0 and region["parent"] is not None:
            calc_level(region["id"], regions)


def organize_hierarchy(regions):
    """
    階層構造を整理してツリー形式で返す
    """
    # ルート領域を特定
    root_regions = [region for region in regions if region["parent"] is None]

    def build_tree(region):
        tree_node = {
            "id": region["id"],
            "bounds": region["bounds"],
            "area": region["area"],
            "level": region["level"],
            "completeness": region["completeness"],
            "cell_names": region["cell_names"],
            "children": [],
        }

        # 包含関係情報を追加
        if "containment_type" in region:
            tree_node["containment_type"] = region["containment_type"]
        if "shared_boundaries" in region:
            tree_node["shared_boundaries"] = region["shared_boundaries"]

        # 子領域を再帰的に構築
        for child_id in region["children"]:
            child_region = regions[child_id]
            tree_node["children"].append(build_tree(child_region))

        # 子領域を面積の大きい順でソート
        tree_node["children"].sort(
            key=lambda x: (-x["area"], x["bounds"][0], x["bounds"][1])
        )

        return tree_node

    # ルートから階層ツリーを構築
    hierarchy_tree = []
    for root_region in root_regions:
        hierarchy_tree.append(build_tree(root_region))

    # ルート領域を面積の大きい順でソート
    hierarchy_tree.sort(key=lambda x: (-x["area"], x["bounds"][0], x["bounds"][1]))

    return hierarchy_tree


def analyze_region_relationships(hierarchy_tree):
    """
    階層構造の関係を分析してレポートを生成
    """
    analysis = {
        "total_regions": 0,
        "max_depth": 0,
        "root_regions": len(hierarchy_tree),
        "region_details": [],
    }

    def analyze_node(node, depth=0):
        analysis["total_regions"] += 1
        analysis["max_depth"] = max(analysis["max_depth"], depth)

        top, left, bottom, right = node["bounds"]
        width = right - left + 1
        height = bottom - top + 1

        region_info = {
            "id": node["id"],
            "bounds": f"行{top}-{bottom}, 列{left}-{right}",
            "size": f"{width}x{height}",
            "area": node["area"],
            "level": depth,
            "children_count": len(node["children"]),
            "cell_names": node["cell_names"],
            "completeness": node["completeness"],
        }
        analysis["region_details"].append(region_info)

        # 子ノードを再帰的に分析
        for child in node["children"]:
            analyze_node(child, depth + 1)

    for root in hierarchy_tree:
        analyze_node(root)

    return analysis


def filter_nested_regions(regions):
    """
    ネストした領域を整理し、同じセル名を含む最大領域のみを保持
    繰り返しコンテナの識別も行う
    """
    filtered = []

    # セル名ごとに最大領域を特定
    cell_name_to_max_region = {}

    for region in regions:
        for cell_name in region["cell_names"]:
            if cell_name not in cell_name_to_max_region:
                cell_name_to_max_region[cell_name] = region
            elif region["area"] > cell_name_to_max_region[cell_name]["area"]:
                cell_name_to_max_region[cell_name] = region

    # 繰り返しコンテナを識別
    unique_regions = list(set(cell_name_to_max_region.values()))

    for region in unique_regions:
        # セル名の階層構造を解析
        hierarchy_info = analyze_cell_name_hierarchy(region["cell_names"])
        region["hierarchy_info"] = hierarchy_info

        # 繰り返しパターンを検出
        repeat_info = detect_repeat_pattern(region["cell_names"])
        region["repeat_info"] = repeat_info

        filtered.append(region)

    return filtered


def analyze_cell_name_hierarchy(cell_names):
    """セル名の階層構造を解析"""
    hierarchy_levels = {}

    for cell_name in cell_names:
        if not cell_name.startswith("json."):
            continue

        parts = cell_name.split(".")
        # 数値部分とフィールド部分を分離
        hierarchy_path = []
        numeric_indices = []

        for part in parts[1:]:  # 'json'を除く
            if part.isdigit() or part == "*":
                numeric_indices.append(part)
            else:
                hierarchy_path.append(part)

        level = len(hierarchy_path)
        if level not in hierarchy_levels:
            hierarchy_levels[level] = []
        hierarchy_levels[level].append(
            {
                "cell_name": cell_name,
                "hierarchy_path": hierarchy_path,
                "numeric_indices": numeric_indices,
            }
        )

    return hierarchy_levels


def detect_repeat_pattern(cell_names):
    """繰り返しパターンを検出"""
    repeat_groups = {}

    for cell_name in cell_names:
        if not cell_name.startswith("json."):
            continue

        # 数値インデックスを除いたベースパターンを作成
        parts = cell_name.split(".")
        base_pattern = []

        for part in parts:
            if part.isdigit():
                base_pattern.append("*")  # 数値を*で置換
            else:
                base_pattern.append(part)

        pattern_key = ".".join(base_pattern)

        if pattern_key not in repeat_groups:
            repeat_groups[pattern_key] = []
        repeat_groups[pattern_key].append(cell_name)

    # 2つ以上のセル名を持つパターンのみを繰り返しとして認識
    return {k: v for k, v in repeat_groups.items() if len(v) > 1}


def extract_cell_names_from_workbook(workbook):
    """ワークブックから名前付き範囲のjson.*セル名を抽出"""
    cell_names_map = {}

    # 名前付き範囲から座標とセル名を取得
    for name, defined_name in workbook.defined_names.items():
        if name.startswith("json."):
            try:
                for sheet_name, coord in defined_name.destinations:
                    # 座標文字列を解析（例: "$S$2" -> (19, 2)）
                    coord_clean = coord.replace("$", "")
                    if ":" in coord_clean:
                        # 範囲の場合は左上のセルを使用
                        coord_clean = coord_clean.split(":")[0]

                    # 列文字を数値に変換
                    col_match = ""
                    row_match = ""
                    for char in coord_clean:
                        if char.isalpha():
                            col_match += char
                        elif char.isdigit():
                            row_match += char

                    if col_match and row_match:
                        col_num = column_index_from_string(col_match)
                        row_num = int(row_match)
                        cell_names_map[(row_num, col_num)] = name

            except Exception as e:
                logger.warning(f"名前付き範囲の解析に失敗: {name} - {e}")

    return cell_names_map


def find_max_enclosing_rectangles(
    worksheet, target_row, target_col, min_row, max_row, min_col, max_col
):
    """指定セルを含む最大の囲み四角形を検出"""
    rectangles = []

    # 段階的に拡張して囲み領域を検出
    for expansion in range(1, 11):  # 最大10セル範囲まで拡張
        for top in range(max(min_row, target_row - expansion), target_row + 1):
            for left in range(max(min_col, target_col - expansion), target_col + 1):
                for bottom in range(
                    target_row, min(max_row, target_row + expansion) + 1
                ):
                    for right in range(
                        target_col, min(max_col, target_col + expansion) + 1
                    ):

                        # 四角形が完全に罫線で囲まれているかチェック
                        completeness = calculate_border_completeness(
                            worksheet, top, left, bottom, right
                        )

                        if completeness >= 0.3:  # 30%以上の完全度で部分矩形と認識
                            bounds = {
                                "top": top,
                                "left": left,
                                "bottom": bottom,
                                "right": right,
                            }
                            rectangles.append(bounds)

                            logger.debug(
                                f"部分矩形検出: {bounds} 完全度={completeness:.2f}"
                            )

    return rectangles


def calculate_border_completeness(worksheet, top, left, bottom, right):
    """四角形の罫線完全度を計算（0.0-1.0）"""
    try:
        total_segments = 0
        bordered_segments = 0

        # 上辺をチェック
        for col in range(left, right + 1):
            total_segments += 1
            if has_border(worksheet, top, col, "top"):
                bordered_segments += 1

        # 下辺をチェック
        for col in range(left, right + 1):
            total_segments += 1
            if has_border(worksheet, bottom, col, "bottom"):
                bordered_segments += 1

        # 左辺をチェック
        for row in range(top, bottom + 1):
            total_segments += 1
            if has_border(worksheet, row, left, "left"):
                bordered_segments += 1

        # 右辺をチェック
        for row in range(top, bottom + 1):
            total_segments += 1
            if has_border(worksheet, row, right, "right"):
                bordered_segments += 1

        return bordered_segments / total_segments if total_segments > 0 else 0.0

    except Exception as e:
        logger.debug(f"罫線完全度計算エラー: {e}")
        return 0.0


def integrate_border_detection_with_containers(worksheet, containers=None):
    """
    罫線検出をコンテナ機能と統合
    """
    # 名前付き範囲からセル名マップを取得
    workbook = worksheet.parent
    cell_names_map = extract_cell_names_from_workbook(workbook)

    if not cell_names_map:
        logger.debug("セル名が見つかりませんでした")
        return {}

    logger.debug(f"セル名マップ: {len(cell_names_map)}個")
    for (row, col), name in cell_names_map.items():
        logger.debug(f"  {name} at ({row}, {col})")

    # 罫線領域を検出
    regions = detect_rectangular_regions(worksheet, cell_names_map)

    logger.debug(f"検出された罫線領域: {len(regions)}個")
    # 古い形式 (tuple) を新しい形式 (dict) に変換
    region_dicts = []
    for region_tuple in regions:
        top, left, bottom, right, completeness = region_tuple
        cell_names_in_region = (
            get_cell_names_in_region(cell_names_map, top, left, bottom, right)
            if cell_names_map
            else []
        )
        region_dict = {
            "bounds": {"top": top, "left": left, "bottom": bottom, "right": right},
            "area": (bottom - top + 1) * (right - left + 1),
            "completeness": completeness,
            "cell_names": cell_names_in_region,
            "repeat_info": {},  # デフォルトで空の辞書
        }
        region_dicts.append(region_dict)
        logger.debug(
            f"領域{len(region_dicts)}: {region_dict['bounds']}, 面積={region_dict['area']}, セル名数={len(region_dict['cell_names'])}"
        )

    regions = region_dicts

    # 既存のコンテナ設定と統合
    integrated_containers = {}

    if containers:
        integrated_containers.update(containers)

    # 罫線検出結果からコンテナを自動生成
    for i, region in enumerate(regions):
        if region["repeat_info"]:  # 繰り返しパターンがある場合のみ
            for pattern, cell_names in region["repeat_info"].items():
                container_name = f"auto_region_{i}_{pattern.replace('*', 'N')}"

                # 座標をExcel形式に変換
                bounds = region["bounds"]
                range_ref = (
                    f"{chr(ord('A') + bounds['left'] - 1)}{bounds['top']}:"
                    + f"{chr(ord('A') + bounds['right'] - 1)}{bounds['bottom']}"
                )

                # コンテナ定義を生成
                integrated_containers[container_name] = {
                    "range": range_ref,
                    "items": extract_field_names_from_pattern(pattern),
                    "direction": "row",  # デフォルト
                    "type": "auto_detected",
                    "source_region": region,
                }

    return integrated_containers


def extract_field_names_from_pattern(pattern):
    """パターンからフィールド名を抽出"""
    parts = pattern.split(".")
    field_names = []

    for part in parts:
        if part != "json" and part != "*" and not part.isdigit():
            field_names.append(part)

    return field_names if field_names else ["value"]


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
            if not cell_name.startswith("json."):
                raise ValueError(f"セル名は'json.'で始まる必要があります: {cell_name}")


def calculate_hierarchy_depth(cell_name):
    """数値インデックスを除外した階層深度を計算"""
    parts = cell_name.split(".")
    # 空の部分も除外し、数値でない部分のみを階層として扱う
    hierarchy_parts = [part for part in parts if part and not part.isdigit()]
    return len(hierarchy_parts) - 1  # 'json'を除く


def validate_container_config(containers):
    """コンテナ設定の妥当性を検証"""
    errors = []

    for container_name, container_def in containers.items():
        # セル名の形式チェック
        if not container_name.startswith("json."):
            errors.append(
                f"コンテナ名は'json.'で始まる必要があります: {container_name}"
            )

        # 必須項目のチェック
        has_range = "range" in container_def
        has_offset = "offset" in container_def

        if not has_range and not has_offset:
            errors.append(
                f"コンテナ{container_name}には'range'または'offset'が必要です"
            )

        if has_range and has_offset:
            errors.append(
                f"コンテナ{container_name}には'range'と'offset'の両方を指定できません"
            )

        # items の検証
        if "items" not in container_def:
            errors.append(f"コンテナ{container_name}には'items'が必要です")
        elif not isinstance(container_def["items"], list):
            errors.append(
                f"コンテナ{container_name}の'items'は配列である必要があります"
            )
        elif len(container_def["items"]) == 0:
            errors.append(f"コンテナ{container_name}の'items'は空にできません")

        # direction の検証
        if "direction" in container_def:
            direction = container_def["direction"]
            if direction not in ["row", "column"]:
                errors.append(
                    f"コンテナ{container_name}の'direction'は'row'または'column'である必要があります: {direction}"
                )

        # increment の検証
        if "increment" in container_def:
            increment = container_def["increment"]
            if not isinstance(increment, int) or increment < 1:
                errors.append(
                    f"コンテナ{container_name}の'increment'は1以上の整数である必要があります: {increment}"
                )

        # offset の検証（子コンテナの場合）
        if has_offset:
            offset = container_def["offset"]
            if not isinstance(offset, int):
                errors.append(
                    f"コンテナ{container_name}の'offset'は整数である必要があります: {offset}"
                )

        # type の検証（明示的指定の場合）
        if "type" in container_def:
            container_type = container_def["type"]
            if container_type not in ["table", "card", "tree"]:
                errors.append(
                    f"コンテナ{container_name}の'type'は'table'、'card'、または'tree'である必要があります: {container_type}"
                )

    return errors


def validate_hierarchy_consistency(containers):
    """コンテナの階層構造の整合性を検証"""
    errors = []

    # 親子関係の検証
    for container_name, container_def in containers.items():
        if "offset" in container_def:  # 子コンテナ
            parent_name = get_parent_container_name(container_name)
            if parent_name and parent_name not in containers:
                errors.append(
                    f"子コンテナ{container_name}の親コンテナ{parent_name}が見つかりません"
                )
            elif parent_name:
                parent_def = containers[parent_name]
                if "offset" in parent_def:
                    errors.append(
                        f"親コンテナ{parent_name}もoffset指定されています（range指定である必要があります）"
                    )

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
        if cell_name and cell_name.startswith("json."):
            result[cell_coord] = cell_name
    return result


def generate_cell_names_from_containers(containers, workbook):
    """
    コンテナ定義からセル名を自動生成し、実際のExcelデータから値を読み取る
    罫線検出機能と統合して自動的に繰り返し領域も検出
    """
    generated_names = {}

    # まず罫線検出による自動コンテナを生成
    if workbook.worksheets:
        worksheet = workbook.active
        auto_containers = integrate_border_detection_with_containers(
            worksheet, containers
        )

        # 手動設定と自動検出を統合
        all_containers = {}
        if containers:
            all_containers.update(containers)

        # 自動検出されたコンテナを追加（重複しない場合のみ）
        for auto_name, auto_def in auto_containers.items():
            if auto_name not in all_containers:
                all_containers[auto_name] = auto_def
                logger.debug(f"自動検出コンテナを追加: {auto_name}")

        containers = all_containers

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
            process_root_container(
                container_name, container_def, workbook, generated_names
            )
        elif has_offset:
            # 子コンテナの処理
            process_child_container(
                container_name, container_def, workbook, generated_names
            )
        else:
            logger.warning(
                f"コンテナ{container_name}にrangeまたはoffsetが指定されていません"
            )

    logger.debug(f"生成されたセル名と値: {generated_names}")
    return generated_names


def sort_containers_by_hierarchy(containers):
    """コンテナを階層の深さでソート（浅い順）"""
    container_items = list(containers.items())
    return sorted(container_items, key=lambda x: calculate_hierarchy_depth(x[0]))


def calculate_hierarchy_depth(cell_name):
    """セル名から階層の深さを計算（数値インデックスを除外）"""
    parts = cell_name.split(".")
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

    logger.debug(
        f"コンテナ設定: direction={direction}, increment={increment}, items={items}, labels={labels}"
    )

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
        container_type = detect_container_type(
            actual_range, workbook, increment, container_def
        )
        logger.debug(f"検出されたコンテナタイプ: {container_type}")

        # タイプ別処理
        if container_type == "table":
            process_table_container(
                container_name, container_def, actual_range, workbook, generated_names
            )
        elif container_type == "card":
            process_card_container(
                container_name, container_def, actual_range, workbook, generated_names
            )
        elif container_type == "tree":
            process_tree_container(
                container_name, container_def, actual_range, workbook, generated_names
            )
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
        process_child_instance(
            container_name,
            container_def,
            parent_instance_idx,
            workbook,
            generated_names,
        )


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
                logger.debug(
                    f"範囲形状によりカード型判定: {width}x{height}, ratio={aspect_ratio:.2f}"
                )
                return "card"

            # 縦長の場合はテーブル型、横長の場合もテーブル型
            logger.debug(f"範囲形状によりテーブル型判定: {width}x{height}")
    except Exception as e:
        logger.warning(f"範囲解析エラー: {e}")

    # デフォルトはテーブル型
    return "table"


def process_table_container(
    container_name, container_def, range_spec, workbook, generated_names
):
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
            logger.debug(
                f"基準位置取得（方法1）: {item} -> {position} (セル名: {base_cell_name})"
            )
        else:
            # 方法2: 最初のインスタンスのセル名（配列指定）
            array_cell_name = f"json.{base_container_name}.{item}.1"
            position = get_cell_position_from_name(array_cell_name, workbook)
            if position:
                logger.debug(
                    f"基準位置取得（方法2）: {item} -> {position} (セル名: {array_cell_name})"
                )
            else:
                # 方法3: 直接のセル名（単一セル）
                direct_cell_name = f"json.{base_container_name}.{item}"
                position = get_cell_position_from_name(direct_cell_name, workbook)
                if position:
                    logger.debug(
                        f"基準位置取得（方法3）: {item} -> {position} (セル名: {direct_cell_name})"
                    )
                else:
                    logger.warning(
                        f"基準セル名が見つかりません: {item} (試行: {base_cell_name}, {array_cell_name}, {direct_cell_name})"
                    )

        if position:
            base_positions[item] = position

    if not base_positions:
        logger.error(
            f"基準位置が見つからないため、コンテナ {container_name} をスキップ"
        )
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

            logger.debug(
                f"セル値読み取り {cell_name}: row={target_position[1]}, col={target_position[0]}, value={cell_value}"
            )

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


def process_card_container(
    container_name, container_def, range_spec, workbook, generated_names
):
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
                logger.debug(
                    f"既存セル名から値取得: {existing_cell_name} -> {cell_value}"
                )
            else:
                # 計算された位置から値を読み取り
                target_position = calculate_target_position(
                    base_positions[item], direction, card_idx, increment
                )
                cell_value = read_cell_value(target_position, ws)
                logger.debug(
                    f"計算位置から値取得: {existing_cell_name} -> {cell_value}"
                )

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
            parts = name[len(prefix) :].split(".")
            if parts and parts[0].isdigit():
                card_indices.add(int(parts[0]))

    return max(card_indices) if card_indices else 0


def process_tree_container(
    container_name, container_def, range_spec, workbook, generated_names
):
    """ツリー型コンテナの処理（階層構造対応）"""
    logger.debug(f"ツリー型コンテナ処理: {container_name}")

    # ツリー型は基本的にテーブル型と同じだが、階層構造を考慮
    process_table_container(
        container_name, container_def, range_spec, workbook, generated_names
    )


def get_cell_position_from_name(cell_name, workbook):
    """セル名から座標位置を取得"""
    if cell_name in workbook.defined_names:
        defined_name = workbook.defined_names[cell_name]
        for sheet_name, coord in defined_name.destinations:
            match = re.match(r"^\$?([A-Z]+)\$?(\d+)$", coord)
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
            if value_stripped.isdigit() or (
                value_stripped.startswith("-") and value_stripped[1:].isdigit()
            ):
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
        if value_stripped.lower() in ["true", "yes", "on", "1", "はい", "真", "オン"]:
            return True
        elif value_stripped.lower() in [
            "false",
            "no",
            "off",
            "0",
            "いいえ",
            "偽",
            "オフ",
        ]:
            return False

        # 日付の判定（簡易版）
        try:
            import datetime

            # ISO形式の日付
            if re.match(r"^\d{4}-\d{2}-\d{2}$", value_stripped):
                return datetime.datetime.strptime(value_stripped, "%Y-%m-%d").date()
            # ISO形式の日時
            if re.match(r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}", value_stripped):
                return datetime.datetime.fromisoformat(
                    value_stripped.replace("Z", "+00:00")
                )
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
            cell_value = cell_value.replace("\r\n", "\n").replace("\r", "\n")

        # 全角・半角の正規化
        if normalize_options.get("normalize_width", False):
            import unicodedata

            cell_value = unicodedata.normalize("NFKC", cell_value)

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
        target_position = calculate_target_position(
            base_position, direction, card_idx, increment
        )
        cell_value = read_cell_value(target_position, worksheet)

        if cell_value:  # 値がある場合はカードが存在
            card_count = card_idx
        else:
            break  # 空の場合は終了

    return card_count


def get_parent_container_name(container_name):
    """子コンテナから親コンテナ名を取得"""
    parts = container_name.split(".")
    if len(parts) >= 3:
        # 数値インデックスを除去して親を特定
        parent_parts = []
        for part in parts[:-1]:  # 最後の部分を除く
            if not part.isdigit():
                parent_parts.append(part)
        return ".".join(parent_parts)
    return None


def find_parent_instances(parent_container_name, generated_names):
    """生成済みセル名から親インスタンスのインデックスを取得"""
    parent_instances = set()
    prefix = parent_container_name + "."

    for cell_name in generated_names.keys():
        if cell_name.startswith(prefix):
            parts = cell_name.replace(prefix, "").split(".")
            if parts and parts[0].isdigit():
                parent_instances.add(int(parts[0]))

    return sorted(parent_instances)


def process_child_instance(
    container_name, container_def, parent_instance_idx, workbook, generated_names
):
    """子コンテナの特定インスタンスを処理"""
    logger.debug(
        f"子インスタンス処理: {container_name}, 親インスタンス: {parent_instance_idx}"
    )

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
            position = estimate_position_from_parent(
                parent_cell_name, workbook, generated_names
            )
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
        container_name,
        parent_instance_idx,
        parent_base_positions,
        direction,
        increment,
        offset,
        workbook,
    )

    logger.debug(
        f"検出された子インスタンス数: {len(child_instances)} (親{parent_instance_idx})"
    )

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

            logger.debug(
                f"子セル値読み取り {cell_name}: row={child_position[1]}, col={child_position[0]}, value={cell_value}"
            )

            if cell_value:
                all_empty = False

            instance_values[cell_name] = cell_value

        # 有効なデータがある場合のみ追加
        if not all_empty:
            generated_names.update(instance_values)
            logger.debug(f"子インスタンス{child_idx}: 有効なデータとして追加")


def detect_child_instances_from_data(
    container_name,
    parent_instance_idx,
    parent_base_positions,
    direction,
    increment,
    offset,
    workbook,
):
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
    try:
        parser = create_argument_parser()
        args = parser.parse_args()
    except SystemExit:
        logger.error("引数の解析に失敗しました")
        return 1

    try:
        config = create_config_from_args(args)
        converter = Xlsx2JsonConverter(config)
        return converter.process_files(config.input_files)
    except (ConfigurationError, FileProcessingError) as e:
        logger.error(f"エラー: {e}")
        return 1
    except Exception as e:
        logger.error(f"予期しないエラー: {e}")
        return 1


def create_argument_parser() -> argparse.ArgumentParser:
    """コマンドライン引数パーサーを作成"""
    parser = argparse.ArgumentParser(description="Excel の名前付き範囲を JSON に変換")
    parser.add_argument("input_files", nargs="*", help="入力 Excel ファイル")
    parser.add_argument("--config", type=Path, help="設定ファイル")
    parser.add_argument("--output-dir", "-o", type=Path, help="出力ディレクトリ")
    parser.add_argument("--prefix", "-p", default="json", help="プレフィックス")
    parser.add_argument("--schema", "-s", type=Path, help="JSONスキーマファイル")
    parser.add_argument(
        "--output-format",
        "-f",
        choices=["json", "yaml"],
        default="json",
        help="出力フォーマット (json/yaml, デフォルト: json)",
    )
    parser.add_argument(
        "--log-level",
        choices=["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"],
        help="ログレベル",
    )
    parser.add_argument("--keep-empty", action="store_true", help="空の値を保持")
    parser.add_argument("--trim", action="store_true", help="文字列の前後の空白を削除")
    parser.add_argument("--container", action="append", help="コンテナ定義 (JSON)")
    parser.add_argument(
        "--transform", action="append", help="変換ルール (複数指定で連続適用)"
    )
    return parser


def create_config_from_args(args) -> ProcessingConfig:
    """コマンドライン引数から設定を作成"""
    # 設定ファイルの読み込み
    config_data = {}
    if args.config:
        try:
            with args.config.open("r", encoding="utf-8") as f:
                config_data = json.load(f)
        except (json.JSONDecodeError, FileNotFoundError, IOError) as e:
            raise ConfigurationError(f"設定ファイルの読み込みに失敗: {e}")

    # コマンドライン引数で設定を上書き
    if args.input_files:
        config_data["input-files"] = args.input_files
    if args.output_dir:
        config_data["output-dir"] = args.output_dir
    # prefixは設定ファイルにない場合のみデフォルト値を使用
    if args.prefix and "prefix" not in config_data:
        config_data["prefix"] = args.prefix
    if args.schema:
        config_data["schema"] = args.schema
    if args.keep_empty:
        config_data["keep-empty"] = True
    if args.trim:
        config_data["trim"] = True
    if args.log_level:
        config_data["log-level"] = args.log_level
    if args.container:
        validate_cli_containers(args.container)
        cli_containers = parse_container_args(args.container)
        config_containers = config_data.get("containers", {})
        config_data["containers"] = {**config_containers, **cli_containers}
    if args.transform:
        config_transforms = config_data.get("transform", [])
        config_data["transform"] = config_transforms + args.transform
    if args.output_format:
        config_data["output-format"] = args.output_format

    # ログレベル設定
    log_level = config_data.get("log-level", "INFO")
    logging.basicConfig(
        level=getattr(logging, log_level), format="%(levelname)s: %(message)s"
    )

    # 必須パラメータのチェック
    if not config_data.get("input-files"):
        raise ConfigurationError("入力ファイルが指定されていません")

    # スキーマの読み込み
    schema = None
    if config_data.get("schema"):
        schema_path = Path(config_data["schema"])
        schema = SchemaLoader.load_schema(schema_path)

    return ProcessingConfig(
        input_files=config_data.get("input-files", []),
        prefix=config_data.get("prefix", "json"),
        trim=config_data.get("trim", False),
        keep_empty=config_data.get("keep-empty", False),
        output_dir=config_data.get("output-dir"),
        output_format=config_data.get("output-format", "json"),
        schema=schema,
        containers=config_data.get("containers", {}),
        transform_rules=config_data.get("transform", []),
    )


# =============================================================================
# Wildcard and 2D Array Transform Extensions
# =============================================================================


def apply_wildcard_transforms(
    data: dict, transform_rules: Dict[str, List[ArrayTransformRule]], prefix: str
) -> dict:
    """
    ワイルドカード変換ルールを適用

    Args:
        data: 変換対象のデータ
        transform_rules: 変換ルール（ワイルドカードを含む）
        prefix: プレフィックス

    Returns:
        変換後のデータ
    """
    if not transform_rules:
        return data

    def match_wildcard_path(pattern: str, actual_path: str) -> bool:
        """ワイルドカードパターンマッチング"""
        pattern_parts = pattern.split(".")
        actual_parts = actual_path.split(".")

        if len(pattern_parts) != len(actual_parts):
            return False

        for pattern_part, actual_part in zip(pattern_parts, actual_parts):
            if pattern_part == "*":
                continue
            elif pattern_part != actual_part:
                return False
        return True

    def get_nested_value(obj: dict, path: str):
        """ネストされた値を取得"""
        parts = path.split(".")
        current = obj
        for part in parts:
            if isinstance(current, dict) and part in current:
                current = current[part]
            else:
                return None
        return current

    def set_nested_value(obj: dict, path: str, value):
        """ネストされた値を設定"""
        parts = path.split(".")
        current = obj
        for part in parts[:-1]:
            if part not in current:
                current[part] = {}
            current = current[part]
        current[parts[-1]] = value

    def find_matching_paths(
        obj: dict, pattern: str, current_path: str = ""
    ) -> List[str]:
        """パターンにマッチするパスを検索"""
        matches = []

        if isinstance(obj, dict):
            for key, value in obj.items():
                new_path = f"{current_path}.{key}" if current_path else key

                if match_wildcard_path(pattern, new_path):
                    matches.append(new_path)

                # 再帰的に探索
                if isinstance(value, (dict, list)):
                    matches.extend(find_matching_paths(value, pattern, new_path))

        return matches

    # 各変換ルールを適用
    for pattern, rule_list in transform_rules.items():
        if "*" not in pattern:
            continue  # ワイルドカードでないルールはスキップ

        matching_paths = find_matching_paths(data, pattern)

        for path in matching_paths:
            current_value = get_nested_value(data, path)
            if current_value is not None:
                try:
                    # 連続適用: 複数の変換ルールを順次適用
                    for rule in rule_list:
                        # 変換実行
                        transformed_value = rule.transform(current_value)
                        current_value = transformed_value

                    # 結果を設定
                    if isinstance(current_value, dict):
                        # dict型の場合はキーによる階層指定を処理
                        for key, val in current_value.items():
                            if key.startswith(prefix):
                                # 絶対指定（json.で始まる場合）
                                abs_path = (
                                    key[len(prefix) :]
                                    if key.startswith(prefix + ".")
                                    else key[len(prefix) :]
                                )
                                if abs_path:
                                    set_nested_value(data, abs_path, val)
                            else:
                                # 相対指定
                                rel_path = f"{path}.{key}"
                                set_nested_value(data, rel_path, val)
                    else:
                        set_nested_value(data, path, current_value)

                except Exception as e:
                    logger.error(
                        f"ワイルドカード変換エラー: パス={path}, ルール={rule_list}, エラー={e}"
                    )

    return data


def apply_2d_array_transforms(
    data: dict,
    transform_rules: Dict[str, List[ArrayTransformRule]],
    workbook,
    prefix: str,
) -> dict:
    """
    二次元配列変換ルールを適用

    Args:
        data: 変換対象のデータ
        transform_rules: 変換ルール
        workbook: Excelワークブック
        prefix: プレフィックス

    Returns:
        変換後のデータ
    """
    if not transform_rules:
        return data

    for path, rule_list in transform_rules.items():
        for rule in rule_list:
            if rule.transform_type != "range":
                continue  # range変換以外はスキップ

            try:
                # 範囲データを取得
                range_data = rule._get_range_data(workbook)

                if range_data is not None:
                    # 変換実行
                    transformed_value = rule._transform_with_range(range_data)

                    # 結果をデータに設定
                    path_parts = path.split(".")
                    current = data
                    for part in path_parts[:-1]:
                        if part not in current:
                            current[part] = {}
                        current = current[part]

                    if isinstance(transformed_value, dict):
                        # dict型の場合はキーによる階層指定を処理
                        for key, val in transformed_value.items():
                            if key.startswith(prefix):
                                # 絶対指定
                                abs_path = (
                                    key[len(prefix) :]
                                    if key.startswith(prefix + ".")
                                    else key[len(prefix) :]
                                )
                                if abs_path:
                                    abs_parts = abs_path.split(".")
                                    abs_current = data
                                    for abs_part in abs_parts[:-1]:
                                        if abs_part not in abs_current:
                                            abs_current[abs_part] = {}
                                        abs_current = abs_current[abs_part]
                                    abs_current[abs_parts[-1]] = val
                            else:
                                # 相対指定
                                current[key] = val
                    else:
                        current[path_parts[-1]] = transformed_value

            except Exception as e:
                logger.error(
                    f"二次元配列変換エラー: パス={path}, ルール={rule}, エラー={e}"
                )

    return data


def extract_column(data, index=0):
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


def table_to_dict(data):
    """テーブルデータを辞書に変換（テスト互換性のため）"""
    if not isinstance(data, list) or len(data) < 2:
        return {}

    headers = data[0] if isinstance(data[0], list) else []
    result = {}

    for i, row in enumerate(data[1:], 1):
        if isinstance(row, list):
            row_data = {}
            for j, header in enumerate(headers):
                if j < len(row):
                    row_data[header] = row[j]
            result[f"row_{i}"] = row_data

    return result


if __name__ == "__main__":
    exit(main())
