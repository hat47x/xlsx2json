"""
xlsx2json - Excel の名前付き範囲を JSON に変換するツール
"""

import argparse
import json
import re
import logging
import shlex
import subprocess
import importlib.util
from pathlib import Path
from typing import Any, Dict, List, Optional, Union, Callable

from openpyxl import load_workbook
from jsonschema import Draft7Validator


# グローバル変数
_global_trim = False
_global_schema = None

logger = logging.getLogger("xlsx2json")


def load_schema(schema_path: Optional[Path]) -> Optional[Dict[str, Any]]:
    """
    指定されたパスから JSON スキーマを読み込む。
    スキーマが指定されていない場合は None を返す。
    """
    if not schema_path:
        return None

    with schema_path.open("r", encoding="utf-8") as f:
        return json.load(f)


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

    logger.info(f"Validation errors logged: {log_file}")


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
    rules = {}
    for rule in array_split_rules:
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


class ArrayTransformRule:
    """配列変換ルールを表すクラス"""

    def __init__(self, path: str, transform_type: str, transform_spec: str):
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


def parse_named_ranges_with_prefix(
    xlsx_path: Path,
    prefix: str,
    array_split_rules: Optional[Dict[str, List[str]]] = None,
    array_transform_rules: Optional[Dict[str, ArrayTransformRule]] = None,
) -> Dict[str, Any]:
    """
    Excel 名前付き範囲(prefix) を解析してネスト dict/list を返す。
    prefixはデフォルトで"json"。
    array_split_rules: 配列化設定の辞書 {path: [delimiter1, delimiter2, ...]}
    array_transform_rules: 配列変換設定の辞書 {path: ArrayTransformRule}
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

    for name, defined_name in wb.defined_names.items():
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
            elif orig_key_path in array_transform_rules:
                return array_transform_rules[orig_key_path]
            # 配列要素の場合は親キーでも判定
            if len(path_keys) > 1 and ".".join(path_keys[:-1]) in array_transform_rules:
                return array_transform_rules[".".join(path_keys[:-1])]
            if (
                len(original_path_keys) > 1
                and ".".join(original_path_keys[:-1]) in array_transform_rules
            ):
                return array_transform_rules[".".join(original_path_keys[:-1])]
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
        insert_json_path(result, path_keys, value, ".".join(path_keys))

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
                if entry.suffix.lower() == ".xlsx":
                    files.append(entry)
        elif p_path.is_file() and p_path.suffix.lower() == ".xlsx":
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


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Excel 名前付き範囲(json.*) -> JSON ファイル出力",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
変換ルールの例:
  配列化:
    1次元配列化: --transform "json.tags=split:,"  (カンマで分割)
    2次元配列化: --transform "json.matrix=split:,|\n"  (カンマ→改行で分割)
    3次元配列化: --transform "json.cube=split:,|\\||\n"  (カンマ→パイプ→改行で分割)
  Python関数:
    モジュール: --transform "json.tags=function:mymodule:split_func"
    ファイル: --transform "json.tags=function:/path/to/script.py:split_func"
  外部コマンド: --transform "json.lines=command:sort -u"

  外部コマンドは標準入力から値を受け取り、標準出力に結果を返します。
        """,
    )
    parser.add_argument(
        "inputs", nargs="*", help="変換対象のファイルまたはフォルダ (.xlsx)"
    )
    parser.add_argument(
        "--output-dir",
        "-o",
        type=Path,
        help="一括出力先フォルダ (省略時は各入力ファイル隣の output-json)",
    )
    parser.add_argument(
        "--schema",
        "-s",
        type=Path,
        help="JSON Schema ファイル (バリデーションとソート用)",
    )
    parser.add_argument(
        "--transform",
        action="append",
        default=[],
        help="変換ルール",
    )
    parser.add_argument(
        "--config",
        type=Path,
        help="設定ファイル",
    )
    parser.add_argument(
        "--trim",
        action="store_true",
        help="配列化時に各要素をトリムする",
    )
    parser.add_argument(
        "--keep-empty",
        action="store_true",
        help="空のセル値も JSON に含める (デフォルトでは空値を除去)",
    )
    parser.add_argument(
        "--log-level",
        choices=["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"],
        help="ログレベルを指定 (デフォルト: INFO)",
    )
    parser.add_argument(
        "--prefix",
        type=str,
        default="json",
        help="Excelセル名のプレフィックス",
    )
    args = parser.parse_args()

    # 設定ファイル読み込み（コマンドライン引数が優先）
    config = {}
    if getattr(args, "config", None):
        try:
            with args.config.open("r", encoding="utf-8") as f:
                config = json.load(f)
        except Exception as e:
            logger.error(f"設定ファイルの読込に失敗しました: {e}")
            config = {}

    def get_opt(key, default=None):
        # コマンドライン→設定ファイル→デフォルト値の順で返す
        val = getattr(args, key, None)
        # inputsの場合は空のリストでも設定ファイルを確認
        if key == "inputs":
            if val:  # 空でないリストの場合
                return val
        else:
            if val is not None and val != []:
                return val
        if key in config:
            return config[key]
        return default

    # 入力ファイル/フォルダ
    inputs = get_opt("inputs")
    if not inputs:
        logger.error(
            "入力ファイル/フォルダが指定されていません。コマンドライン引数または --config で指定してください。"
        )
        return
    if isinstance(inputs, str):
        inputs = [inputs]

    # 出力先
    output_dir = get_opt("output_dir")
    # スキーマ
    schema_path = get_opt("schema")
    if schema_path:
        schema_path = Path(schema_path)
    # 変換ルール
    transform_list = list(args.transform) if args.transform else []
    if "transform" in config and isinstance(config["transform"], list):
        transform_list.extend(
            [t for t in config["transform"] if t not in transform_list]
        )
    # プレフィックス
    prefix = get_opt("prefix", "json")
    # 空値保持
    keep_empty = get_opt("keep_empty", False)
    # ログレベル
    log_level = get_opt("log_level", "INFO")

    # 既存のハンドラをクリア（basicConfigが効かない問題対策）
    if logging.root.handlers:
        logging.root.handlers.clear()

    # ログ設定
    logging.basicConfig(
        level=getattr(logging, log_level), format="[%(levelname)s] %(message)s"
    )
    logger.setLevel(getattr(logging, log_level))

    logger.debug(f"入力ファイルリスト: {inputs}")
    xlsx_files = collect_xlsx_files(inputs)
    logger.debug(f"収集されたxlsxファイル: {xlsx_files}")
    if not xlsx_files:
        logger.error("対象の .xlsx ファイルが見つかりませんでした。")
        return

    schema = load_schema(schema_path) if schema_path else None
    logger.debug(f"ロードしたスキーマ: {schema is not None}")
    validator = Draft7Validator(schema) if schema else None
    suppress_empty = not keep_empty

    transform_rules = parse_array_transform_rules(transform_list, prefix, schema)
    logger.debug(f"変換設定: {transform_rules}")

    global _global_trim
    _global_trim = get_opt("trim", False)

    global _global_schema
    _global_schema = schema

    for xlsx_path in xlsx_files:
        logger.debug(f"処理開始: {xlsx_path}")
        try:
            data = parse_named_ranges_with_prefix(
                xlsx_path,
                prefix=prefix,
                array_split_rules=None,
                array_transform_rules=transform_rules,
            )
            logger.debug(f"parse_named_ranges_with_prefix結果: {data}")
        except Exception as e:
            logger.exception(f"例外が発生しました: {e}")
            data = None

        out_dir = output_dir if output_dir else xlsx_path.parent / "output-json"
        if isinstance(out_dir, str):
            out_dir = Path(out_dir)
        out_file = out_dir / f"{xlsx_path.stem}.json"
        logger.debug(f"出力先: {out_file}")
        if data is None:
            logger.error(f"データがNoneのため、出力をスキップします: {out_file}")
            continue
        if isinstance(data, dict) and not data:
            logger.warning(
                f"データが空のdictです。出力ファイルは空になります: {out_file}"
            )
        write_json(data, out_file, schema, validator, suppress_empty)


if __name__ == "__main__":
    main()
