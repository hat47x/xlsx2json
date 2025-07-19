# xlsx2json

Excel の名前付き範囲（`json.*`）を読み込んで JSON 形式に変換し、ファイル出力するコマンドラインツールです。

---

## 主な機能

### 🔄 名前付き範囲の自動解析
セル名に `json.` プレフィックスを付けるだけで、自動的に JSON 形式に変換します。

### 📋 JSON Schema サポート
`--schema` オプションで JSON Schema を指定することで、データのバリデーションとキー順序の整理が可能です。

### 🔀 記号ワイルドカード対応
Excelの「名前の定義」で使えない記号も、アンダーバー（`_`）を1文字ワイルドカードとしてJSON Schemaの項目名にユニークにマッチする場合は置き換えて出力されます。

### 🛠️ 多次元配列の柔軟な分割
複数の区切り文字を使って1次元から多次元配列まで対応。エスケープ処理により、パイプ文字（`|`）も区切り文字として利用可能です。

### 📊 エラーログ出力
スキーマ違反があった場合は `<basename>.error.log` に詳細なエラー情報を出力します。

### 🚫 空値の自動除去
空のセル値はデフォルトで JSON から除外され、必要に応じて `--keep-empty` オプションで保持可能です。

### 📁 柔軟な出力先設定
個別フォルダへの出力、または `--output-dir` による一括出力先指定が可能です。

---

## インストール

```bash
git clone https://github.com/hat47x/xlsx2json.git
cd xlsx2json
pip install openpyxl jsonschema
```

---

## 使用方法

### 基本構文

```bash
python xlsx2json.py [INPUT1 ...] [OPTIONS]
```

### オプション一覧

| オプション | 説明 |
|:----------|:-----|
| `INPUT1 ...` | 変換対象のファイルまたはフォルダ（.xlsx）。フォルダ指定時は直下の `.xlsx` ファイルを対象とします。 |
| `-o, --output-dir` | 一括出力先フォルダを指定。省略時は各入力ファイルと同じディレクトリの `output-json/` に出力されます。 |
| `-s, --schema` | JSON Schema ファイルを指定。バリデーションとキー順序の整理に使用されます。 |
| `--transform RULE` | 変換設定。指定した名前付き範囲の値に対し、split（区切り文字による配列化）、function（Python関数）、command（外部コマンド）による変換を適用（複数指定可）。 |
| `--config FILE` | 設定ファイル（JSON形式）から全オプションを一括指定。コマンドライン引数が優先されます。 |
| `--keep-empty` | 空のセル値も JSON に含めます（デフォルトでは空値を除去）。 |
| `--prefix PREFIX` | Excel 名前付き範囲のプレフィックスを指定（デフォルト: `json`）。 |
| `--log-level LEVEL` | ログレベルを指定（`DEBUG`/`INFO`/`WARNING`/`ERROR`/`CRITICAL`、デフォルト: `INFO`）。 |

---

## 設定ファイルによる一括指定

`--config` オプションで、変換ルール以外も含めた全オプションをJSON形式で一括指定できます。

### 設定ファイル例（config.json）
```json
{
  "inputs": ["sample.xlsx"],
  "output_dir": "output-json",
  "schema": "sample-schema.json",
  "transform": [
    "json.tags=split:,",
    "json.matrix=split:,|\n"
  ],
  "prefix": "json",
  "keep_empty": false,
  "log_level": "INFO"
}
```

### 実行例
```bash
python xlsx2json.py --config config.json
```

コマンドライン引数で指定した値は、設定ファイルより優先されます。

---

## 処理の流れ

1. **ファイル収集**: 入力パスから `.xlsx` ファイルを収集
2. **データ読み込み**: 名前付き範囲（`json.*`）を読み込み、ネスト構造の辞書・リストを生成
3. **変換処理**: `--transform` オプションで指定したルールに従い、値の配列化・関数変換・コマンド変換等を適用
4. **空値処理**: （デフォルト）空の値を除去（`--keep-empty` オプションで保持可能）
5. **バリデーション**: （オプション）Schema バリデーションとエラーログ出力
6. **キー整理**: （オプション）Schema の `properties` 順にキーを整理
7. **ファイル出力**: JSON ファイルとして書き出し

---

## 名前付き範囲の命名規則

### 基本ルール

#### 1. プレフィックス
必ず `json.` で始めてください。ツールはこのプレフィックスを検知して処理対象とします。

#### 2. 階層表現
プレフィックスの後にドット `.` 区切りでキー階層を指定します。
```
json.customer.name → { "customer": { "name": ... } }
```

#### 3. 配列指定
数字のみのセグメントは 1 始まりのインデックスとして扱い、配列要素を作成します。
```
json.items.1 → items[0] に値を設定
json.parent.2.1 → { "parent": [[...],[value,...]] }
```

### キー名の制約

#### 対応文字
- Unicode 文字をサポート（日本語の漢字・ひらがな・カタカナ、英数字、アンダースコア `_` 等）
- 大文字・小文字は区別されます

#### 禁止文字
- ドット `.` とスペースは区切り文字のため、キー名に含めないでください

#### 文字数制限
- Excel の仕様上、名前の長さは 255 文字以内です

#### 全角記号の注意点
全角記号をキー名に使用する際は、実際に使用できる文字種であるかを確認してください。
- ✅ 使用可能: "。", "・", "＿"
- ❌ 使用不可: "、", "！", "～", "／", "（", "）" など

### 重要な注意事項

- **重複禁止**: 同じシート内、またはブック全体で同名の名前付き範囲を複数定義しないでください
- **命名の一貫性**: チーム開発では命名規則を統一することをお勧めします

### サンプル例

```
json.user.id              → { "user": { "id": ... } }
json.order.details.price  → { "order": { "details": { "price": ... } } }
json.tags.1               → { "tags": [value1, ...] }
json.tags.2               → { "tags": [value1, value2, ...] }
json.orders.1.items.3.name → { "orders": [{ "items": [null, null, { "name": ... }] }] }
```

---


## 変換ルール指定（--transform）

Excelの名前付き範囲から抽出した値に対して、柔軟な変換ルールを指定できます。

### 配列化（split）

区切り文字で値を配列化したい場合は、`split:` 変換タイプを使います。

#### 実行例

```bash
# カンマ区切りで配列化
python xlsx2json.py sample.xlsx --transform "json.tags=split:,"
# → ["apple", "banana", "orange"]

# 改行区切りで配列化
python xlsx2json.py sample.xlsx --transform "json.parent.1=split:\n"
# → ["A", "B", "C"]

# 多次元配列（セミコロン→カンマ）
python xlsx2json.py sample.xlsx --transform "json.matrix=split:;|,"
# → [["A", "B"], ["C", "D"]]

# パイプ文字を区切り文字として使用
python xlsx2json.py sample.xlsx --transform "json.data=split:\|"
# → ["A", "B", "C"]

# 3次元配列（セミコロン→パイプ→カンマ）
python xlsx2json.py sample.xlsx --transform "json.cube=split:;|\||,"
# → [[["A", "B"], ["C", "D"]], [["E", "F"], ["G", "H"]]]
```

区切り文字は `|` で区切って多次元配列に対応します。
パイプ文字自体を区切り文字にする場合は `\|` でエスケープします。
改行（`\n`）、タブ（`\t`）、復帰（`\r`）も利用可能です。

### Python関数による変換

```bash
python xlsx2json.py sample.xlsx --transform "json.tags=function:mymodule:split_func"
python xlsx2json.py sample.xlsx --transform "json.tags=function:/path/to/script.py:split_func"
# → Python関数で値を変換
```

### 外部コマンドによる変換

```bash
python xlsx2json.py sample.xlsx --transform "json.lines=command:sort -u"
# → sortコマンドで値を変換
```

## 記号ワイルドカード対応

### 概要

Excelの「名前の定義」で使えない記号や一部の日本語記号をJSONのキー名に使いたい場合、アンダーバー（`_`）を1文字ワイルドカードとして利用できます。

### 動作原理

1. Excelの名前付き範囲で `json.` 以降のキー名にアンダーバー（`_`）を使用
2. その位置が「任意の1文字」としてSchemaのプロパティ名にマッチします
3. マッチするSchemaのプロパティが1つだけの場合、そのプロパティ名に自動で置換されます
4. 複数マッチまたはマッチしない場合は、元のキー名のまま出力されます

### 使用例

#### JSON Schema例
```json
{
  "type": "object",
  "properties": {
    "user_name": {"type": "string"},
    "user／group": {"type": "string"},
    "user！": {"type": "string"},
    "user？": {"type": "string"}
  }
}
```

#### Excelでの名前付き範囲の定義

| Excelの名前定義 | 対応するJSONキー | 備考 |
|---|---|---|
| `json.user_name` | `user_name` | そのまま一致 |
| `json.user_group` | `user／group` | ワイルドカードによるマッチング |
| `json.user_` | `user！, user？` | 複数マッチ時は置換されない |

### 注意点

- アンダーバーの数だけ1文字ずつ任意の文字にマッチします
- マッチするSchemaプロパティが複数ある場合は、Excelで指定した名前がそのまま使われます
- ワイルドカード置換は各階層ごとに適用されます
- 置換の有無やマッチ状況は、`--log-level DEBUG` で確認できます

### 実行例

```bash
# 基本的な使用
python xlsx2json.py sample.xlsx --schema sample-schema.json

# ログでマッチ状況を確認
python xlsx2json.py sample.xlsx --schema sample-schema.json --log-level DEBUG
```

---

## JSON Schema のサンプル

```json
{
  "$schema": "http://json-schema.org/draft-07/schema#",
  "title": "Sample",
  "type": "object",
  "properties": {
    "parent": {
      "type": "array",
      "description": "4次元配列(縦×横×セル内縦×横)",
      "items": {
        "type": "array",
        "description": "3次元配列(横×セル内縦×横)",
        "items": {
          "type": "array",
          "description": "2次元配列(セル内縦×横)",
          "items": {
            "type": "array",
            "description": "1次元配列(セル内横)",
            "items": {
              "type": "string",
              "description": "文字列"
            }
          }
        }
      }
    },
    "customer": {
      "type": "object",
      "description": "顧客",
      "properties": {
        "name": {
          "type": "string",
          "description": "氏名"
        },
        "address": {
          "type": "string",
          "description": "住所"
        }
      },
      "required": ["name", "address"],
      "additionalProperties": false
    },
    "日本語！": {
      "type": "object",
      "properties": {
        "記号！": {
          "type": "string",
          "description": "Excelの「名前の定義」で使用できない記号をパスに含むケース"
        },
        "（記／号）～": {
          "type": "string",
          "description": "Excelの「名前の定義」で使用できない記号をパスに含むケース２"
        }
      }
    }
  },
  "required": ["customer", "parent", "日本語！"],
  "additionalProperties": false
}
```

## 変換ルール指定（--transform）

Excelの名前付き範囲から抽出した値に対して、柔軟な変換ルールを指定できます。

### 配列化（split）

区切り文字で値を配列化したい場合は、`split:` 変換タイプを使います。

例：

```
--transform "json.tags=split:,"
--transform "json.matrix=split:;|,"
--transform "json.data=split:\||,"
```

区切り文字は `|` で区切って多次元配列に対応します。
パイプ文字自体を区切り文字にする場合は `\|` でエスケープします。

### Python関数による変換

```
--transform "json.tags=function:mymodule:split_func"
--transform "json.tags=function:/path/to/script.py:split_func"
```

### 外部コマンドによる変換

```
--transform "json.lines=command:sort -u"
```

### 実際のワークフロー例

```bash
# 開発環境での検証
python xlsx2json.py sample.xlsx --schema validation.json --log-level DEBUG

# 本番環境での一括処理
python xlsx2json.py ./production_data/ \
  --output-dir ./json_output/ \
  --schema production_schema.json \
  --transform "json.categories=split:," \
  --transform "json.nested_data=split:;|,"
```

## ライセンス

MIT