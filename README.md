# xlsx2json

Excelのセル名を用いて事前に定義されたデータ構造にもとづき、JSON形式に変換するツールです。

## 特徴

Excelの「名前付き範囲」（以降「セル名」）を用いて事前に定義されたデータ構造にもとづき、JSON形式に変換するツールです。Excelファイルの様式が入り組んでいても、 **セル名さえつけてしまえば様式ごとの変換ロジックが不要** な点が本ツールの特徴です。

一方デメリットとしては、事前にセル名をつける作業が必要となります。特に、様式にもとづき自動で繰り返し要素を識別することができないため、例えば5列×100行の入力項目が存在する場合、それらすべてにセル名をつける必要があります。また行数を増やしたい場合は追加でセル名をつける必要があります。
※作業負荷を軽減するため「セル名インポート・エクスポート用マクロ.xlsm」を同梱しています

上記のトレードオフを念頭に、本ツールの利用をご検討ください。

### 🔄 セル名でJSON出力時の階層構造を指定
セル名にドット `.` 区切りのキー階層を記載することで、自動的にネストした JSON 構造に変換します。複雑な階層データも直感的に定義可能です。
セル名の先頭には `json.` を付加してください。本ツールではこのプレフィックスがついたセル名をJSON出力に用います。

### 📋 JSON Schema サポート
`--schema` オプションで JSON Schema を指定することで、データのバリデーションおよびキー順序の指定が可能です。バリデーションエラーは `<basename>.error.log` に出力されます。

### 🔀 セル名の禁則文字をJSON項目名に使用可能
セル名に使えない記号をJSONの項目名に使用したい場合、アンダーバー（`_`）への置き換え および JSON Schemaの併用により対応可能です。セル名のアンダーバーは1文字のワイルドカードとみなされ、JSON Schemaの項目名と照合～置き換えて出力されます。

### 🛠️ ユーザ定義の変換ルール
セル名ごとに「変換ルール」を指定することで、セル値の編集や1次元～多次元配列への変換に対応。
変換ルールには区切り文字を指定した配列変換や、任意のPython関数・外部コマンドを利用可能。

---

## インストール
```bash
git clone https://github.com/hat47x/xlsx2json.git
cd xlsx2json

# 依存関係をインストール
pip install -r requirements.txt
```

---

## クイックスタート

### 1. サンプルファイルで動作確認
```bash
# サンプルファイルを変換
python xlsx2json.py samples/sample.xlsx

# 結果確認（samples/output-json/sample.json が生成されます）
cat samples/output-json/sample.json
```

### 2. JSON Schema を使ったバリデーション
```bash
# スキーマ付きでバリデーション
python xlsx2json.py samples/sample.xlsx --schema samples/schema.json
```

### 3. 変換ルールを適用
```bash
# カンマ区切りデータを配列に変換
python xlsx2json.py samples/sample.xlsx --transform "json.parent=split:,"

# 設定ファイルを用いることで大量の変換ルールにも対応
python xlsx2json.py samples/sample.xlsx --config samples/config.json
```

---

## 前提条件

### システム要件
- **Python**: 3.8+

### 依存関係
- openpyxl
- jsonschema

### 対応ファイル形式
- **入力**: Excel (.xlsx) ファイル
- **出力**: JSON (.json) ファイル

---

## 使用方法

### 基本構文

```bash
python xlsx2json.py [INPUT1 ...] [OPTIONS]
```

### オプション一覧

| オプション | 説明 |
|:----------|:-----|
| `INPUT1 ...` | 変換対象のファイル（.xlsx）またはフォルダ。フォルダ指定時は直下の `.xlsx` ファイルを対象とします。省略時は `--config` で指定が必要です。 |
| `-o, --output-dir` | 一括出力先フォルダを指定。省略時は各入力ファイルと同じディレクトリの `output-json/` に出力されます。 |
| `-s, --schema` | JSON Schema ファイルを指定。バリデーションやキー順序の整理などに使用されます。 |
| `--transform RULE` | 変換ルールを指定。指定したセル名の値に対し、split（区切り文字による配列化）、function（Python関数）、command（外部コマンド）による変換を適用（複数指定可）。 |
| `--keep-empty` | 空のセル値も JSON に含めます（デフォルト: true）。 |
| `--prefix PREFIX` | Excel セル名のプレフィックスを指定（デフォルト: `json`）。 |
| `--log-level LEVEL` | ログレベルを指定（`DEBUG`/`INFO`/`WARNING`/`ERROR`/`CRITICAL`、デフォルト: `INFO`）。 |
| `--config FILE` | 設定ファイル（JSON形式）から全オプションを一括指定。コマンドライン引数が優先されます。 |

---

## 設定ファイルによるオプション指定

`--config` オプションで、変換ルール以外も含めた全オプションをJSON形式で一括指定できます。

### 設定ファイル例（config.json）
```json
{
  "inputs": ["samples/sample.xlsx"],
  "output_dir": "output-json",
  "schema": "samples/schema.json",
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
# 設定ファイルから全オプションを指定
python xlsx2json.py --config config.json

# 設定ファイル + コマンドライン引数の組み合わせ（コマンドライン引数が優先）
python xlsx2json.py --config config.json --log-level DEBUG
```

コマンドライン引数で指定した値は、設定ファイルより優先されます。

---

## 処理の流れ

1. **ファイル収集**: 入力パスから `.xlsx` ファイルを収集
2. **データ読み込み**: セル名（`json.*`）を読み込み、ネスト構造の辞書・リストを生成
3. **変換処理**: `--transform` オプションで指定したルールに従い、値の配列化・関数変換・コマンド変換等を適用
4. **空値処理**: （デフォルト）空の値を除去（`--keep-empty` オプションで保持可能）
5. **バリデーション**: （オプション）Schema バリデーションとエラーログ出力
6. **キー整理**: （オプション）Schema の `properties` 順にキーを整理
7. **ファイル出力**: JSON ファイルとして書き出し

---

## セル名の命名規則

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
- Unicode 文字をサポート（日本語の漢字・ひらがな・カタカナ、英数字、アンダーバー `_` 等）
  - 記号文字は全角・半角を問わず多くが未サポートである点に注意
- 大文字・小文字は区別されます

#### 文字数制限
- Excel の仕様上、名前の長さは 255 文字以内です

#### 記号文字の注意点
記号文字をキー名に使用する際は、実際に使用できる文字種であるかを確認してください。
- ✅ 使用可能: "。", "・", "＿"
- ❌ 使用不可: "、", "！", "～", "／", "（", "）" など

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

Excelのセル名から抽出した値に対して、柔軟な変換ルールを指定できます。

### 配列化（split）

区切り文字で値を配列化したい場合は、`split:` 変換タイプを使います。

#### 実行例

```bash
# カンマ区切りで配列化
python xlsx2json.py samples/sample.xlsx --transform "json.tags=split:,"
# → ["apple", "banana", "orange"]

# 改行区切りで配列化
python xlsx2json.py samples/sample.xlsx --transform "json.parent.1=split:\n"
# → ["A", "B", "C"]

# 多次元配列（セミコロン→カンマ）
python xlsx2json.py samples/sample.xlsx --transform "json.matrix=split:;|,"

# 3次元配列（セミコロン→パイプ→カンマ）
python xlsx2json.py samples/sample.xlsx --transform "json.cube=split:;|\||,"
# → [[["A", "B"], ["C", "D"]], [["E", "F"], ["G", "H"]]]

# パイプ文字を区切り文字として使用（エスケープが必要）
python xlsx2json.py samples/sample.xlsx --transform "json.data=split:\|"
# → ["A", "B", "C"]
```

**重要**: 区切り文字は `|` で区切って多次元配列に対応します。パイプ文字自体を区切り文字にする場合は `\|` でエスケープしてください。改行（`\n`）、タブ（`\t`）、復帰（`\r`）も利用可能です。

### Python関数による変換（function）

```bash
# モジュール内の関数を指定
python xlsx2json.py samples/sample.xlsx --transform "json.tags=function:mymodule:split_func"
# ファイル内の関数を指定
python xlsx2json.py samples/sample.xlsx --transform "json.tags=function:/path/to/script.py:split_func"
```

### 外部コマンドによる変換（command）

```bash
# sortコマンドを指定
python xlsx2json.py samples/sample.xlsx --transform "json.lines=command:sort -u"
```

## 記号ワイルドカード対応

### 概要

Excelの「名前の定義」で使えない記号も、アンダーバー（`_`）を1文字ワイルドカードとしてJSON Schemaの項目名にユニークにマッチする場合は置き換えてJSONのキー名として出力されます。
置き換え後のキー名は、変換ルール内でも使うことができます。

### 動作原理

1. Excelのセル名で `json.` 以降のキー名にアンダーバー（`_`）を使用
2. その位置が「任意の1文字」としてJSON Schemaのプロパティ名にマッチします
3. マッチするJSON Schemaのプロパティが1つだけの場合、そのプロパティ名に自動で置換されます
4. 複数マッチまたはマッチしない場合は、元のキー名のまま出力されます

### 使用例

#### JSON Schema例
```json
{
  "type": "object",
  "properties": {
    "test_name": {"type": "string"},
    "test-group": {"type": "string"},
    "test！": {"type": "string"},
    "test？": {"type": "string"},
    "★test！？": {"type": "string"}
  }
}
```

#### Excelでのセル名の定義

| Excelの名前定義 | 対応するJSONキー | 備考 |
|---|---|---|
| `json.test_name` | `test_name` | そのまま一致 |
| `json.test_group` | `test-group` | ワイルドカードによるマッチング |
| `json.test_` | `test！, test？` | 複数マッチ時は置換されない（警告ログ出力） |
| `json._test__` | `★test！？` | ワイルドカードは1文字ずつのマッチング |

### 注意点

- アンダーバーの数だけ1文字ずつ任意の文字にマッチします
- マッチするSchemaプロパティが複数ある場合は、Excelで指定した名前がそのまま使われます
- ワイルドカード置換は各階層ごとに適用されます
- 置換の有無やマッチ状況は、`--log-level DEBUG` で確認できます

### 実行例

```bash
# 基本的な使用
python xlsx2json.py samples/sample.xlsx --schema samples/schema.json

# ログでマッチ状況を確認
python xlsx2json.py samples/sample.xlsx --schema samples/schema.json --log-level DEBUG

# 変換ルール内のキー名にも利用可能
python xlsx2json.py samples/sample.xlsx --schema samples/schema.json --transform "json.★test！？=split:,"

# アンダーバー表記での指定も可能
python xlsx2json.py samples/sample.xlsx --schema samples/schema.json --transform "json._test__=split:,"
```

---

## 制限事項

### 出力フォーマット
- JSON Schema Draft-07
- UTF-8

---

## 開発・テスト

```bash
# 開発用依存関係をインストール
pip install -r requirements-dev.txt

# ユニットテスト実行
python -m pytest test_xlsx2json.py -v

# コードカバレッジ確認
python -m pytest test_xlsx2json.py --cov=xlsx2json --cov-report=html
open htmlcov/index.html
```

---

## ライセンス

MIT