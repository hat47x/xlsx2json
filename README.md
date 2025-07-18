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
| `--keep-empty` | 空のセル値も JSON に含めます（デフォルトでは空値を除去）。 |
| `--prefix PREFIX` | Excel 名前付き範囲のプレフィックスを指定（デフォルト: `json.`）。 |
| `--log-level LEVEL` | ログレベルを指定（`DEBUG`/`INFO`/`WARNING`/`ERROR`/`CRITICAL`、デフォルト: `INFO`）。 |
| `--array-split RULE` | 配列化設定。指定した名前付き範囲の値を区切り文字で分割して配列化（複数指定可）。 |

---

## 処理の流れ

1. **ファイル収集**: 入力パスから `.xlsx` ファイルを収集
2. **データ読み込み**: 名前付き範囲（`json.*`）を読み込み、ネスト構造の辞書・リストを生成
3. **空値処理**: （デフォルト）空の値を除去（`--keep-empty` オプションで保持可能）
4. **バリデーション**: （オプション）Schema バリデーションとエラーログ出力
5. **キー整理**: （オプション）Schema の `properties` 順にキーを整理
6. **ファイル出力**: JSON ファイルとして書き出し

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

## 配列化オプション（--array-split）

### 基本的な使用方法

`--array-split` オプションは、Excelの名前付き範囲の値を区切り文字で分割し、JSON配列として出力します。

#### 構文
```bash
--array-split "json.パス=区切り文字"
```

#### 基本例
```bash
# カンマ区切りで配列化
python xlsx2json.py sample.xlsx --array-split "json.tags=,"

# 改行区切りで配列化
python xlsx2json.py sample.xlsx --array-split "json.parent.1=\n"

# 複数指定
python xlsx2json.py sample.xlsx --array-split "json.tags=," --array-split "json.parent.1=\n"
```

### 多次元配列の対応

#### 区切り文字の指定方法
パイプ文字（`|`）で区切って、1次元目、2次元目、3次元目...の区切り文字を指定できます。

```bash
# 1次元配列: カンマ区切り
--array-split "json.data=,"
# "A,B,C" → ["A", "B", "C"]

# 2次元配列: セミコロンで1次元目、カンマで2次元目を区切り
--array-split "json.matrix=;|,"
# "A,B;C,D" → [["A", "B"], ["C", "D"]]

# 3次元配列: セミコロン、パイプ、カンマで各次元を区切り
--array-split "json.cube=;|\||,"
# "A,B|C,D;E,F|G,H" → [[["A", "B"], ["C", "D"]], [["E", "F"], ["G", "H"]]]
```

### エスケープ処理

#### パイプ文字を区切り文字として使用
パイプ文字（`|`）を実際の区切り文字として使用したい場合は、`\|` でエスケープします。

```bash
# パイプ文字を1次元目の区切り文字として使用
--array-split "json.data=\\|"
# "A|B|C" → ["A", "B", "C"]

# パイプ文字を1次元目、カンマを2次元目の区切り文字として使用
--array-split "json.matrix=\\||,"
# "A,B|C,D" → [["A", "B"], ["C", "D"]]

# セミコロンを1次元目、パイプ文字を2次元目の区切り文字として使用
--array-split "json.mixed=;|\\|"
# "A|B;C|D" → [["A", "B"], ["C", "D"]]

# 複雑な例：3次元配列でパイプ文字を含む
--array-split "json.complex=;|\\||,"
# "A,B|C,D;E,F|G,H" → [[["A", "B"], ["C", "D"]], [["E", "F"], ["G", "H"]]]
```

#### 標準的なエスケープシーケンス
```bash
# 改行、タブ、復帰文字
--array-split "json.data=\n"  # 改行区切り
--array-split "json.data=\t"  # タブ区切り
--array-split "json.data=\r"  # 復帰文字区切り
```

### 実際の使用例

| Excelの名前定義 | セル値 | オプション例 | 出力例 |
|---|---|---|---|
| `json.tags` | `apple,banana,orange` | `--array-split "json.tags=,"` | `["apple", "banana", "orange"]` |
| `json.parent.1` | `A\nB\nC` | `--array-split "json.parent.1=\n"` | `["A", "B", "C"]` |
| `json.matrix` | `A,B;C,D` | `--array-split "json.matrix=;|,"` | `[["A", "B"], ["C", "D"]]` |
| `json.data` | `A\|B\|C` | `--array-split "json.data=\\|"` | `["A", "B", "C"]` |

### 注意点

- プレフィックス（`json.`など）は自動的に除去されます
- 区切り文字が含まれない場合は1要素の配列となります
- 複数の `--array-split` を指定すると、それぞれのパスに個別に適用されます
- 配列化対象のパスは、Excelの名前付き範囲の階層表現で指定します

---

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
  "type": "object",
  "properties": {
    "customer_name": {"type": "string"},
    "parent": {
      "type": "array",
      "items": {
        "type": "array",
        "items": {"type": "string"}
      }
    }
  },
  "required": ["customer_name", "parent"],
  "additionalProperties": false
}
```

---

## 実用的な使用例

### 基本的な使用

```bash
# 単一ファイルの変換
python xlsx2json.py data.xlsx

# 複数ファイルの変換
python xlsx2json.py file1.xlsx file2.xlsx

# フォルダ内のすべての .xlsx ファイルを変換
python xlsx2json.py ./data_folder/
```

### 高度な使用例

```bash
# 出力先を指定
python xlsx2json.py data.xlsx --output-dir ./json_output/

# JSON Schema を使用してバリデーション
python xlsx2json.py data.xlsx --schema schema.json

# 空のセル値も含めて出力
python xlsx2json.py data.xlsx --keep-empty

# 名前付き範囲のプレフィックスを変更
python xlsx2json.py data.xlsx --prefix myprefix.

# ログレベルをDEBUGにして詳細な処理ログを出力
python xlsx2json.py data.xlsx --log-level DEBUG

# 複数のオプションを組み合わせ
python xlsx2json.py ./data/ \
  --output-dir ./output/ \
  --schema schema.json \
  --keep-empty \
  --prefix myprefix. \
  --log-level DEBUG \
  --array-split "json.tags=," \
  --array-split "json.matrix=;|,"
```

### 実際のワークフロー例

```bash
# 開発環境での検証
python xlsx2json.py sample.xlsx --schema validation.json --log-level DEBUG

# 本番環境での一括処理
python xlsx2json.py ./production_data/ \
  --output-dir ./json_output/ \
  --schema production_schema.json \
  --array-split "json.categories=," \
  --array-split "json.nested_data=;|,"
```

---

## ライセンス

MIT