# xlsx2json

Excel### 🔄 繰り返し構造の自動処理
コンテナ機能により、Excel の繰り返し構造（テーブル、カード型レイアウト、階層構造）を自動検出・処理できます。罫線解析による構造判定で、手動でのセル名設定作業を大幅に軽減します。

### 📋 JSON Schema サポート
`--schema` オプションで JSON Schema を指定することで、データのバリデーションおよびキー順序の指定が可能です。バリデーションエラーは `<basename>.error.log` に出力されます。

### 📝 YAML設定ファイル対応
設定ファイル（`--config` オプション）でJSON形式またはYAML形式を選択できます。YAMLではコメントの記述や、より読みやすい階層表現が可能です。既存のJSON設定ファイルもそのまま使用できます（JSONはYAMLのサブセット）。

### 🔀 セル名の禁則文字をJSON項目名に使用可能て事前に定義されたデータ構造にもとづき、JSON形式またはYAML形式に変換するツールです。

## 特徴

Excelの「名前付き範囲」（以降「セル名」）を用いて事前に定義されたデータ構造にもとづき、JSON形式またはYAML形式に変換するツールです。Excelファイルの様式が入り組んでいても、 **セル名さえつけてしまえば様式ごとの変換ロジックが不要** な点が本ツールの特徴です。

一方デメリットとしては、事前にセル名をつける作業が必要となります。特に、様式にもとづき自動で繰り返し要素を識別することができないため、例えば5列×100行の入力項目が存在する場合、それらすべてにセル名をつける必要があります。また行数を増やしたい場合は追加でセル名をつける必要があります。
※作業負荷を軽減するため「セル名インポート・エクスポート用マクロ.xlsm」を同梱しています

上記のトレードオフを念頭に、本ツールの利用をご検討ください。

### 🔄 セル名でJSON出力時の階層構造を指定
セル名にドット `.` 区切りのキー階層を記載することで、自動的にネストした JSON 構造に変換します。複雑な階層データも直感的に定義可能です。
セル名の先頭には `json.` を付加してください。本ツールではこのプレフィックスがついたセル名をJSON出力に用います。

### � 複数の出力フォーマット対応
`--output-format` オプションで JSON または YAML 形式での出力を選択できます。YAML 形式では、より読みやすい形式でデータを表示でき、設定ファイルや人間が確認するためのドキュメントに適しています。

### �🔄 繰り返し構造の自動処理
コンテナ機能により、Excel の繰り返し構造（テーブル、カード型レイアウト、階層構造）を自動検出・処理できます。罫線解析による構造判定で、手動でのセル名設定作業を大幅に軽減します。

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
# サンプルファイルを変換（JSON形式 - デフォルト）
python xlsx2json.py samples/sample.xlsx

# 結果確認（samples/output/sample.json が生成されます）
cat samples/output/sample.json

# YAML形式で出力
python xlsx2json.py samples/sample.xlsx --output-format yaml

# 結果確認（samples/output/sample.yaml が生成されます）
cat samples/output/sample.yaml
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

# シンプルな関数名で文字列正規化
python xlsx2json.py samples/sample.xlsx --transform "json.name=function:samples/transform.py:normalize"

# 範囲データの合計計算 (Excelでjson.totalセル名が範囲を指している場合)
python xlsx2json.py samples/sample.xlsx --transform "json.total=function:samples/transform.py:total"

# 連続適用（チェーン変換）
python xlsx2json.py samples/sample.xlsx \
  --transform "json.data=split:," \
  --transform "json.data=function:samples/transform.py:clean" \
  --transform "json.data=function:samples/transform.py:normalize"
```

### 4. コンテナによる繰り返し構造の自動処理
```bash
# テーブル形式の繰り返しデータを自動検出・変換
python xlsx2json.py samples/sample.xlsx \
  --container '{"json.orders":{"range":"A1:C10","direction":"row","items":["date","customer","amount"]}}'
```

### 5. 設定ファイルの利用
```bash
# JSON形式の設定ファイルで大規模な変換ルールや複雑な階層構造を定義
python xlsx2json.py samples/sample.xlsx --config samples/config.json

# YAML形式の設定ファイル（コメント付きで読みやすい）
python xlsx2json.py samples/sample.xlsx --config samples/config.yaml
```

---

## 前提条件

### システム要件
- **Python**: 3.8+

### 依存関係
- openpyxl
- jsonschema
- pyyaml

### 対応ファイル形式
- **入力**: Excel (.xlsx) ファイル
- **出力**: JSON (.json) ファイル、YAML (.yaml) ファイル

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
| `-o, --output-dir` | 一括出力先フォルダを指定。省略時は各入力ファイルと同じディレクトリの `output/` に出力されます。 |
| `-f, --output-format FORMAT` | 出力フォーマットを指定（`json` または `yaml`、デフォルト: `json`）。 |
| `-s, --schema` | JSON Schema ファイルを指定。バリデーションやキー順序の整理などに使用されます。 |
| `--transform RULE` | 変換ルールを指定。同一セル名に対して複数指定した場合は連続適用（チェーン）されます。split（区切り文字による配列化）、function（Python関数）、command（外部コマンド）による変換を適用可能。セル名が指し示すデータ形式（値・1次元配列・2次元配列）に応じて自動的に適切な形式で変換関数に渡します。 |
| `--container DEFINITION` | コンテナ定義を指定。Excel の繰り返し構造（テーブル、カード、階層構造）を自動検出・処理（複数指定可）。JSON形式で定義します。 |
| `--keep-empty` | 空のセル値も JSON に含めます（デフォルト: true）。 |
| `--prefix PREFIX` | Excel セル名のプレフィックスを指定（デフォルト: `json`）。 |
| `--log-level LEVEL` | ログレベルを指定（`DEBUG`/`INFO`/`WARNING`/`ERROR`/`CRITICAL`、デフォルト: `INFO`）。 |
| `--config FILE` | 設定ファイルから全オプションを一括指定。コマンドライン引数が優先されます。 |

---

## 設定ファイルによるオプション指定

`--config` オプションで、変換ルール以外も含めた全オプションをYAML形式で一括指定できます。JSONはYAMLのサブセットなので、JSONで記述することもできます。


### 設定ファイル例（config.yaml）
```yaml
# YAML形式の設定ファイル例
input-files:
  - samples/sample.xlsx
output-dir: output
output-format: yaml
schema: samples/schema.json

# 変換ルール
transform:
  - "json.tags=split:,"
  - "json.matrix=split:,|\n"
  - "json.orders.*.date=function:date:parse_japanese_date"
  - "json.orders.*.amount=function:math:parse_currency"
  - "json.orders.*.items.*.unit_price=function:math:parse_currency"

# コンテナ定義
containers:
  json.orders:
    range: orders_range
    direction: row
    increment: 1
    items:
      - date
      - customer_id
      - amount
    labels:
      - 注文日
      - 顧客ID
      - 金額
  
  json.orders.1.items:
    offset: 3
    items:
      - product_code
      - quantity
      - unit_price
    labels:
      - 商品コード
      - 数量
      - 単価

# その他のオプション
prefix: json
keep-empty: false
log-level: INFO
```

### 実行例
```bash

# 設定ファイルから全オプションを指定
python xlsx2json.py --config config.yaml

# 設定ファイル + コマンドライン引数の組み合わせ（コマンドライン引数が優先）
python xlsx2json.py --config config.yaml --log-level DEBUG

# 設定ファイル + 追加のコンテナ定義
python xlsx2json.py --config config.yaml \
  --container '{"json.additional":{"range":"E1:G10","direction":"row","items":["extra1","extra2"]}}'
```

コマンドライン引数で指定した値は、設定ファイルより優先されます。

---

## 処理の流れ

1. **ファイル収集**: 入力パスから `.xlsx` ファイルを収集
2. **データ読み込み**: セル名（`json.*`）を読み込み、ネスト構造の辞書・リストを生成
3. **変換処理**: `--transform` オプションで指定したルールに従い、値の配列化・関数変換・コマンド変換等を適用。セル名が指し示すデータ形式（値・1次元配列・2次元配列）に応じて自動的に適切な形式で変換関数に渡します。同一セル名に対する複数ルールは順次連続適用（チェーン）されます。
4. **空要素削除**: 全フィールドが未設定（空値・空リスト・空辞書）の要素を再帰的に除去
5. **空値処理**: （デフォルト）空の値を除去（`--keep-empty` オプションで保持可能）
6. **バリデーション**: （オプション）Schema バリデーションとエラーログ出力
7. **キー整理**: （オプション）Schema の `properties` 順にキーを整理
8. **ファイル出力**: JSON または YAML ファイルとして書き出し

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

### 🔗 連続適用（チェーン変換）

同一セル名に対して複数の`--transform`を指定することで、**変換ルールを順次連続適用**できます。これにより、複雑なデータ処理パイプラインを構築可能です。

```bash
# 基本的な連続適用
python xlsx2json.py sample.xlsx \
  --transform "json.data=split:," \
  --transform "json.data=function:samples/transform.py:clean" \
  --transform "json.data=function:samples/transform.py:normalize"

# ワイルドカードと連続適用の組み合わせ
python xlsx2json.py sample.xlsx \
  --transform "json.orders.*.amount=split:," \
  --transform "json.orders.*.amount=function:samples/transform.py:total" \
  --transform "json.orders.*.amount=function:samples/transform.py:normalize"
```

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

### セル範囲データの自動判定変換

セル名が単一値、1次元配列、2次元配列のいずれを指すかによって、変換関数に適切な形式でデータを渡します。

#### データ形式の自動判定

- **単一セル（値）**: セル値そのものが変換関数に渡されます
- **1次元範囲（行または列）**: `[値1, 値2, 値3, ...]` として配列で渡されます  
- **2次元範囲**: `[[行1の値...], [行2の値...], ...]` として2次元配列で渡されます

#### 実行例

```bash
# セル名 json.matrix が単一セル A1 を指している場合
# → セル値がそのまま変換関数に渡される
python xlsx2json.py samples/sample.xlsx --transform "json.matrix=function:mymodule:transform_value"

# セル名 json.matrix が1次元範囲 A1:A5 を指している場合  
# → [値1, 値2, 値3, 値4, 値5] として配列で渡される
python xlsx2json.py samples/sample.xlsx --transform "json.matrix=function:mymodule:transform_array"

# セル名 json.matrix が2次元範囲 A1:C3 を指している場合
# → [[行1の値...], [行2の値...], [行3の値...]] として2次元配列で渡される
python xlsx2json.py samples/sample.xlsx --transform "json.matrix=function:mymodule:transform_matrix"

# 外部コマンドでの処理
# 1次元データは改行区切りで渡される
python xlsx2json.py samples/sample.xlsx --transform "json.sorted_list=command:sort -n"

# 2次元データはTSV形式で渡される  
python xlsx2json.py samples/sample.xlsx --transform "json.sorted_table=command:sort -t$'\t' -k2,2n"
```

#### 辞書戻り値による動的セル名構築

任意の変換関数（function、command、split）が辞書を返すことで、実行時にセル名と値の関係を動的に構築できます：

```python
# カスタム変換関数の例（mymodule.py）
def create_dynamic_fields(data):
    """データから動的にセル名を生成"""
    result = {}
    if isinstance(data, list):
        for i, value in enumerate(data):
            if value is not None:
                result[f"field_{i+1}"] = str(value).strip()
                result[f"field_{i+1}_upper"] = str(value).upper()
    elif isinstance(data, str):
        # カンマ区切り文字列をパース
        parts = data.split(",")
        for i, part in enumerate(parts):
            result[f"item_{i+1}"] = part.strip()
    return result

def parse_key_value_data(data):
    """キー=値形式のデータを解析"""
    result = {}
    if isinstance(data, list):
        for row in data:
            if isinstance(row, list) and len(row) >= 2:
                key = str(row[0]).strip() if row[0] else None
                value = row[1] if row[1] else None
                if key and value:
                    result[key] = value
    elif isinstance(data, str):
        # "key1=value1;key2=value2" 形式をパース
        pairs = data.split(";")
        for pair in pairs:
            if "=" in pair:
                k, v = pair.split("=", 1)
                result[k.strip()] = v.strip()
    return result
```

**使用例:**
```bash
# 変換関数で動的フィールドを生成し、生成されたフィールドにさらに変換を適用
python xlsx2json.py sample.xlsx \
  --transform "json.dynamic=function:mymodule:create_dynamic_fields" \
  --transform "json.dynamic.field_1=function:samples/transform.py:normalize" \
  --transform "json.dynamic.field_*_upper=split:,"

# キー=値ペアを解析して動的セル名を構築
python xlsx2json.py sample.xlsx \
  --transform "json.config=function:mymodule:parse_key_value_data" \
  --transform "json.config.*=function:samples/transform.py:normalize"

# 外部コマンドでJSONを返すことで動的セル名構築
python xlsx2json.py sample.xlsx \
  --transform "json.api_data=command:curl -s api.example.com/config | jq -r '.fields'"
```

**辞書戻り値の仕様:**
- **絶対指定**: キーが `json.` で始まる場合、JSONルートからの絶対パスとして扱われます
- **相対指定**: それ以外のキーは、元の変換対象パスからの相対パスとして扱われます
- **連続適用対応**: 生成されたセル名に対しても、指定された変換ルールが自動的に適用されます

### Python関数による変換（function）

#### 基本的な使用方法
```bash
# モジュール内の関数を指定
python xlsx2json.py samples/sample.xlsx --transform "json.tags=function:mymodule:split_func"
# ファイル内の関数を指定
python xlsx2json.py samples/sample.xlsx --transform "json.tags=function:/path/to/script.py:split_func"
```

#### samples/transform.py サンプル関数

`samples/transform.py` には即座に使える便利な変換関数が用意されています。シンプルで覚えやすい関数名で、様々なデータ変換が可能です。

##### 🔤 文字列変換
```bash
# CSV文字列を配列に変換
python xlsx2json.py samples/sample.xlsx --transform "json.data=function:samples/transform.py:csv"

# 改行区切りを配列に変換  
python xlsx2json.py samples/sample.xlsx --transform "json.lines=function:samples/transform.py:lines"

# 文字列正規化（トリム・全角半角変換・置換）
python xlsx2json.py samples/sample.xlsx --transform "json.name=function:samples/transform.py:normalize"
```

##### 📊 配列・行列操作（セル名が範囲を指している場合）
```bash
# 指定列を抽出（セル名 json.names が範囲 A1:C10 を指している場合）
python xlsx2json.py samples/sample.xlsx --transform "json.names=function:samples/transform.py:column"

# 行と列を入れ替え（転置）（セル名 json.matrix が範囲 A1:C3 を指している場合）
python xlsx2json.py samples/sample.xlsx --transform "json.matrix=function:samples/transform.py:flip"

# 空でない行のみを残す（セル名 json.data が範囲 A1:D20 を指している場合）
python xlsx2json.py samples/sample.xlsx --transform "json.data=function:samples/transform.py:clean"
```

##### 🔢 数値計算（セル名が範囲を指している場合）
```bash
# 全要素の合計（セル名 json.total が範囲 A1:C3 を指している場合）
python xlsx2json.py samples/sample.xlsx --transform "json.total=function:samples/transform.py:total"

# 数値要素の平均（セル名 json.average が範囲 A1:C3 を指している場合）
python xlsx2json.py samples/sample.xlsx --transform "json.average=function:samples/transform.py:avg"

# 指定列の合計（セル名 json.sum が範囲 A1:C10 を指している場合）
python xlsx2json.py samples/sample.xlsx --transform "json.sum=function:samples/transform.py:sum_col"
```

##### 💡 サンプル関数の特徴
- **シンプルな関数名**: `csv()`, `flip()`, `total()`, `normalize()` など覚えやすい名前
- **後方互換性**: 従来の関数名（`csv_split`, `extract_column`等）も引き続き使用可能
- **カスタマイズ可能**: `samples/transform.py` を参考に独自の変換関数を作成
- **詳細情報**: `samples/README.md` に全関数の説明あり

### 外部コマンドによる変換（command）

```bash
# sortコマンドを指定
python xlsx2json.py samples/sample.xlsx --transform "json.lines=command:sort -u"
```

### ワイルドカード対応

変換ルールでワイルドカード `*` を使用することで、複数のセル名に対して一括でルールを適用できます。これにより、コンテナ機能で自動生成されるセル名にも効率的にルールを適用可能です。

#### 基本的なワイルドカード使用例

```bash
# 全ての orders インスタンスの amount フィールドに適用
python xlsx2json.py sample.xlsx --transform "json.orders.*.amount=function:math:parse_currency"

# 全ての items の price フィールドに適用  
python xlsx2json.py sample.xlsx --transform "json.orders.*.items.*.price=function:math:parse_currency"

# 全ての date フィールドに適用
python xlsx2json.py sample.xlsx --transform "json.*.date=function:date:parse_japanese_date"
```

#### 階層ワイルドカード

```bash
# 複数レベルの階層に適用
python xlsx2json.py sample.xlsx --transform "json.customers.*.orders.*.date=function:date:parse"

# ツリー構造の seq フィールドに適用
python xlsx2json.py sample.xlsx --transform "json.tree_data.lv1.*.seq=function:math:parse_number"
python xlsx2json.py sample.xlsx --transform "json.tree_data.lv1.*.lv2.*.seq=function:math:parse_number"
```

#### コンテナ機能との連携例

**コンテナ定義:**
```json
{
  "json.orders": {
    "range": "orders_range",
    "direction": "row", 
    "items": ["date", "customer_id", "amount", "items"]
  },
  "json.orders.1.items": {
    "offset": 3,
    "items": ["product_code", "quantity", "unit_price"]
  }
}
```

**ワイルドカード変換ルール:**
```bash
# 全ての注文の日付を変換
--transform "json.orders.*.date=function:date:parse_japanese_date"

# 全ての金額を通貨形式で変換  
--transform "json.orders.*.amount=function:math:parse_currency"

# 全ての商品単価を通貨形式で変換
--transform "json.orders.*.items.*.unit_price=function:math:parse_currency"

# 全ての数量を数値に変換
--transform "json.orders.*.items.*.quantity=function:math:parse_number"
```

**自動生成されるセル名（例）:**
- `json.orders.1.date`, `json.orders.2.date`, `json.orders.3.date` ...
- `json.orders.1.items.1.unit_price`, `json.orders.1.items.2.unit_price` ...
- `json.orders.2.items.1.unit_price`, `json.orders.2.items.2.unit_price` ...

**適用結果:**
ワイルドカード `json.orders.*.date` は、自動生成された全ての注文日付セル名（`json.orders.1.date`, `json.orders.2.date` など）にマッチし、指定した変換関数が適用されます。

#### ワイルドカード使用時の注意点

1. **パフォーマンス**: ワイルドカードは処理時にパターンマッチングを行うため、大量のセル名がある場合は処理時間が増加する場合があります。

2. **優先順位**: より具体的なルールが優先されます。
   ```bash
   # 具体的なルールが優先される
   --transform "json.orders.1.amount=function:custom:special_parse"
   --transform "json.orders.*.amount=function:math:parse_currency"
   ```

3. **複数マッチ**: 一つのセル名が複数のワイルドカードルールにマッチする場合、最後に指定されたルールが適用されます。

4. **連続適用**: 同一セル名に対する複数の`--transform`指定は、**指定された順序で連続適用（チェーン）されます**。これにより複雑なデータ変換パイプラインを構築できます。
   ```bash
   # 以下の場合、split → normalize の順で適用される
   --transform "json.orders.*.amount=split:,"
   --transform "json.orders.*.amount=function:samples/transform.py:normalize"
   ```
   
   **連続適用の実行例:**
   ```bash
   # 1. CSV分割 → 2. データクリーニング → 3. 正規化
   python xlsx2json.py sample.xlsx \
     --transform "json.data=split:," \
     --transform "json.data=function:samples/transform.py:clean" \
     --transform "json.data=function:samples/transform.py:normalize"
   ```

5. **辞書戻り値による動的セル名構築**: 変換関数が辞書を返すことで、実行時にセル名と値の関係を動的に構築できます。生成されたセル名に対しても、追加の変換ルールが自動的に適用されます。
   ```bash
   # 変換関数で辞書を返し、生成されたセル名にさらに変換を適用
   --transform "json.dynamic=function:mymodule:create_dict"
   --transform "json.dynamic.item1=function:samples/transform.py:normalize"
   --transform "json.dynamic.item2=split:,"
   ```

---

## コンテナ指定（--container）

Excelの繰り返し構造（テーブル、カード型レイアウト、階層構造）を自動検出・処理するためのコンテナ定義をコマンドラインで指定できます。

### 基本書式

```bash
--container '{"セル名": JSON定義}'
```

JSON定義は、設定ファイルの `containers` セクションと同じ形式を使用します。

#### 使用例

##### 基本的なテーブル定義
```bash
# シンプルなテーブル
python xlsx2json.py sample.xlsx \
  --container '{"json.orders":{"range":"A1:C10","direction":"row","items":["date","customer","amount"]}}'

# ラベル検証付き
python xlsx2json.py sample.xlsx \
  --container '{"json.orders":{"range":"orders_range","direction":"row","items":["date","customer","amount"],"labels":["注文日","顧客名","金額"]}}'
```

##### 親子関係（ネスト構造）
```bash
# 親コンテナ
python xlsx2json.py sample.xlsx \
  --container '{"json.orders":{"range":"orders_range","direction":"row","items":["date","customer","items"]}}' \
  --container '{"json.orders.1.items":{"offset":3,"items":["product","quantity","price"]}}'
```

##### カード型レイアウト
```bash
# 多段組みカード
python xlsx2json.py sample.xlsx \
  --container '{"json.customers":{"range":"A1:C20","direction":"column","increment":5,"items":["name","phone","address"]}}'
```

##### 階層構造（ツリー）
```bash
# 3階層のツリー構造
python xlsx2json.py sample.xlsx \
  --container '{"json.tree_data":{"range":"tree_range","direction":"row","items":["name","value","seq"]}}' \
  --container '{"json.tree_data.lv1.1.lv2":{"offset":1,"items":["name","value","seq"]}}' \
  --container '{"json.tree_data.lv1.1.lv2.1.lv3":{"offset":1,"items":["name","value","seq"]}}'
```

### 設定ファイルとの組み合わせ

```bash
# 設定ファイル + 追加のコンテナ定義
python xlsx2json.py --config config.json \
  --container '{"json.additional":{"range":"Z1:Z10","direction":"row","items":["extra_field"]}}'
```

### 注意点

- JSON文字列はシェルでのエスケープに注意（シングルクォートを推奨）
- 複数のコンテナは `--container` を複数回指定
- コマンドライン指定は設定ファイルの `containers` より優先されます
- 同じセル名の定義が重複した場合、後から指定したものが優先されます

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