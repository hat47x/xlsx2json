# xlsx2json

Excel の名前付き範囲（`json.*`）を読み込んで JSON 形式に変換し、ファイル出力するコマンドラインツールです。

---

## 特徴

- **名前付き範囲の自動解析**
  セル名に `json.` プレフィックスを付けるだけで、自動的に JSON 形式に変換します。
- **JSON Schema サポート**
  `--schema` オプションで JSON Schema を指定することで、データのバリデーションとキー順序の整理が可能です。
- **記号ワイルドカード対応**
  Excelの「名前の定義」で使えない記号も、アンダーバー（_）を1文字ワイルドカードとしてJSON Schemaの項目名にユニークにマッチする場合は置き換えて出力されます。
- **エラーログ出力**
  スキーマ違反があった場合は `<basename>.error.log` に詳細なエラー情報を出力します。
- **空値の自動除去**
  空のセル値はデフォルトで JSON から除外され、必要に応じて `--keep-empty` オプションで保持可能です。
- **柔軟な出力先設定**
  個別フォルダへの出力、または `--output-dir` による一括出力先指定が可能です。

---

## インストール

```bash
git clone https://github.com/hat47x/xlsx2json.git
cd xlsx2json
pip install openpyxl jsonschema
```

---

## 使い方


```bash
python xlsx2json.py [INPUT1 ...] [--output-dir OUT_DIR] [--schema SCHEMA_FILE] [--keep-empty] [--prefix PREFIX] [--log-level LEVEL]
```

### 引数一覧


| オプション                | 説明                                                                                                 |
|:-------------------------|:----------------------------------------------------------------------------------------------------|
| `INPUT1 ...`             | 変換対象のファイルまたはフォルダ（.xlsx）。フォルダ指定時は直下の `.xlsx` ファイルを対象とします。         |
| `-o, --output-dir`       | 一括出力先フォルダを指定。省略時は各入力ファイルと同じディレクトリの `output-json/` に出力されます。     |
| `-s, --schema`           | JSON Schema ファイルを指定。バリデーションとキー順序の整理に使用されます。                           |
| `--keep-empty`           | 空のセル値も JSON に含めます（デフォルトでは空値を除去）。                                         |
| `--prefix PREFIX`        | Excel 名前付き範囲のプレフィックスを指定（デフォルト: `json.`）。                                   |
| `--log-level LEVEL`      | ログレベルを指定（`DEBUG`/`INFO`/`WARNING`/`ERROR`/`CRITICAL`、デフォルト: `INFO`）。               |

---

## 処理の流れ

1. 入力パスから `.xlsx` ファイルを収集
2. 名前付き範囲（`json.*`）を読み込み、ネスト構造の辞書・リストを生成
3. （デフォルト）空の値を除去（`--keep-empty` オプションで保持可能）
4. （オプション）Schema バリデーションとエラーログ出力
5. （オプション）Schema の `properties` 順にキーを整理
6. JSON ファイルとして書き出し

---

## 名前付き範囲の命名規則

### 基本ルール

1. **プレフィックス**: 必ず `json.` で始めてください。ツールはこのプレフィックスを検知して処理対象とします。
2. **階層表現**: プレフィックスの後にドット `.` 区切りでキー階層を指定します。
   - 例: `json.customer.name` → `{ "customer": { "name": ... } }`
3. **配列指定**: 数字のみのセグメントは 1 始まりのインデックスとして扱い、配列要素を作成します。
   - 例: `json.items.1` → `items[0]` に値を設定
   - ネスト配列: `json.parent.2.1` → `{ "parent": [[...],[value,...]] }`

### キー名の制約
Excelの「名前の定義」に使える文字種のみ使用できます。
特に全角記号をキー名に使用する際は、実際に使用できる文字種であるかを確認したうえでご利用ください。
  - "。", "・", "＿" は OKですが "、", "！", "～", "／", "（", "）"など慣用的に項目名に用いる多くの文字が使用できない点に注意

4. **対応文字**: Unicode 文字をサポート（日本語の漢字・ひらがな・カタカナ、英数字、アンダースコア `_` 等）
5. **大文字・小文字**: 区別されます
6. **禁止文字**: ドット `.` とスペースは区切り文字のため、キー名に含めないでください
7. **最大長**: Excel の仕様上、名前の長さは 255 文字以内です。階層が深すぎると管理が困難になるため、適切な設計を推奨します

### 注意事項

8. **重複禁止**: 同じシート内、またはブック全体で同名の名前付き範囲を複数定義しないでください
9. **命名の一貫性**: チーム開発では命名規則を統一することをお勧めします

### サンプル例

```
json.user.id              → { "user": { "id": ... } }
json.order.details.price  → { "order": { "details": { "price": ... } } }
json.tags.1               → { "tags": [value1, ...] }
json.tags.2               → { "tags": [value1, value2, ...] }
json.orders.1.items.3.name → { "orders": [{ "items": [null, null, { "name": ... }] }] }
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

## 記号ワイルドカード対応の具体的な利用方法

本ツールは、Excelの「名前の定義」で使えない記号や一部の日本語記号をJSONのキー名に使いたい場合、アンダーバー（`_`）を1文字ワイルドカードとして利用できます。これにより、Excel上で定義できないキー名も、JSON Schemaの項目名にユニークにマッチすれば自動で置換されて出力されます。

### 仕組み

- Excelの名前付き範囲で「json.」以降のキー名にアンダーバー（`_`）を使うと、その位置が「任意の1文字」としてSchemaのプロパティ名にマッチします。
- マッチするSchemaのプロパティが1つだけの場合、そのプロパティ名に自動で置換されます。
- 複数マッチまたはマッチしない場合は、元のキー名のまま出力されます。

### 例

#### 1. JSON Schema例

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

#### 2. Excelでの名前付き範囲の定義例

| Excelの名前定義         | 対応するJSONキー   | 備考                         |
|-------------------------|-------------------|------------------------------|
| json.user_name          | user_name         | そのまま一致                |
| json.user_group         | user／group       | ワイルドカードによるマッチング |
| json.user_              | user！, user？    | 複数マッチ時は置換されない   |

#### 3. 注意点

- アンダーバーの数だけ1文字ずつ任意の文字にマッチします。
- マッチするSchemaプロパティが複数ある場合は、Excelで指定した名前がそのまま使われます（自動置換されません）。
- ワイルドカード置換は各階層ごとに適用されます。
- 置換の有無やマッチ状況は、`--log-level DEBUG` で実行するとログで確認できます。

#### 4. 実行例

```bash
# 例: user-名, user。名 などExcelで直接定義できないキー名をワイルドカードで指定
python xlsx2json.py sample.xlsx --schema sample-schema.json

# ログでマッチ状況を確認したい場合
python xlsx2json.py sample.xlsx --schema sample-schema.json --log-level DEBUG
```

---

## 使用例

### 基本的な使用

```bash
# 単一ファイルの変換
python xlsx2json.py data.xlsx

# 複数ファイルの変換
python xlsx2json.py file1.xlsx file2.xlsx

# フォルダ内のすべての .xlsx ファイルを変換
python xlsx2json.py ./data_folder/
```

### オプション付きの使用

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
python xlsx2json.py ./data/ --output-dir ./output/ --schema schema.json --keep-empty --prefix myprefix. --log-level DEBUG
```

---

## ライセンス

MIT
