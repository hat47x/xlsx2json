# xlsx2json

Excel の名前付き範囲（`json.*`）を読み込んで JSON 形式に変換し、ファイル出力するコマンドラインツールです。

---

## 特徴

- **名前付き範囲の自動解析**  
  セル名に `json.` プレフィックスを付けるだけで、自動的に JSON 形式に変換します。
- **JSON Schema サポート**  
  `--schema` オプションで JSON Schema を指定することで、データのバリデーションとキー順序の整理が可能です。
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
python xlsx2json.py [INPUT1 ...] [--output-dir OUT_DIR] [--schema SCHEMA_FILE] [--keep-empty]
```

### 引数一覧

| オプション                | 説明                                                         |
|:-------------------------|:------------------------------------------------------------|
| `INPUT1 ...`             | 変換対象のファイルまたはフォルダ（.xlsx）。フォルダ指定時は直下の `.xlsx` ファイルを対象とします。 |
| `-o, --output-dir`       | 一括出力先フォルダを指定。省略時は各入力ファイルと同じディレクトリの `output-json/` に出力されます。         |
| `-s, --schema`           | JSON Schema ファイルを指定。バリデーションとキー順序の整理に使用されます。                     |
| `--keep-empty`           | 空のセル値も JSON に含めます（デフォルトでは空値を除去）。                     |

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

# すべてのオプションを組み合わせ
python xlsx2json.py ./data/ --output-dir ./output/ --schema schema.json --keep-empty
```

---

## ライセンス

MIT