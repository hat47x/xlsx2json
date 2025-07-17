# xlsx2json

Excel の名前付き範囲（`json.*`）を読み込んで JSON に変換し、ファイル出力するコマンドラインツールです。

---

## 特徴

- **名前付き範囲の自動パース**  
  セル名に `json.` プレフィックスを付けるだけで、自動的に JSON 形式に変換。
- **JSON Schema サポート**  
  `--schema` オプションで JSON Schema を指定し、内容のバリデーション＆キー順リオーダーが可能。
- **エラーログ出力**  
  スキーマ違反があった場合は `<basename>.error.log` に詳細なエラーを出力。
- **柔軟な出力先設定**  
  個別フォルダ、または `--output-dir` で一括指定可能。

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
python xlsx2json.py [INPUT1 ...] [--output-dir OUT_DIR] [--schema SCHEMA_FILE]
```

### 引数一覧

| オプション                | 説明                                                         |
|:-------------------------|:------------------------------------------------------------|
| `INPUT1 ...`             | 変換対象のファイルまたはフォルダ（.xlsx）。フォルダ指定時は直下の `.xlsx` を対象。 |
| `-o, --output-dir`       | 一括出力先フォルダを指定。省略時は各入力ファイル隣の `output-json/` に出力。         |
| `-s, --schema`           | JSON Schema ファイルを指定。バリデーション＆キー順ソートに使用。                     |

---

## 処理の流れ

1. 入力パスから `.xlsx` ファイルを収集
2. 名前付き範囲（`json.*`）を読み込み、ネスト構造の dict/list を生成
3. （オプション）Schema バリデーション＆エラーログ出力
4. （オプション）Schema の `properties` 順にキーをソート
5. JSON ファイルとして書き出し

---

## 名前付き範囲の命名規則

1. **プレフィックス**: 必ず `json.` で始める。ツールはこのプレフィックスを検知して処理対象とします。
2. **階層表現**: プレフィックスの後にドット `.` 区切りでキー階層を指定。
   - 例: `json.customer.name` → `{ "customer": { "name": ... } }`
3. **配列指定**: 数字のみのセグメントは 1 始まりのインデックスとして扱い、配列要素を作成。
   - 例: `json.items.1` → `items[0]` に値
   - ネスト配列: `json.parent.2.1` → `{ "parent": [[...],[value,...]] }`
4. **キー名の形式**:
   - Unicode 文字をサポート（日本語の漢字・ひらがな・カタカナ、英数字、アンダースコア `_` 等）
   - 大文字・小文字は区別されます
5. **禁止文字**: ドット `.` とスペースは区切り文字のためキー名に含めないでください。
6. **最大階層数**: Excel の仕様上、名前の長さは 255 文字以内。階層が深すぎると管理が困難になるため、適切な設計を推奨。
7. **重複禁止**: 同じシート内、またはブック全体で同名の名前付き範囲を複数定義しないこと。
8. **サンプル例**:
   - 単一値: `json.user.id`
   - ネストオブジェクト: `json.order.details.price`
   - 配列: `json.tags.1`, `json.tags.2`, ... → `"tags": [value1, value2, ...]`
   - 複合: `json.orders.1.items.3.name`

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

## ライセンス

MIT
```
