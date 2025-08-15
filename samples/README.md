# サンプルファイル

このフォルダには xlsx2json の使用例として以下のファイルが含まれています：

## ファイル一覧

- **sample.xlsx**: サンプルのExcelファイル（セル名を含む）
- **config.json**: 設定ファイルのサンプル
- **schema.json**: JSON Schemaのサンプル
- **transform.py**: データ変換関数のサンプル集（ユーザ定義関数の例）

## 使用方法

```bash
# 基本的な使用例
python ../xlsx2json.py samples/sample.xlsx --output-dir output

# スキーマを使用してバリデーション
python ../xlsx2json.py samples/sample.xlsx --schema samples/schema.json --output-dir output

# 設定ファイルを使用
python ../xlsx2json.py samples/sample.xlsx --config samples/config.yaml --output-dir output
```

## transform.py の関数一覧

### 🔤 文字列変換
- **`csv(value)`** - CSV文字列を配列に分割
- **`lines(value)`** - 改行区切りの文字列を配列に分割  
- **`words(value)`** - 空白区切りの文字列を配列に分割

### 📊 配列・行列操作
- **`column(data, index=0)`** - 指定列を抽出
- **`sum_col(data, index=0)`** - 指定列の合計を計算
- **`flip(data)`** - 行と列を入れ替え（転置）
- **`clean(data)`** - 空でない行のみを残す

### 🔢 数値計算
- **`total(data)`** - 全要素の合計
- **`avg(data)`** - 数値要素の平均

### 🛠️ 便利関数
- **`normalize(value)`** - 文字列を正規化（トリム・全角半角変換・置換など）
- **`parse_json(value)`** - JSON文字列を解析
- **`upper(value)`** - 大文字に変換
- **`lower(value)`** - 小文字に変換

### 使用例

```json
{
  "transform": [
    "json.data=function:samples/transform.py:csv",
    "json.matrix=range:A1:C3:function:samples/transform.py:total",
    "json.names=function:samples/transform.py:normalize"
  ]
}
```
