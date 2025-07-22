# コンテナ仕様書

## 基本概念

宣言的コンテナ定義による Excel 繰り返し構造の自動処理

## 設定

config.json に `containers` セクションを追加：

```json
{
  "containers": {
    "json.orders": {
      "range": "orders_range", 
      "direction": "row",
      "items": ["date", "customer_id", "amount", "items"],
      "labels": ["注文日", "顧客ID", "金額"]
    },
    "json.orders.1.items": {
      "offset": 3,
      "items": ["product_code", "quantity", "unit_price"],
      "labels": ["商品コード", "数量", "単価"]
    },
    "json.tree_data": {
      "range": "tree_range",
      "direction": "row",
      "items": ["name", "value", "seq"],
      "labels": ["項目名", "値"]
    },
    "json.tree_data.lv1.1.lv2": {
      "offset": 1,
      "items": ["name", "value", "seq"],
      "labels": ["項目名", "値"]
    },
    "json.tree_data.lv1.1.lv2.1.lv3": {
      "offset": 1,
      "items": ["name", "value", "seq"],
      "labels": ["項目名", "値"]
    }
  },
  "transform": [
    "json.orders.*.date=function:date:parse_japanese_date",
    "json.orders.*.amount=function:math:parse_currency",
    "json.orders.*.items.*.unit_price=function:math:parse_currency",
    "json.tree_data.lv1.*.seq=function:math:parse_number",
    "json.tree_data.lv1.*.lv2.*.seq=function:math:parse_number",
    "json.tree_data.lv1.*.lv2.*.lv3.*.seq=function:math:parse_number"
  ]
}
```

## スキーマ

### ルートコンテナ
```json
{
  "range": "range_name",
  "direction": "row|column", 
  "increment": 1,
  "items": ["field1", "field2", ...],
  "labels": ["必須ラベル1", "必須ラベル2", ...]
}
```

**パラメータ説明:**
- `range`: Excel名前付き範囲または範囲文字列（例: "A1:C10"）
- `direction`: データの配置方向（"row"=行方向, "column"=列方向）
- `increment`: 複数インスタンス間の間隔（Card型で使用）
- `items`: 各インスタンスに含まれるフィールド名の配列
- `labels`: 構造同一性確認のための必須ラベル配列（オプション、`json.*`セルは検索対象外）

### 子コンテナ
```json
{
  "offset": 3,
  "items": ["field1", "field2", ...],
  "labels": ["必須ラベル1", "必須ラベル2", ...]
}
```

**パラメータ説明:**
- `offset`: 親コンテナからの相対位置（行または列単位）
- `items`: 子インスタンスに含まれるフィールド名の配列
- `labels`: 構造同一性確認のための必須ラベル配列（オプション、`json.*`セルは検索対象外）

**用途:**
- 親コンテナ内の各インスタンスに対して、関連する子データを定義
- 明細データやサブテーブルの処理に使用
- `labels`により子構造の存在確認と同一性検証を実行

## 動作

### 罫線・構造解析による自動判定
- **Table判定**: 周囲+内部の規則的な罫線パターン
- **Card判定**: 周囲の囲み罫線、内部は空白多め
  - **フラット型**: 単純なラベル-値ペア
  - **階層型**: インデントによる親子関係（囲み罫線から判断）
- **多段組み対応**: `increment`間隔での罫線確認
- **構造同一性確認**: `labels`配列で指定された必須ラベルの存在確認（`json.*`セルは値用のため除外）

### セル名生成
- テンプレート: `json.orders.1.date`
- 自動展開: `json.orders.1.date`, `json.orders.2.date`, `json.orders.3.date`...

### 座標計算
- **ルート**: 明示的な `range` 範囲
- **子**: 親範囲 + `offset` から計算

### インスタンス検出
- **行方向**: 空行まで計数
- **列方向**: 空列まで計数
- **構造検証**: `labels`で指定されたラベルがすべて存在する範囲のみを有効なインスタンスとして判定
- **ラベル判定除外**: セル名（`json.*`）が設定されたセルはラベル検索対象から除外

### 階層解決アルゴリズム

コンテナ階層は名前パターンの解析により自動構築されます：

1. **親子関係の識別**: `json.parent.1.child` パターンから親子関係を自動検出
2. **階層レベル自動判定**: セル名の構造パターン（ドット区切りの深さ）から階層レベルを判定
3. **構造同一性確認**: `labels`配列で指定された必須ラベルの存在確認による構造検証（`json.*`セルは除外）
4. **インデックス自動生成**: 罫線解析結果に基づく動的インデックス採番
5. **シーケンス番号管理**: 各階層レベルでの順序番号の自動割り当て

**処理フロー:**
1. 基本セル名パターン（インデックス1）の定義を解析
2. 対象範囲内の罫線構造を分析して構造タイプを判定
3. `labels`で指定された必須ラベルの存在確認により構造同一性を検証（`json.*`セルは除外）
4. セル名の構造パターン（数値インデックスを除いた階層セグメント）から階層レベルの深さを判定
5. 各レベルでの項目数を計算してインデックス範囲を決定
6. 階層構造に応じたJSONオブジェクトツリーを構築

## 罫線解析パターン

### Table（テーブル）
- 周囲の完全な囲み罫線
- 内部の規則的なグリッド罫線
- ヘッダー行の下線強調
- `increment`間隔での区切り線確認

### Card（カード型）  

カード型は、1つの論理的なまとまりを持つデータブロックを扱います。

**特徴:**
- 周囲の囲み罫線による明確な境界
- 内部は空白セルを含む自由なレイアウト
- ラベル-値ペアの柔軟な配置

**階層構造の自動判定:**
- セル名の構造パターン（ドット区切りの階層深度）から階層レベルを自動識別
- インデント解析：セルのインデントレベルから親子関係を自動判定
- セル結合パターン：結合セルの配置から階層境界を特定
- 自動インデックス管理：各レベルでの項目数を動的に計算

## 例

### シンプルテーブル
```json
"json.orders": {
  "range": "orders_range",
  "direction": "row",
  "items": ["id", "date", "amount"],
  "labels": ["ID", "注文日", "金額"]
}
```
→ 罫線解析でTable判定 → ラベル存在確認 → 生成: `json.orders.1.id`, `json.orders.1.date`, `json.orders.1.amount`...

### ネスト構造（マルチレベル）
```json
"json.orders": {
  "range": "orders_range",
  "direction": "row",
  "items": ["date", "items"]
},
"json.orders.1.items": {
  "offset": 1,
  "items": ["product", "qty"]
},
"json.tree_data": {
  "range": "tree_range",
  "direction": "row",
  "items": ["name", "value", "seq"]
},
"json.tree_data.lv1.1.lv2": {
  "offset": 1,
  "items": ["name", "value", "seq"]
},
"json.tree_data.lv1.1.lv2.1.lv3": {
  "offset": 1,
  "items": ["name", "value", "seq"]
}
```

**特徴:**
- 既存のネスト構造定義と同じパターンで階層構造を表現
- セル名の構造パターン（ドット区切りの階層深度）から階層レベルを自動判定
- 特定のレベル名（`lv1`, `lv2`など）に依存せず、任意の階層構造に対応

**自動生成セル名:**
- `json.tree_data.lv1.1.name`, `json.tree_data.lv1.2.name`...
- `json.tree_data.lv1.1.lv2.1.name`, `json.tree_data.lv1.1.lv2.2.name`...
- `json.tree_data.lv1.1.lv2.1.lv3.1.name`, `json.tree_data.lv1.1.lv2.1.lv3.2.name`...

**JSON出力:**
```json
{
  "tree_data": [
    {
      "name": "ルート項目1",
      "seq": 1,
      "lv2": [
        {
          "name": "中間項目1-1", 
          "seq": 1,
          "lv3": [
            {"name": "詳細項目1-1-1", "seq": 1},
            {"name": "詳細項目1-1-2", "seq": 2}
          ]
        }
      ]
    }
  ]
}

### カードレイアウト
```json
"json.customers": {
  "range": "customers_range", 
  "direction": "column",
  "increment": 10,
  "items": ["name", "address", "phone"],
  "labels": ["顧客名", "住所", "電話番号"]
}
```
→ 罫線解析でCard判定 → ラベル存在確認 → 1枚目カード: A1-A3、2枚目カード: K1-K3


## 統合

### コマンドラインオプション対応 (--container)

#### 書式
```bash
--container '{"セル名": JSON定義}'
```

#### 仕様
- コマンドライン指定 > 設定ファイル内の `containers`
- 同じセル名の定義が重複した場合、後から指定したものが優先
- 複数の `--container` オプションを統合処理

### 変換ルール対応 (--transformのワイルドカード対応)
- ワイルドカード指定: `json.orders.*.amount=function:math:parse_currency`
- 子要素へのルール適用: `json.orders.*.items.*.price=function:math:multiply`
- 階層ワイルドカード: `json.customers.*.orders.*.date=function:date:parse`

### 既存機能との互換性
- 既存の `transform` ルールと互換
- 手動セル名定義と併用可能
- コンテナ階層から自動JSONパス生成

## 実装要件

### 0. CLIオプション対応（--container）
- **引数パース**: JSON文字列解析
- **設定統合**: CLI優先でマージ
- **複数指定対応**: 複数オプション統合処理
- **書式検証**: セル名 `json.*` プレフィックス必須

```python
def parse_container_args(container_args, config_containers=None):
    combined_containers = config_containers.copy() if config_containers else {}
    for container_arg in container_args:
        container_def = json.loads(container_arg)
        combined_containers.update(container_def)
    return combined_containers
```

### 1. コンテナとセル名の統合
- **階層構造の明示**: コンテナ定義内でセル名の階層構造を完全表現
- **位置関係の算出**: `range`, `direction`, `offset`から構造間の相対位置を計算
- **競合解決**: 手動セル名 < コンテナ生成セル名（優先、警告ログ出力）
- **セル名生成**: コンテナ定義 → 自動セル名生成

### 2. 罫線解析エンジン（セル名ベース階層解析）
```python
def calculate_hierarchy_depth(cell_name):
    """数値インデックスを除外した階層深度を計算"""
    parts = cell_name.split('.')
    hierarchy_parts = [part for part in parts if not part.isdigit()]
    return len(hierarchy_parts) - 1  # 'json'を除く

def analyze_border_structure(range_coords, cell_names):
    # セル名を持つセルのみを解析対象として特定
    target_cells = filter_cells_with_names(range_coords, cell_names)
    
    # 各セルの階層深度を計算（数値インデックス除外）
    hierarchy_depth = {
        cell: calculate_hierarchy_depth(name) 
        for cell, name in target_cells.items()
    }
    
    # 罫線による囲み構造を検出
    border_nesting = detect_border_nesting(range_coords)
    
    # 囲み数と階層深度の対応検証
    validate_structure_consistency(border_nesting, hierarchy_depth)
    
    return build_validated_structure(border_nesting, hierarchy_depth)
```

**方針:**
- **セル名限定解析**: `json.*`セル名を持つセルのみを構造解析対象とする
- **階層深度計算**: セル名の階層構造（数値インデックス除外）から階層レベルを算出
- **囲み数検証**: 罫線による囲み構造の深さを計測
- **対応関係確認**: 囲み数 = 階層深度の厳密な対応を検証
- **不整合処理**: 対応しない場合は警告ログ + 可能な範囲で補正

### 階層対応検証アルゴリズム
```python
def validate_structure_consistency(border_nesting, hierarchy_depth):
    for cell_coord, expected_depth in hierarchy_depth.items():
        actual_nesting = get_nesting_level(border_nesting, cell_coord)
        
        if actual_nesting != expected_depth:
            logger.warning(
                f"階層不整合検出: セル{cell_coord} "
                f"期待階層:{expected_depth} 罫線囲み:{actual_nesting}"
            )
            
            # 補正可能かチェック
            if can_adjust_structure(cell_coord, expected_depth, actual_nesting):
                adjust_structure_mapping(cell_coord, expected_depth)
            else:
                logger.error(f"セル{cell_coord}の階層補正不可 - 出力除外")
                exclude_from_output(cell_coord)
```

**階層対応ルール:**
- **セル名 `json.a.b.c`** → 階層深度3 → 罫線囲み3重が必要
- **セル名 `json.orders.1.items.2.name`** → 階層深度3 → 罫線囲み3重が必要（数値インデックスは階層に含まない）
- **不整合時の補正**: 可能な場合のみ構造マッピングを調整
- **補正不可時**: 当該セルを出力対象から除外

### 3. 座標計算システム（セル名対応座標特定）
```python
def calculate_coordinates(container_def, parent_range=None):
    if parent_range is None:
        # ルートコンテナ: 明示的範囲指定
        base_range = parse_range(container_def['range'])
    else:
        # 子コンテナ: 親範囲 + offset計算
        base_range = apply_offset(parent_range, container_def['offset'], 
                                 container_def.get('direction', 'row'))
    
    # 範囲内のセル名を取得
    cell_names = get_cell_names_in_range(base_range)
    
    # セル名の階層深度と罫線構造の整合性確認
    structure_validation = validate_cell_hierarchy_consistency(
        base_range, cell_names, container_def
    )
    
    if not structure_validation.is_valid:
        logger.warning(f"構造不整合: {structure_validation.errors}")
        
    return base_range, structure_validation
```

**処理手順:**
1. **基準範囲特定**: `range`指定または親範囲+offset計算
2. **セル名収集**: 範囲内の`json.*`セル名を全て取得
3. **階層深度算出**: 各セル名の階層構造（数値インデックス除外）から階層レベル計算
4. **罫線構造解析**: 範囲内の囲み罫線パターンを解析
5. **整合性検証**: 階層深度と囲み数の対応関係を確認
6. **補正または除外**: 不整合セルの処理方針決定

### 4. 繰り返し検出アルゴリズム（階層整合性重視）
```python
def detect_repetitions(start_coords, direction, labels, expected_hierarchy):
    instances = []
    current_pos = start_coords
    
    while True:
        # セル名の存在確認
        cell_names = get_cell_names_at_position(current_pos)
        if not cell_names:
            break  # セル名なし → 終了
        
        # 階層深度の整合性確認
        hierarchy_valid = validate_hierarchy_at_position(
            current_pos, cell_names, expected_hierarchy
        )
        if not hierarchy_valid:
            logger.warning(f"位置{current_pos}で階層不整合 - スキップ")
            current_pos = advance_position(current_pos, direction)
            continue
        
        # 罫線構造の確認
        border_structure = analyze_border_at_position(current_pos)
        if not matches_expected_structure(border_structure, expected_hierarchy):
            break  # 構造不一致 → 終了
        
        # ラベル検証  
        if not validate_labels(current_pos, labels):
            break  # ラベル不一致 → 終了
            
        instances.append(extract_instance_data(current_pos, cell_names))
        current_pos = advance_position(current_pos, direction)
        
    return instances
```

**終了条件:**
- **セル名未定義**: `json.*`セル名が存在しない領域
- **階層深度不整合**: セル名の階層と罫線囲み数が不一致
- **罫線構造不一致**: 期待する囲み パターンが見つからない  
- **必須ラベル欠如**: `labels`配列のラベルが不完全
- **範囲外到達**: 指定範囲の境界に達した

### 5. エラーハンドリング（階層整合性重視）
- **階層不整合検出**: セル名階層深度と罫線囲み数の不一致時
  - 警告ログ出力: `"階層不整合: セルA1の期待階層3、実際囲み2"`
  - 補正処理: 可能な場合のみ構造マッピング調整
  - 補正不可時: 該当セルを出力対象から除外
- **セル名未定義**: `json.*`パターン以外のセルは解析対象外
- **座標範囲外**: 範囲外アクセス時は安全に終了
- **設定矛盾**: 設定エラー時は詳細ログで原因を明示

### 階層整合性チェック例
```python
def check_hierarchy_consistency(cell_name, border_nesting_level):
    # 数値インデックスを除外して階層深度を計算
    parts = cell_name.split('.')
    hierarchy_parts = [part for part in parts if not part.isdigit()]
    expected_depth = len(hierarchy_parts) - 1  # 'json'を除く
    
    if border_nesting_level != expected_depth:
        return {
            'valid': False,
            'expected': expected_depth,
            'actual': border_nesting_level,
            'correctable': abs(expected_depth - border_nesting_level) <= 1
        }
    return {'valid': True}
```

### 6. 実装方針（可読性重視）
- **シンプル設計**: 複雑な最適化より分かりやすいロジック
- **確実性優先**: エラー処理を丁寧に、予期しない状況に対応
- **段階的処理**: 各ステップを独立したメソッドに分割
- **豊富なログ**: デバッグしやすい詳細ログ出力

### 7. 設定検証（起動時チェック）
```python
def validate_container_config(containers, source="config"):
    for container_name, config in containers.items():
        validate_required_params(container_name, config, source)
        validate_coordinate_consistency(container_name, config, source)
        validate_hierarchy_logic(container_name, config, source)
        validate_arrays(container_name, config, source)

def validate_cli_containers(container_args):
    for i, container_arg in enumerate(container_args):
        container_def = json.loads(container_arg)
        for cell_name in container_def.keys():
            if not cell_name.startswith('json.'):
                raise ValueError(f"セル名は'json.'で始まる必要があります: {cell_name}")
        validate_container_config(container_def, f"CLI引数{i+1}")
```
