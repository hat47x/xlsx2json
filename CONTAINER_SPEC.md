# コンテナ仕様書

## 基本概念

宣言的コンテナ定義による Excel 繰り返し構造の自動処理

## 新仕様による不整合解消

以下の新仕様により、既存仕様の不整合が解消されます：

### 旧仕様の問題点
1. **`range`パラメータの重複**: キーとrangeパラメータが同じセル範囲を指すため冗長
2. **`items`配列の冗長性**: セル名から自動識別可能なため不要
3. **`labels`の複雑性**: 構造検証の複雑化により廃止
4. **`offset`の曖昧性**: 親子関係の位置計算が不明確

### 新仕様による解決
1. **キーによる範囲特定**: Excelセル範囲名をキーとして直接使用
2. **セル名からの自動項目抽出**: 既存セル名を解析して構成要素を自動識別
3. **罫線による繰り返し判定**: 罫線パターンによる客観的な構造識別
4. **相対位置による階層計算**: 第1要素の相対位置から動的計算

## 設定

config.yaml に `containers` セクションを追加：

```yaml
containers:
  json.orders: {}  # 親要素はデフォルト値で十分
  json.orders.1:
    direction: row
    increment: 5
  json.orders.1.items.1:
    direction: row
    increment: 2
  json.tree_data: {}  # 親要素はデフォルト値で十分
  json.tree_data.lv1.1:
    direction: row
    increment: 3
  json.tree_data.lv1.1.lv2.1:
    direction: column
    increment: 1
  json.tree_data.lv1.1.lv2.1.lv3.1: {}  # incrementのみ省略（direction=row, increment=0）

transform:
  - "json.orders.*.date=function:date:parse_japanese_date"
  - "json.orders.*.amount=function:math:parse_currency"
    "json.orders.*.items.*.unit_price=function:math:parse_currency",
    "json.tree_data.lv1.*.seq=function:math:parse_number",
    "json.tree_data.lv1.*.lv2.*.seq=function:math:parse_number",
    "json.tree_data.lv1.*.lv2.*.lv3.*.seq=function:math:parse_number"
  ]
```

## スキーマ

### コンテナ定義
```json
{
  "direction": "row|column",
  "increment": 数値
}
```

**パラメータ説明:**
- `direction`: データの繰り返し方向（"row"=行方向, "column"=列方向）（デフォルト: "row"）
- `increment`: 繰り返し方向への増分（デフォルト: 0）

**重要な変更点:**
- ~~`range`~~: キーと重複するため廃止
- ~~`items`~~: セル名から自動識別されるため廃止
- ~~`labels`~~: 複雑性により廃止
- ~~`offset`~~: 相対位置計算により廃止

### キーの構成規則

1. **コンテナキー形式**: `"json.{範囲名}.{インデックス}"`
   - 例: `"json.orders.1"`, `"json.tree_data.lv1.1"`

2. **範囲指定**: キーのExcelセル範囲名で範囲を特定
   - `"json.orders.1"` → Excelの名前付き範囲 `json.orders.1` を参照

3. **繰り返し処理**: 末尾数値（例: `.1`）を基準に次要素を動的生成
   - `json.orders.1` → `json.orders.2`, `json.orders.3`, ...

4. **階層構造**: ドット区切りによる親子関係
   - 親: `json.orders.1`
   - 子: `json.orders.1.items.1`

5. **親要素による範囲制限（オプション）**: 
   - 親要素（例: `"json.orders"`）でセル範囲の上限を事前指定可能
   - 罫線判定によるオーバーラン防止のため
   - 子要素は親要素のセル範囲を超えることができない

### 省略時の動作

- **`direction`省略**: デフォルト値 `"row"` を使用
- **`increment`省略**: デフォルト値 `0` を使用（繰り返しなし）
- **両方省略**: 単一範囲指定として動作（階層構造定義で曖昧さ回避に活用）

## 動作

### 自動セル名識別・動的生成

1. **セル名の自動識別**:
   - キー（例: `"json.orders.1"`）直下の全セル名を自動識別
   - 例: `json.orders.1.no`, `json.orders.1.date`, `json.orders.1.items`, ...

2. **相対位置の特定**:
   - キーのセル範囲先頭と第1要素項目の相対位置を計算
   - 例: `json.orders.1=A5:E10`, `json.orders.1.no=B6` → 相対位置(1,1)

3. **動的セル名生成**:
   - `direction`、`increment`により2以降の位置を計算
   - 計算された位置に動的にセル名を生成
   - 例: `json.orders.2.no=B11`, `json.orders.3.no=B16`, ...

### 罫線による繰り返し判定

**基本原則**: 罫線で四角形領域が描画されている場合のみ「繰り返し」として識別

1. **繰り返し要素数の計算**:
   - `direction`と`increment`の方向に罫線で囲まれた四角形を検索
   - 連続する四角形の数を繰り返し要素数として特定

2. **四角形判定条件**:
   - 四方が罫線で囲まれている
   - `increment`間隔で同様の四角形パターンが継続
   - 四角形内にセル名が配置されている

3. **非繰り返し要素**:
   - 罫線で囲まれない要素は繰り返しコンテナとして識別されない
   - 単一範囲として処理される

### 階層構造の位置補正

**親子階層がある場合の計算補正**:

1. **子要素数の動的計算**:
   - 親コンテナ（例: `json.orders.1`）が子要素（例: `json.orders.1.items.1`）を持つ場合
   - 子階層の実際の要素数を動的に計算（最低1件は必須、2件目以降は不定）
   - 子要素の総高さ・幅を親コンテナの次要素位置補正に使用

2. **incrementの正確な適用**:
   - `increment`は子要素数を意識せず、Excelファイル上の位置関係のみを指定
   - 子要素の動的サイズと`increment`を組み合わせて次要素位置を計算
   - 子要素が空欄の場合でも入力欄は1件分確保される（セル名生成のため必須）

3. **補正計算の例**:
   ```
   親: json.orders.1 (A5:E10, increment=5)
   子: json.orders.1.items.1 (C6:E8, increment=2, 実際の要素数=3)
   
   通常の次要素位置: A10 (5行増分)
   子要素による補正: 実際の子要素数3 × increment2 = 6行
   実際の次要素位置: A16 (5 + 6 + 次の開始マージン)
   ```

4. **直近親要素参照による範囲制限**:
   - 親子孫階層がある場合、**最も直近の親要素**を参照して範囲制限を適用
   - 例: `json.orders.1.items.1.details.1` の場合、`json.orders.1.items.1` を直近親として参照
   - 階層を飛び越えた参照（孫→祖父）は行わず、常に1つ上の階層のみを参照
   - フォールバック機能: 直近親が未定義の場合のみ、上位階層の親要素を検索

5. **親要素による範囲制限**:
   - 直近親要素（例: `json.orders.1`）が定義されている場合、その範囲を上限とする
   - 罫線判定による意図しないオーバーランを防止
   - 子要素は直近親要素のセル範囲を超えることができない

6. **子要素の必須要件**:
   - 子要素の入力欄は最低1件分必須（セル名生成のため）
   - 0件表現は全入力欄を空欄にすることで実現（現行ロジック対応済）
   - 動的な子要素数計算は1件から開始

## 汎用解析アルゴリズム

### 統一解析原則

**すべてのコンテナタイプ（テーブル、カード、ツリー等）を単一アルゴリズムで処理**

1. **矩形領域検出**:
   - 基準要素の矩形境界を特定
   - 四方向の罫線有無は問わない（部分罫線も対応）
   - セル名配置による論理的境界も考慮

2. **繰り返しパターン判定**:
   - `direction` + `increment` による次要素位置を計算
   - 計算位置にセル名または入力欄が存在するかチェック
   - 継続条件：セル名パターンの一致または親要素範囲内

3. **終了条件**:
   - セル名が存在しない
   - 親要素範囲を超える
   - 1000要素制限到達（無限ループ防止）

### 解析不能なケース

**この仕様では以下のExcel構造を解析できません:**

1. **不規則構造**:
   - `increment`で予測できない不規則な配置
   - セル名命名規則に従わない構造

2. **動的レイアウト**:
   - 要素ごとに異なるサイズの繰り返し
   - 条件付きで表示・非表示が変わる項目

3. **グラフィカル要素**:
   - 図形オブジェクトとセルの混合レイアウト
   - 背景色のみによる領域区分
   - インデントのみによる階層表現

4. **動的レイアウト構造**:
   - 要素ごとに異なるサイズの繰り返し
    - ただし子要素の増減による親要素のサイズの増減については解析可
   - 条件付きで表示・非表示が変わる項目
   - 可変列数のテーブル構造

5. **グラフィカル要素混合**:
   - 図形オブジェクトとセルの混合レイアウト
   - 画像埋め込みによるレイアウト破綻
   - マクロによる動的レイアウト変更

**回避策:**
- 不規則構造は手動セル名定義で対応
- 罫線なし構造には明示的な境界罫線を追加
- 複雑な結合セルは単純な矩形レイアウトに変更

## 例

### シンプルテーブル
```json
"json.orders.1": {
  "direction": "row",
  "increment": 5
}
```

**動作説明:**
1. **セル範囲**: Excelの名前付き範囲 `json.orders.1` を参照
2. **セル名自動識別**: `json.orders.1.id`, `json.orders.1.date`, `json.orders.1.amount` を自動検出
3. **罫線判定**: 四角形領域の有無で繰り返し要素数を特定
4. **動的生成**: `json.orders.2.id`, `json.orders.3.id`... （5行ずつ増分）

### ネスト構造（YAML・デフォルト値最適化）
```yaml
containers:
  json.orders: {}  # 親要素（direction=row, increment=0）
  json.orders.1:
    increment: 8   # direction=row（デフォルト）
  json.orders.1.items.1:
    increment: 2   # direction=row（デフォルト）
  json.orders.1.items.1.details.1:
    increment: 1   # direction=row（デフォルト）
  json.tree_data: {}  # 親要素（direction=row, increment=0）
  json.tree_data.lv1.1:
    increment: 4   # direction=row（デフォルト）
  json.tree_data.lv1.1.lv2.1:
    direction: column
    increment: 1
  json.tree_data.lv1.1.lv2.1.lv3.1: {}  # increment=0（デフォルト）
```

**特徴:**
- 親要素はデフォルト値で十分（{}で省略）
- 繰り返しがない要素もデフォルト値活用
- **直近親要素参照**: `json.orders.1.items.1.details.1` は `json.orders.1.items.1` を直近親として参照
- **親要素による範囲制限**: `json.orders`, `json.tree_data` のように親要素を指定することで、`json.orders.1`のような子要素の繰り返し範囲の上限を指定可（省略も可）

**自動生成セル名:**
- `json.orders.1.date`, `json.orders.2.date`, `json.orders.3.date`...
- `json.orders.1.items.1.product`, `json.orders.1.items.2.product`...
- `json.tree_data.lv1.1.name`, `json.tree_data.lv1.2.name`...
- `json.tree_data.lv1.1.lv2.1.name`, `json.tree_data.lv1.1.lv2.2.name`...

**位置補正の例（直近親要素参照による動的計算）:**
```
範囲制限階層: json.orders (A1:H50) - 最上位範囲制限
直近親要素: json.orders.1 (A5:H12, increment=8) - 主コンテナ
子要素: json.orders.1.items.1 (C7:F9, increment=2, 実際の要素数=3)
孫要素: json.orders.1.items.1.details.1 (D8:E8, increment=1, 実際の要素数=2)

孫要素の直近親参照: json.orders.1.items.1 (子要素を直近親として参照)
孫要素の動的サイズ計算: 2要素 × increment1 = 2行
子要素の動的サイズ計算: (3要素 × increment2) + 孫要素補正2 = 8行
通常の次要素位置: A13 (8行増分)
実際の次要素位置: A21 (5 + 8 + 8)

直近親範囲制限チェック: 各要素が直近親のセル範囲内であることを確認
```

**JSON出力:**
```json
{
  "orders": [
    {
      "date": "2024-01-01",
      "amount": 1000,
      "items": [
        {"product": "商品A", "qty": 2},
        {"product": "商品B", "qty": 1},
        {"product": "商品C", "qty": 3}
      ]
    },
    {
      "date": "2024-01-02", 
      "amount": 2000,
      "items": [...]
    }
  ],
  "tree_data": {
    "lv1": [
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
}
```

### カードレイアウト（新仕様）
```json
"json.customers.1": {
  "direction": "column",
  "increment": 10
}
```

**動作説明:**
1. **セル範囲**: `json.customers.1` 名前付き範囲を参照
2. **セル名自動識別**: `json.customers.1.name`, `json.customers.1.address`, `json.customers.1.phone` を自動検出
3. **罫線判定**: 四角形カードレイアウトの有無で要素数を特定
4. **動的生成**: 列方向に10列ずつ増分でカード配置

## 統合

### コマンドラインオプション対応 (--container)

#### 書式
```bash
--container '{"json.orders.1": {"direction": "row", "increment": 5}}'
```

#### 仕様
- コマンドライン指定 > 設定ファイル内の `containers`
- 同じキーの定義が重複した場合、後から指定したものが優先
- 複数の `--container` オプションを統合処理
- キー形式検証: `json.{範囲名}.{数値}` 必須

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
- **YAML解析**: YAML文字列解析
- **設定統合**: CLI優先でマージ

```python
def parse_container_args(container_args, config_containers=None):
    combined_containers = config_containers.copy() if config_containers else {}
    for container_arg in container_args:
        container_def = yaml.safe_load(container_arg)
        combined_containers.update(container_def)
    return combined_containers
```

### 最小実装要件

1. **セル名自動識別**:
   - キー直下のセル名を自動抽出
   - 相対位置から次要素のセル名を動的生成

2. **汎用解析エンジン**:
```python
def analyze_container_elements(container_key, direction='row', increment=0):
    """コンテナ要素数を動的に特定（最小実装）"""
    
    base_range = get_excel_named_range(container_key)
    element_count = 1  # 第1要素は必須
    current_pos = base_range
    
    while element_count < 1000:  # 無限ループ防止
        # 次要素位置を計算
        next_pos = calculate_next_position(current_pos, direction, increment)
        
        # セル名または入力欄の存在確認
        if not has_cell_names_at_position(next_pos):
            break
        
        # 親要素範囲内かチェック
        parent_limit = get_parent_range_limit(container_key)
        if parent_limit and is_outside_range(next_pos, parent_limit):
            break
        
        element_count += 1
        current_pos = next_pos
    
    return element_count

def generate_cell_names(container_key, element_count, direction='row', increment=0):
    """動的セル名生成（最小実装）"""
    
    # 第1要素のセル名を自動識別
    base_cells = get_cell_names_in_range(container_key)
    relative_positions = calculate_relative_positions(base_cells)
    
    # 2以降の要素のセル名を生成
    generated = {}
    for i in range(1, element_count + 1):
        for field, rel_pos in relative_positions.items():
            pos = calculate_position(i, direction, increment, rel_pos)
            cell_name = f"{container_key.replace('.1', f'.{i}')}.{field}"
            generated[pos] = cell_name
    
    return generated

def extract_field_name(cell_name, container_key):
    """セル名からフィールド名を抽出"""
    # "json.orders.1.date" → "date"
    # "json.orders.1.items.1.product" → "items.1.product" (子階層の場合)
    
    parts = cell_name.split('.')
    container_parts = container_key.split('.')
    
    # コンテナキー部分を除去してフィールド部分を取得
    field_parts = parts[len(container_parts):]
    
    return '.'.join(field_parts)

def generate_cell_name(container_key, element_index, field_name):
    """動的セル名生成"""
    # "json.orders.1" + element_index=2 + "date" → "json.orders.2.date"
    
    base_parts = container_key.split('.')
    base_parts[-1] = str(element_index)  # 末尾の数値を置換
    
    if field_name:
        return '.'.join(base_parts + field_name.split('.'))
    else:
        return '.'.join(base_parts)
```

## 最小実装のポイント

### 1. **YAML統一とデフォルト値活用**
- 設定書式をYAMLに統一
- 親要素は`{}`でデフォルト値を活用
- 繰り返しなしの要素もデフォルト値で簡素化

### 2. **汎用解析アルゴリズム**
- テーブル・カード・ツリー等の類型分岐を廃止
- 単一アルゴリズムですべてのコンテナタイプに対応
- 罫線パターンに依存しない汎用的な検出

### 3. **キー形式検証の廃止**
- 複雑なキー形式検証を除去
- ユーザーの自由度を最大限確保

### 4. **最小限のコード量**
- 要件実現に必要最小限の実装
- コード量に対するユーザー貢献を最大化

この仕様により、実装コストを最小限に抑えつつ、ユーザーの要件を最大限満たす効率的なコンテナシステムが実現されます。
    
```python
    # 親要素の場合
    if re.match(r'^json\.[^.]+$', container_key):
        # increment=0であることを推奨
        config = all_containers[container_key]
        increment = config.get('increment', 0)
        if increment != 0:
            logger.warning(f"親要素のincrementは0推奨: {container_key} increment={increment} ({source})")
    
    # 子要素の場合
    elif re.match(r'^json\.[^.]+\.\d+', container_key):
        # 対応する親要素が存在するかチェック
        parent_key = extract_parent_key_from_child(container_key)
        if parent_key and parent_key not in all_containers:
            logger.info(f"親要素未定義（オプション）: {parent_key} for {container_key} ({source})")

def extract_parent_key_from_child(child_key):
    """子要素キーから親要素キーを抽出"""
    # "json.orders.1.items.1" → "json.orders"
    parts = child_key.split('.')
    if len(parts) >= 3:
        parent_parts = []
        for part in parts[1:-1]:  # 'json'と末尾数値を除外
            if not part.isdigit():
                parent_parts.append(part)
        
        if parent_parts:
            return 'json.' + '.'.join(parent_parts)
    
    return None

def validate_direction_parameter(container_key, config, source):
    """direction パラメータの検証"""
    direction = config.get('direction', 'row')
    if direction not in ['row', 'column']:
        raise ValueError(f"無効なdirection: {direction} (row|column のみ許可) - {container_key} ({source})")

def validate_increment_parameter(container_key, config, source):
    """increment パラメータの検証"""
    increment = config.get('increment', 0)
    if not isinstance(increment, int) or increment < 0:
        raise ValueError(f"無効なincrement: {increment} (0以上の整数必須) - {container_key} ({source})")

def check_deprecated_parameters(container_key, config, source):
    """廃止パラメータの警告"""
    deprecated_params = ['range', 'items', 'labels', 'offset']
    found_deprecated = []
    
    for param in deprecated_params:
        if param in config:
            found_deprecated.append(param)
    
    if found_deprecated:
        logger.warning(f"廃止パラメータ検出: {found_deprecated} - {container_key} ({source})")
        logger.warning("新仕様では以下が変更されました:")
        logger.warning("- range: キーから自動識別")
        logger.warning("- items: セル名から自動抽出")
        logger.warning("- labels: 罫線による構造判定に変更")
        logger.warning("- offset: 相対位置による自動計算")

def validate_cli_containers_new_spec(container_args):
    """新仕様対応のCLI引数検証"""
    for i, container_arg in enumerate(container_args):
        try:
            container_def = json.loads(container_arg)
        except json.JSONDecodeError as e:
            raise ValueError(f"CLI引数{i+1}: 無効なJSON形式 - {e}")
        
        validate_container_config_new_spec(container_def, f"CLI引数{i+1}")
```

## YAML設定例

### 基本設定例
```yaml
containers:
  json.orders: {}  # 親要素（デフォルト値）
  json.orders.1:
    increment: 5   # direction=row（デフォルト）
  json.orders.1.items.1:
    increment: 2   # direction=row（デフォルト）
```

### 省略設定例
```yaml
containers:
  json.single_range.1: {}  # デフォルト値のみ
  json.card_layout: {}     # 親要素（デフォルト値）
  json.card_layout.1:
    direction: column
    increment: 8
  json.hierarchy_def.1: {} # increment=0（デフォルト）
```

### CLIオプション例
```bash
# 基本使用（YAML文字列）
--container 'json.orders: {}' \
--container 'json.orders.1: {increment: 5}'

# 複数コンテナ指定
--container 'json.orders.1: {increment: 5}' \
--container 'json.orders.1.items.1: {increment: 2}'

# デフォルト値使用
--container 'json.simple.1: {}'
```

## 新仕様実装のポイント

### 1. **キーとExcel範囲の直接対応**
- 従来の `range` パラメータを廃止
- キー自体がExcelの名前付き範囲を指定
- 設定の重複と冗長性を解消

### 2. **自動セル名識別による簡素化**
- 従来の `items` 配列を廃止
- 既存のセル名を解析して構成要素を自動抽出
- 手動設定の手間を大幅削減

### 3. **罫線による客観的構造判定**
- 従来の `labels` による複雑な検証を廃止
- 四角形の罫線パターンによる明確な判定基準
- 人間の視覚的判断と一致する解析結果

### 4. **相対位置による階層計算**
- 従来の `offset` による曖昧な指定を廃止
- 第1要素の相対位置から自動計算
- 子階層の要素数による位置補正を実装

### 5. **親要素による範囲制限機能**
- オプションで親要素のセル範囲制限を追加
- 罫線判定による意図しないオーバーラン防止
- 子要素の動的計算における安全性向上

### 6. **子要素の動的サイズ計算**
- 子要素数を意識しないincrementの適用
- 実際のExcelファイル構造に基づく動的計算
- 最低1件確保による空データ表現の対応

この新仕様により、コンテナ設定が大幅に簡素化され、Excel構造との一致性が向上し、設定の保守性が飛躍的に改善されます。また、親要素による範囲制限と子要素の動的計算により、より正確で安全な構造解析が可能になります。
