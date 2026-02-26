# ガントチャート PPTX 生成ツール

Excelファイルからガントチャートの PowerPoint スライドを自動生成するツールです。

## セットアップ

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

> 2回目以降は `source venv/bin/activate` だけでOKです。

## クイックスタート

```bash
# Excelから生成
python main.py --input sample_data.xlsx --output gantt.pptx

# 四半期モードで生成
python main.py -i sample_data.xlsx -o gantt.pptx --mode quarterly

# 期間を指定
python main.py -i data.xlsx -o gantt.pptx --start 2025-04 --end 2027-03
```

## コマンドオプション一覧

| オプション | 短縮 | 説明 | デフォルト |
|--|--|--|--|
| `--input` | `-i` | 入力Excelファイル（必須） | — |
| `--output` | `-o` | 出力PPTXファイル | `output.pptx` |
| `--sheet` | `-s` | Excelのシート名 | 最初のシート |
| `--start` | | 表示開始月 `YYYY-MM` | 自動 |
| `--end` | | 表示終了月 `YYYY-MM` | 自動 |
| `--mode` | `-m` | カレンダー表示モード | `auto` |
| `--external-label` | | 矢羽テキスト外出し閾値（月数） | `3` |

## カレンダー表示モード（`--mode`）

| モード | ヘッダー構成 | 推奨期間 |
|--|--|--|
| `auto` | 月数に応じて自動選択 | — |
| `bme` | FY → 月 → B/M/E（上旬/中旬/下旬） | 〜12ヶ月 |
| `monthly` | FY → 月 | 13〜30ヶ月程度 |
| `quarterly` | FY → Q1/Q2/Q3/Q4 | 24ヶ月〜 |

**四半期の定義（日本の会計年度）：**
Q1 = 4〜6月 / Q2 = 7〜9月 / Q3 = 10〜12月 / Q4 = 1〜3月

## Excelフォーマット

1行目はヘッダー行、2行目以降がデータです。

| 列 | ヘッダー | 説明 | 例 |
|--|--|--|--|
| A | 大項目 | プロジェクト名。同じ値が続く行はグループ化 | Phase1: 企画 |
| B | 種別 | `milestone` / `period` / `task` | task |
| C | 名前 | アイテム名 | 要件定義 |
| D | 開始日 | 日付（YYYY-MM-DD またはExcel日付） | 2025-04-01 |
| E | 終了日 | 日付（milestoneは空欄可） | 2025-06-30 |

### 入力例

| 大項目 | 種別 | 名前 | 開始日 | 終了日 |
|--|--|--|--|--|
| 企画フェーズ | milestone | キックオフ | 2025-04-01 | |
| 企画フェーズ | period | 企画 | 2025-04-01 | 2025-06-30 |
| 企画フェーズ | task | 市場調査 | 2025-04-01 | 2025-04-30 |
| 企画フェーズ | task | 企画書作成 | 2025-05-01 | 2025-05-31 |
| 設計フェーズ | period | 設計 | 2025-07-01 | 2025-09-30 |
| 設計フェーズ | task | 基本設計 | 2025-07-01 | 2025-08-15 |

### ルール
- **大項目**が同じ連続行は1つのレーンにまとめられます
- **period**: カレンダー上にバー（青）で表示
- **milestone**: ひし形（赤）で表示。終了日は不要
- **task**: 矢羽（グレー）で表示。1レーン内に縦に並びます

## ファイル構成

```
ppt_guntt/
├── main.py              # CLI エントリーポイント
├── gantt_generator.py   # ガントチャート描画エンジン
├── excel_reader.py      # Excel → データ変換
├── generate_samples.py  # サンプルPPTX一括生成
├── sample_data.xlsx     # サンプルExcel
└── requirements.txt
```

## Python から直接使う場合

```python
from gantt_generator import GanttChartGenerator

data = [
    {
        "name": "設計フェーズ",
        "type": "project",
        "items": [
            {"name": "設計", "type": "period", "start": "2025-04-01", "end": "2025-06-30"},
            {"name": "完了", "type": "milestone", "date": "2025-06-30"}
        ],
        "tasks": [
            {"name": "基本設計", "start": "2025-04-01", "end": "2025-05-15"},
            {"name": "詳細設計", "start": "2025-05-01", "end": "2025-06-30"}
        ]
    }
]

gen = GanttChartGenerator(data, calendar_mode="auto")
gen.generate("output.pptx")
```
