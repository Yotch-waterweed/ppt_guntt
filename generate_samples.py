"""
3ヶ月、6ヶ月、12ヶ月、24ヶ月のガントチャートサンプルを生成するスクリプト
"""
from gantt_generator import GanttChartGenerator


def make_3month_data():
    """3ヶ月サンプル（4月〜6月）"""
    return [
        {
            "name": "TOBE設計",
            "type": "project",
            "items": [
                {"name": "キックオフ", "type": "milestone", "date": "2025-04-01"},
                {"name": "設計フェーズ", "type": "period", "start": "2025-04-01", "end": "2025-06-30"}
            ],
            "tasks": []
        },
        {
            "name": "業務要件定義",
            "type": "project",
            "items": [
                {"name": "業務要件定義", "type": "period", "start": "2025-04-01", "end": "2025-04-30"}
            ],
            "tasks": [
                {"name": "現状業務フロー整理", "start": "2025-04-01", "end": "2025-04-10"},
                {"name": "課題・要望ヒアリング", "start": "2025-04-08", "end": "2025-04-18"},
                {"name": "業務要件書作成・レビュー", "start": "2025-04-16", "end": "2025-04-30"}
            ]
        },
        {
            "name": "システム要件定義",
            "type": "project",
            "items": [
                {"name": "システム要件定義", "type": "period", "start": "2025-04-21", "end": "2025-06-15"},
                {"name": "要件FIX", "type": "milestone", "date": "2025-06-15"}
            ],
            "tasks": [
                {"name": "機能要件洗い出し", "start": "2025-04-21", "end": "2025-05-05"},
                {"name": "非機能要件定義", "start": "2025-04-28", "end": "2025-05-10"},
                {"name": "外部IF仕様検討", "start": "2025-05-06", "end": "2025-05-20"},
                {"name": "データモデル設計", "start": "2025-05-12", "end": "2025-05-31"},
                {"name": "要件定義書作成・承認", "start": "2025-05-26", "end": "2025-06-15"}
            ]
        }
    ]


def make_6month_data():
    """6ヶ月サンプル（4月〜9月）"""
    return [
        {
            "name": "TOBE設計",
            "type": "project",
            "items": [
                {"name": "キックオフ", "type": "milestone", "date": "2025-04-01"},
                {"name": "前半フェーズ", "type": "period", "start": "2025-04-01", "end": "2025-06-30"},
                {"name": "中間レビュー", "type": "milestone", "date": "2025-06-30"},
                {"name": "後半フェーズ", "type": "period", "start": "2025-07-01", "end": "2025-09-30"}
            ],
            "tasks": []
        },
        {
            "name": "業務要件定義",
            "type": "project",
            "items": [
                {"name": "業務要件定義", "type": "period", "start": "2025-04-01", "end": "2025-05-31"}
            ],
            "tasks": [
                {"name": "現状業務フロー整理", "start": "2025-04-01", "end": "2025-04-15"},
                {"name": "課題・要望ヒアリング", "start": "2025-04-10", "end": "2025-04-30"},
                {"name": "業務要件書作成・レビュー", "start": "2025-04-25", "end": "2025-05-31"}
            ]
        },
        {
            "name": "システム設計",
            "type": "project",
            "items": [
                {"name": "システム設計", "type": "period", "start": "2025-05-01", "end": "2025-07-31"},
                {"name": "設計完了", "type": "milestone", "date": "2025-07-31"}
            ],
            "tasks": [
                {"name": "基本設計", "start": "2025-05-01", "end": "2025-06-15"},
                {"name": "詳細設計", "start": "2025-06-01", "end": "2025-07-15"},
                {"name": "設計レビュー", "start": "2025-07-10", "end": "2025-07-31"}
            ]
        },
        {
            "name": "開発・テスト",
            "type": "project",
            "items": [
                {"name": "開発・テスト", "type": "period", "start": "2025-07-01", "end": "2025-09-30"},
                {"name": "リリース", "type": "milestone", "date": "2025-09-30"}
            ],
            "tasks": [
                {"name": "フロントエンド開発", "start": "2025-07-01", "end": "2025-08-31"},
                {"name": "バックエンド開発", "start": "2025-07-01", "end": "2025-08-31"},
                {"name": "結合テスト", "start": "2025-08-15", "end": "2025-09-15"},
                {"name": "受入テスト", "start": "2025-09-10", "end": "2025-09-30"}
            ]
        }
    ]


def make_12month_data():
    """12ヶ月サンプル（4月〜3月：1年度分）"""
    return [
        {
            "name": "企画フェーズ",
            "type": "project",
            "items": [
                {"name": "プロジェクト開始", "type": "milestone", "date": "2025-04-01"},
                {"name": "企画", "type": "period", "start": "2025-04-01", "end": "2025-06-30"}
            ],
            "tasks": [
                {"name": "市場調査・分析", "start": "2025-04-01", "end": "2025-04-30"},
                {"name": "企画書作成", "start": "2025-05-01", "end": "2025-05-31"},
                {"name": "経営承認", "start": "2025-06-01", "end": "2025-06-30"}
            ]
        },
        {
            "name": "設計フェーズ",
            "type": "project",
            "items": [
                {"name": "設計", "type": "period", "start": "2025-06-01", "end": "2025-09-30"},
                {"name": "設計完了", "type": "milestone", "date": "2025-09-30"}
            ],
            "tasks": [
                {"name": "要件定義", "start": "2025-06-01", "end": "2025-07-15"},
                {"name": "基本設計", "start": "2025-07-01", "end": "2025-08-15"},
                {"name": "詳細設計", "start": "2025-08-01", "end": "2025-09-30"}
            ]
        },
        {
            "name": "開発フェーズ",
            "type": "project",
            "items": [
                {"name": "開発", "type": "period", "start": "2025-09-01", "end": "2026-01-31"}
            ],
            "tasks": [
                {"name": "バックエンド実装", "start": "2025-09-01", "end": "2025-11-30"},
                {"name": "フロントエンド実装", "start": "2025-10-01", "end": "2025-12-31"},
                {"name": "API連携開発", "start": "2025-11-01", "end": "2026-01-31"}
            ]
        },
        {
            "name": "テスト・移行",
            "type": "project",
            "items": [
                {"name": "テスト・移行", "type": "period", "start": "2026-01-01", "end": "2026-03-31"},
                {"name": "本番リリース", "type": "milestone", "date": "2026-03-31"}
            ],
            "tasks": [
                {"name": "統合テスト", "start": "2026-01-01", "end": "2026-01-31"},
                {"name": "性能テスト", "start": "2026-01-15", "end": "2026-02-15"},
                {"name": "データ移行", "start": "2026-02-01", "end": "2026-02-28"},
                {"name": "本番移行・運用開始", "start": "2026-03-01", "end": "2026-03-31"}
            ]
        }
    ]


def make_24month_data():
    """24ヶ月サンプル（2025年4月〜2027年3月：2年度分）"""
    return [
        {
            "name": "Phase1: 企画・構想",
            "type": "project",
            "items": [
                {"name": "PJ開始", "type": "milestone", "date": "2025-04-01"},
                {"name": "企画・構想", "type": "period", "start": "2025-04-01", "end": "2025-09-30"}
            ],
            "tasks": [
                {"name": "現状分析", "start": "2025-04-01", "end": "2025-06-30"},
                {"name": "構想策定", "start": "2025-06-01", "end": "2025-08-31"},
                {"name": "RFP作成・ベンダー選定", "start": "2025-08-01", "end": "2025-09-30"}
            ]
        },
        {
            "name": "Phase2: 設計",
            "type": "project",
            "items": [
                {"name": "設計", "type": "period", "start": "2025-10-01", "end": "2026-03-31"},
                {"name": "設計完了", "type": "milestone", "date": "2026-03-31"}
            ],
            "tasks": [
                {"name": "要件定義", "start": "2025-10-01", "end": "2025-12-31"},
                {"name": "基本設計", "start": "2026-01-01", "end": "2026-02-28"},
                {"name": "詳細設計", "start": "2026-02-01", "end": "2026-03-31"}
            ]
        },
        {
            "name": "Phase3: 開発",
            "type": "project",
            "items": [
                {"name": "開発", "type": "period", "start": "2026-04-01", "end": "2026-12-31"}
            ],
            "tasks": [
                {"name": "Sprint 1-5", "start": "2026-04-01", "end": "2026-06-30"},
                {"name": "Sprint 6-10", "start": "2026-07-01", "end": "2026-09-30"},
                {"name": "Sprint 11-15", "start": "2026-10-01", "end": "2026-12-31"}
            ]
        },
        {
            "name": "Phase4: テスト・移行",
            "type": "project",
            "items": [
                {"name": "テスト・移行", "type": "period", "start": "2027-01-01", "end": "2027-03-31"},
                {"name": "本番カットオーバー", "type": "milestone", "date": "2027-03-31"}
            ],
            "tasks": [
                {"name": "総合テスト", "start": "2027-01-01", "end": "2027-01-31"},
                {"name": "UAT", "start": "2027-02-01", "end": "2027-02-28"},
                {"name": "本番移行・安定化", "start": "2027-03-01", "end": "2027-03-31"}
            ]
        }
    ]


def make_18month_data():
    """18ヶ月サンプル（2025年4月〜2026年9月）"""
    return [
        {
            "name": "企画・構想",
            "type": "project",
            "items": [
                {"name": "PJ開始", "type": "milestone", "date": "2025-04-01"},
                {"name": "企画・構想", "type": "period", "start": "2025-04-01", "end": "2025-07-31"}
            ],
            "tasks": [
                {"name": "現状分析", "start": "2025-04-01", "end": "2025-05-15"},
                {"name": "構想策定", "start": "2025-05-01", "end": "2025-06-30"},
                {"name": "企画承認", "start": "2025-07-01", "end": "2025-07-31"}
            ]
        },
        {
            "name": "設計",
            "type": "project",
            "items": [
                {"name": "設計", "type": "period", "start": "2025-07-01", "end": "2025-12-31"},
                {"name": "設計完了", "type": "milestone", "date": "2025-12-31"}
            ],
            "tasks": [
                {"name": "要件定義", "start": "2025-07-01", "end": "2025-08-31"},
                {"name": "基本設計", "start": "2025-08-15", "end": "2025-10-31"},
                {"name": "詳細設計", "start": "2025-10-01", "end": "2025-12-31"}
            ]
        },
        {
            "name": "開発",
            "type": "project",
            "items": [
                {"name": "開発", "type": "period", "start": "2026-01-01", "end": "2026-06-30"}
            ],
            "tasks": [
                {"name": "バックエンド開発", "start": "2026-01-01", "end": "2026-04-30"},
                {"name": "フロントエンド開発", "start": "2026-02-01", "end": "2026-05-31"},
                {"name": "API連携", "start": "2026-04-01", "end": "2026-06-30"}
            ]
        },
        {
            "name": "テスト・リリース",
            "type": "project",
            "items": [
                {"name": "テスト", "type": "period", "start": "2026-06-01", "end": "2026-09-30"},
                {"name": "本番リリース", "type": "milestone", "date": "2026-09-30"}
            ],
            "tasks": [
                {"name": "結合テスト", "start": "2026-06-01", "end": "2026-07-15"},
                {"name": "総合テスト", "start": "2026-07-01", "end": "2026-08-15"},
                {"name": "受入テスト・移行", "start": "2026-08-01", "end": "2026-09-30"}
            ]
        }
    ]


def make_32month_data():
    """32ヶ月サンプル（2025年4月〜2027年11月）"""
    return [
        {
            "name": "Phase1: 企画",
            "type": "project",
            "items": [
                {"name": "PJ開始", "type": "milestone", "date": "2025-04-01"},
                {"name": "企画", "type": "period", "start": "2025-04-01", "end": "2025-09-30"}
            ],
            "tasks": [
                {"name": "現状分析・課題整理", "start": "2025-04-01", "end": "2025-06-30"},
                {"name": "RFP・ベンダー選定", "start": "2025-07-01", "end": "2025-09-30"}
            ]
        },
        {
            "name": "Phase2: 設計",
            "type": "project",
            "items": [
                {"name": "設計", "type": "period", "start": "2025-10-01", "end": "2026-06-30"},
                {"name": "設計完了", "type": "milestone", "date": "2026-06-30"}
            ],
            "tasks": [
                {"name": "要件定義", "start": "2025-10-01", "end": "2026-01-31"},
                {"name": "基本設計", "start": "2026-01-01", "end": "2026-04-30"},
                {"name": "詳細設計", "start": "2026-04-01", "end": "2026-06-30"}
            ]
        },
        {
            "name": "Phase3: 開発",
            "type": "project",
            "items": [
                {"name": "開発", "type": "period", "start": "2026-07-01", "end": "2027-06-30"}
            ],
            "tasks": [
                {"name": "Sprint群 前半", "start": "2026-07-01", "end": "2026-12-31"},
                {"name": "Sprint群 後半", "start": "2027-01-01", "end": "2027-06-30"}
            ]
        },
        {
            "name": "Phase4: テスト",
            "type": "project",
            "items": [
                {"name": "テスト", "type": "period", "start": "2027-06-01", "end": "2027-09-30"},
                {"name": "テスト完了", "type": "milestone", "date": "2027-09-30"}
            ],
            "tasks": [
                {"name": "統合テスト", "start": "2027-06-01", "end": "2027-07-31"},
                {"name": "UAT", "start": "2027-08-01", "end": "2027-09-30"}
            ]
        },
        {
            "name": "Phase5: 移行・安定化",
            "type": "project",
            "items": [
                {"name": "移行", "type": "period", "start": "2027-10-01", "end": "2027-11-30"},
                {"name": "カットオーバー", "type": "milestone", "date": "2027-11-30"}
            ],
            "tasks": [
                {"name": "データ移行", "start": "2027-10-01", "end": "2027-10-31"},
                {"name": "本番移行・安定化", "start": "2027-11-01", "end": "2027-11-30"}
            ]
        }
    ]


def make_48month_data():
    """48ヶ月サンプル（2025年4月〜2029年3月：4年度）"""
    return [
        {
            "name": "Year1: 企画・構想",
            "type": "project",
            "items": [
                {"name": "PJ開始", "type": "milestone", "date": "2025-04-01"},
                {"name": "企画・構想", "type": "period", "start": "2025-04-01", "end": "2026-03-31"}
            ],
            "tasks": [
                {"name": "現状調査", "start": "2025-04-01", "end": "2025-09-30"},
                {"name": "構想策定・承認", "start": "2025-10-01", "end": "2026-03-31"}
            ]
        },
        {
            "name": "Year2: 設計・開発(1)",
            "type": "project",
            "items": [
                {"name": "設計開発1", "type": "period", "start": "2026-04-01", "end": "2027-03-31"},
                {"name": "Phase1完了", "type": "milestone", "date": "2027-03-31"}
            ],
            "tasks": [
                {"name": "要件定義・基本設計", "start": "2026-04-01", "end": "2026-09-30"},
                {"name": "詳細設計・開発", "start": "2026-10-01", "end": "2027-03-31"}
            ]
        },
        {
            "name": "Year3: 開発(2)・テスト",
            "type": "project",
            "items": [
                {"name": "開発テスト", "type": "period", "start": "2027-04-01", "end": "2028-03-31"},
                {"name": "Phase2完了", "type": "milestone", "date": "2028-03-31"}
            ],
            "tasks": [
                {"name": "追加開発", "start": "2027-04-01", "end": "2027-09-30"},
                {"name": "統合テスト・UAT", "start": "2027-10-01", "end": "2028-03-31"}
            ]
        },
        {
            "name": "Year4: 移行・運用",
            "type": "project",
            "items": [
                {"name": "移行・運用", "type": "period", "start": "2028-04-01", "end": "2029-03-31"},
                {"name": "全面稼働", "type": "milestone", "date": "2029-03-31"}
            ],
            "tasks": [
                {"name": "段階移行", "start": "2028-04-01", "end": "2028-09-30"},
                {"name": "全面移行・安定化", "start": "2028-10-01", "end": "2029-03-31"}
            ]
        }
    ]


if __name__ == "__main__":
    samples = [
        ("sample_3month.pptx",  make_3month_data(),  99, "auto"),
        ("sample_6month.pptx",  make_6month_data(),  99, "auto"),
        ("sample_12month.pptx", make_12month_data(), 3,  "auto"),
        ("sample_18month.pptx", make_18month_data(), 3,  "auto"),
        ("sample_24month.pptx", make_24month_data(), 3,  "auto"),
        ("sample_32month.pptx", make_32month_data(), 3,  "auto"),
        ("sample_48month.pptx", make_48month_data(), 3,  "auto"),
        # 四半期モード
        ("sample_32month_q.pptx", make_32month_data(), 3, "quarterly"),
        ("sample_48month_q.pptx", make_48month_data(), 3, "quarterly"),
    ]
    
    for filename, data, threshold, mode in samples:
        gen = GanttChartGenerator(data, force_external_label_months=threshold, calendar_mode=mode)
        gen.generate(filename)
