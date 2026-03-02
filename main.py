"""
ガントチャート生成 CLI エントリーポイント

Usage:
    python main.py --input data.xlsx --output gantt.pptx
    python main.py --input data.xlsx --output gantt.pptx --mode quarterly
    python main.py --input data.xlsx --output gantt.pptx --start 2025-04 --end 2026-03
"""
import argparse
from excel_reader import read_excel
from gantt_generator import GanttChartGenerator


def main():
    parser = argparse.ArgumentParser(description="Excel → ガントチャート PPTX 生成ツール")
    parser.add_argument("--input", "-i", required=True, help="入力Excelファイルのパス")
    parser.add_argument("--output", "-o", default="output.pptx", help="出力PPTXファイルのパス（デフォルト: output.pptx）")
    parser.add_argument("--sheet", "-s", default=None, help="Excelのシート名（デフォルト: 最初のシート）")
    parser.add_argument("--start", default=None, help="表示開始月 (YYYY-MM)")
    parser.add_argument("--end", default=None, help="表示終了月 (YYYY-MM)")
    parser.add_argument("--mode", "-m", default="auto",
                        choices=["auto", "bme", "monthly", "quarterly"],
                        help="カレンダー表示モード（デフォルト: auto）")
    parser.add_argument("--external-label", action="store_true", default=False,
                        help="矢羽テキストを図形の外に配置する")
    
    args = parser.parse_args()
    
    # 1. Excelからデータ読み込み
    print(f"[読み込み] {args.input}")
    data = read_excel(args.input, sheet_name=args.sheet)
    
    if not data:
        print("エラー: データが見つかりませんでした。Excelのフォーマットを確認してください。")
        return
    
    print(f"  → {len(data)} プロジェクト検出")
    for d in data:
        items_count = len(d.get("items", []))
        tasks_count = len(d.get("tasks", []))
        print(f"    - {d['name']}: {items_count} items, {tasks_count} tasks")
    
    # 2. ガントチャート生成
    gen = GanttChartGenerator(
        data,
        start_month=args.start,
        end_month=args.end,
        external_label=args.external_label,
        calendar_mode=args.mode,
    )
    gen.generate(args.output)


if __name__ == "__main__":
    main()
