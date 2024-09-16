"""
Markdownで書かれたテスト仕様書をエクセルファイルに変換します。

Usage:
    python converter.py -h
    python converter.py [-f] <file> [-m]
"""

import argparse
from pathlib import Path

from src.config_loader import load_config
from src.excel import ExcelWriter
from src.markdown import MarkdownTestParser, read_markdown_file

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Markdownで書かれたテスト仕様書をエクセルファイルに変換します。")
    parser.add_argument("-f", "--file", type=str, required=True, help="入力ファイルパス")
    parser.add_argument("-m", "--merge", action="store_true", help="セルをマージするか")
    args = parser.parse_args()

    config = load_config(Path(__file__).parent.joinpath("config.yaml"))

    markdown_content_example = read_markdown_file(Path(args.file))
    parser = MarkdownTestParser(markdown_content_example, config)
    df = parser.parse()
    print(f"-------\n{df}\n-------")

    writer = ExcelWriter(df, config)
    output_path = writer(Path(args.file).parent.joinpath(f"{Path(args.file).stem}.xlsx"), merge_cells=args.merge)
    print(f"\nDone! The file is saved at `{output_path}`.")
