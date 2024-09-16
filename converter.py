# coding: utf-8

"""
Markdownで書かれたテスト仕様書をエクセルファイルに変換します。

Usage:
    converter.py [-f] <file> [-m]

Options:
    -f, --file             入力ファイルパス
    -m, --merge            エクセルセルをマージするか

Requirements:
    - pandas
    - openpyxl 3.0.0 or higher
    - PyYAML 5.0.0 or higher
    - docopt

Notes:
    - 変換するMarkdownは以下の形式で記述してください

    # テスト名
    ## 大項目
    ### 中項目
    #### [正常|異常] [OK|NG|--] テストケース名
    1. 確認手順
    2. 確認手順
    * [ ] 期待値
    - 備考

"""

import os
import sys

try:
    import yaml
    import pandas as pd
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.styles.borders import BORDER_THIN
    from docopt import docopt
except ModuleNotFoundError as e:
    sys.stderr.write("This program requires pandas/docopt/openpyxl>=3.0.0.")
    sys.exit(1)
assert yaml.__version__ >= "5.0.0", "This program requires PyYAML>=5.0.0.\n$ pip install pyyaml==5.3.1"
assert openpyxl.__version__ >= "3.0.0", "This program requires openpyxl>=3.0.0.\b$ pip install openpyxl==3.0.5"

from markdown import convert_md_to_df
from excel import convert_df_to_excel


def load_config() -> dict:
    try:
        with open("config.yaml", "r", encoding="utf-8_sig") as f:
            config = yaml.load(f, Loader=yaml.FullLoader)
    except FileNotFoundError:
        sys.stderr.write("[ERROR] 設定ファイル（config.yaml）が見つかりません。")
        sys.exit(1)

    return config


if __name__ == "__main__":
    args = docopt(__doc__)
    config = load_config()
    df = convert_md_to_df(args["<file>"], config_md=config["md"])
    convert_df_to_excel(df, config_excel=config["excel"],
                        output_path=os.path.splitext(os.path.basename(args["<file>"]))[0] + ".xlsx",
                        merge_cells=args["--merge"])
    print("Done")
