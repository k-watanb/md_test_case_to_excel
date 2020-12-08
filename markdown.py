# coding: utf-8

import re
import sys

import pandas as pd


def load_md(input_path: str):
    try:
        input_file = open(input_path, "r", encoding="utf-8")
        return input_file
    except FileNotFoundError:
        sys.stderr.write("[ERROR] Markdownファイル（.md）が見つかりません。")
        sys.exit(1)


def append_df(df: pd.DataFrame, current_item_dict: dict) -> pd.DataFrame:
    # ここまでの項目をデータフレームに追加
    _df = df.append(pd.Series([v for _, v in current_item_dict.items()], index=df.columns), ignore_index=True)

    # 終わったら初期化
    current_item_dict["steps"] = ""
    current_item_dict["expected"] = ""
    current_item_dict["notes"] = ""

    return _df


def can_append_df(current_item_dict: dict) -> bool:
    return (current_item_dict["steps"] and current_item_dict["expected"])


def convert_md_to_df(input_path: str, config_md: dict) -> pd.DataFrame:
    """
    Markdown形式で記述したテスト仕様書をExcelに変換するツール

    Args:
        input_path:        入力ファイルパス
        config_md:         マークダウン部分に関する設定

    Returns:
        df:                データフレーム型テスト仕様書
    """

    # Excelデータ変換用の空データフレーム
    df = pd.DataFrame(columns=[k for k, _ in config_md["name"].items()])
    current_item_dict = {k: "" for k, _ in config_md["name"].items()}
    input_file = load_md(input_path)

    test_case_num = 1
    for i, line in enumerate(input_file):

        if re.match(config_md["mark"]["steps"], line):
            current_item_dict["steps"] += line
        elif line.startswith(config_md["mark"]["expected"]):
            current_item_dict["expected"] += "・" + line.replace(config_md["mark"]["expected"], "").lstrip()
        elif line.startswith(config_md["mark"]["notes"]):
            current_item_dict["notes"] += "・" + line.replace(config_md["mark"]["notes"], "").lstrip()

        elif line.startswith(config_md["mark"]["item"]):
            if can_append_df(current_item_dict):
                df = append_df(df, current_item_dict)

            current_item_dict["item"] = line.replace(config_md["mark"]["item"], "").replace("\n", "").lstrip()
        elif line.startswith(config_md["mark"]["sub-item"]):
            if can_append_df(current_item_dict):
                df = append_df(df, current_item_dict)
            current_item_dict["sub-item"] = line.replace(config_md["mark"]["sub-item"], "").replace("\n", "").lstrip()
        elif line.startswith(config_md["mark"]["sub-sub-item"]):

            if can_append_df(current_item_dict):
                df = append_df(df, current_item_dict)

            if line[5:7] not in ["正常", "異常"] or line[8:10] not in ["OK", "NG", "--"]:
                sys.stderr.write(f"Markdownファイルの{i}行目に構文エラーがあります。\n"
                                 "小項目は以下のフォーマットで入力してください。\n\n"
                                 "#### [正常|異常|準正常] [OK|NG|--] 小項目\n")
                sys.exit(1)
            current_item_dict["pos-neg"] = line[5:7]
            current_item_dict["result"] = line[8:10]
            current_item_dict["sub-sub-item"] = line[11:-1].lstrip()
            current_item_dict["number"] = test_case_num
            test_case_num += 1

    # ファイル終了時点の最後の項目を追加
    if can_append_df(current_item_dict):
        df = append_df(df, current_item_dict)

    return df
