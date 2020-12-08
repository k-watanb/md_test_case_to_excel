# md_test_case_to_excel

Markdownで書かれたテスト仕様書をExcel形式に変換します。Markdown+GitHubでテスト仕様書を差分管理したいヒトはご活用ください。

![image-20201130175253467](attachments/image-20201130175253467.png)

## Get Started

### Write a test specification by Markdown

 ```markdown
# テストケース名

## 大項目
### 中項目
#### [正常|異常] [OK|NG|--] 小項目
1. 確認手順
2. 確認手順
* [ ] 想定動作
* [ ] 想定動作
- 備考
 ```

### Convert markdown to xlsx

```bash
$ python3 converter.py -h
$ python3 converter.py -f sample.md -m
```

実行時に指定できるオプションとして以下があります。

|オプション名|説明|
|:---|:---|
|-h, --help| 引数のヘルプ表示|
|-f, --file| 入力ファイルパス|
|-m, --merge| エクセルセルをマージするか|

## Environment

- Python 3.6 or higher
- pandas
- openpyxl 3.0.0 or higher
- PyYAML 5.0.0 or higher
- docopt

## Directory Structure

主なファイルは以下の通り。

```
.
|-- config.yaml    # ユーザ設定ファイル 
|-- converter.py   # MAIN
|-- markdown.py    # markdown関係の処理
|-- excel.py       # excel関係の処理
|-- sample.md      # markdown型テスト仕様書のサンプル（中身は適当）
```

## Release Notes

### v0.1.0 (2020/12/08)

- First Commit!

## Acknowledgements

The format of markdown borrows heavily from [ryuta46/eval-spec-maker](https://github.com/ryuta46/eval-spec-maker). The python code borrows from [torisawa/convert.py](https://gist.github.com/toriwasa/37c690862ddf67d43cfd3e1af4e40649)

## Limitations

Headers in xlsx are in Japanese only.
