import os, sys
# モジュールインポート
sys.path.append(os.path.join(os.path.dirname(__file__), 'site-packages'))
import fitz
import pandas as pd
import numpy as np

# 変数定義（後で外部定義）
ignoreColumnList = ['名称', '識別2', '備考'] # 無視する列
dayColumn = '経過日数' # 時間で展開する列
reflectColumn = '菌数' # 反映する列

# 入力
print('PDFファイルパスを入力')
inputFilePath = input() # 入力ファイル
# inputFilePath = 'sample.pdf'
print('Excelファイルパスを入力')
outputFilePath = input() # 出力ファイル
# outputFilePath = 'debug'

# PDFの読み込み
doc = fitz.open(inputFilePath)

# ページの指定
page = doc[0]

# 表データの検索
tables = page.find_tables()

if tables.tables:
    table_data = tables[0].extract()
    # 列名にするデータを取得
    columns = table_data[0]

    # 表のデータを取得
    data_rows = table_data[1:]

    # データフレームを作成
    df = pd.DataFrame(data_rows, columns=columns)
    
    # 整形
    df = df.replace("", np.nan)  # nanに置き換え
    df = df.dropna(how="all")  # すべての列がnanの行を削除

    # 無視対象列
    for c in ignoreColumnList:
        df = df.drop(c, axis=1)
    
    # 表内のヘッダー削除（ロジック要改善）
    df = df.drop(df.index[df[df.columns[0]] == df.columns[0]])

    # 日付展開
    df[dayColumn] = df[dayColumn].ffill().astype(int) # nan -> 0に置き換えてから変換
    maxDay = df[dayColumn].max()
    for d in range(maxDay):
        c = 'D+{}'.format(d+1)
        df[c] = df[reflectColumn]
        df.loc[df[dayColumn] != d+1, c] = None
    
    if outputFilePath == 'debug':
        print(df)
    else:
        # データフレームをExcelに保存
        df.to_excel(outputFilePath, index=False)