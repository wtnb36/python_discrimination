# new.pyで作成したExcelデータのカテゴリ分けし書き込む
import openpyxl
import pandas as pd

#Excel読み込み
df = pd.read_excel("test.xlsx", sheet_name = 0)

#df["都道府県"]を取得しリスト化
prefecture_list = df["都道府県"].to_list()

#重複項目の削除
prefecture_list_unique = list(set(prefecture_list))

#都道府県毎にエクセルファイルを新規作成、書き込み、保存
for i in range(len(prefecture_list_unique)):
  df_prefecture = df.loc[df["都道府県"] == prefecture_list_unique[i]]
  df_prefecture.to_excel(prefecture_list_unique[i] + ".xlsx", index = False)

