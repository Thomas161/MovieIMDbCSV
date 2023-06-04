from openpyxl import Workbook, load_workbook
# from collections import Counter
import pandas as pd
# import seaborn as sns

title = pd.read_excel('films.xlsx', sheet_name='Films', usecols='B')
titleList = title['Title'].to_list()
year = pd.read_excel('films.xlsx', sheet_name='Films', usecols='D')
yearList = year['Year'].to_list()
df = pd.DataFrame({'year': yearList, 'movie': titleList})
df2 = df.pivot_table(index=['year', 'movie'], values=['movie'], aggfunc='size')
print(df2)
