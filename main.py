from openpyxl import Workbook, load_workbook
# from collections import Counter
# import pygwalker as pyg
import pandas as pd
# import matplotlib.pyplot as plt

title = pd.read_excel('films.xlsx', sheet_name='Films', usecols='B')
titleList = title['Title'].to_list()
year = pd.read_excel('films.xlsx', sheet_name='Films', usecols='D')
yearList = year['Year'].to_list()

# def conditionalcol(x):
#     if x < 1950:
#         return 'background-color: lightgreen'
#     elif x < 1960:
#         return 'background-color: yellow'
#     elif x < 1970:
#         return 'background-color: red'
#     elif x < 1980:
#         return 'background-color: MediumSeaGreen'
#     elif x < 1990:
#         return 'background-color: HotPink'
#     elif x < 2000:
#         return 'background-color: Chocolate'
#     elif x < 2010:
#         return 'background-color: Olive'
#     elif x < 2020:
#         return 'background-color: FireBrick'
#     else:
#         return 'background-color: orange'


# df = pd.DataFrame({'year': yearList, 'movie': titleList}).style.applymap(
#     subset=['year'], func=conditionalcol)

# df = pd.DataFrame({'year': yearList, 'movie': titleList})
# df.style.background_gradient(subset=['movie'], cmap='Dark2')
# df.applymap(subset=['movie'])

# df['year'].plt(kind='bar')
# plt.title = "years"
# plt.xlabel = "year"
# plt.ylabel = "movie"
# df.plot.bar(x='movie', y='year')
# print(df)
# .style
# .style.format(subset=['movie']).bar(align='left', color='magenta')
# technologies = ['Spark', 'Pandas', 'Java', 'Python', 'PHP']
# fee = [25000, 20000, 15000, 15000, 18000]
# duration = ['5o Days', '35 Days', np.nan, '30 Days', '30 Days']
# discount = [2000, 1000, 800, 500, 800]
# columns = ['Courses', 'Fee', 'Duration', 'Discount']

# Create DataFrame from multiple lists
# df = pd.DataFrame(
#     list(zip(technologies, fee, duration, discount)), columns=columns)
# print(df)
# # df.style.background_gradient()
# df.to_excel('films.xlsx', sheet_name='pivot')
# df2 = df.pivot_table(index=['year', 'movie'], values=[
#                      'movie'])
# print(df2.plot())
# df['movie'] = df['movie'].astype('str')
# df2.plot(kind='bar', figsize=(6, 4))
# plt.title = "Films"
# plt.xlabel = "Movies"
# plt.ylabel = "Years"

# plt.plot()
# df.style.background_gradient(cmap='green')
# print(df.head())
# df2.style.format({"Year": "Year"}).highlight_min(color='#cd4f39')
# print(df.style.bar())
# s = sns.
# df2.style.background_gradient(cmap='green')
# df2.to_excel('pivot.xlsx')
# with pd.ExcelWriter('films.xlsx') as writer:
#     df.to_excel(writer, sheet_name='Sheet')
# print(pivot)
