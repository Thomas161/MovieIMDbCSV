from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from imdb import Cinemagoer
from collections import Counter
import pandas as pd

im = Cinemagoer()
wb = Workbook()
ws = wb.active

ws = wb.create_sheet("Sheet_A")
ws.title = "Films"

center_align = Alignment(horizontal='center', vertical='center')
ws['B1'] = "Title"
ws['B1'].font = Font(name='Verdana', size=18, bold=True, color='00FFFFFF')
ws['B1'].fill = PatternFill("solid", start_color='00339966')
ws['C1'] = "Director"
ws['C1'].font = Font(name='Verdana', size=18, bold=True, color='00FFCC00')
ws['C1'].fill = PatternFill("solid", start_color='00FF8080')
ws['D1'] = "Year"
ws['D1'].font = Font(name='Verdana', size=18, bold=True, color='00FF6600')
ws['D1'].fill = PatternFill("solid", start_color='00003366')
ws['E1'] = "Box Office Global"
ws['E1'].font = Font(name='Verdana', size=18, bold=True, color='00FFFFFF')
ws['E1'].fill = PatternFill("solid", start_color='00800000')
ws['F1'] = "Genre"
ws['F1'].font = Font(name='Verdana', size=18, bold=True, color='000066CC')
ws['F1'].fill = PatternFill("solid", start_color='00FFFF00')
# ws['A103'] = "Year with most Films"
# ws['A103'].font = Font(name='Verdana', size=18, bold=True, color='00FFFF99')
# ws['A103'].fill = PatternFill("solid", start_color='00003366')
# ws['A104'] = "Most Films by Director"
# ws['A104'].font = Font(name='Verdana', size=18, bold=True, color='00FFFF99')
# ws['A104'].fill = PatternFill("solid", start_color='00003366')

for c in ws['A2:A100']:
    c[0].alignment = center_align

for c in ws['B2:D100']:
    c[0].font = Font(size=16, italic=True)

for i in range(1, 100):
    ws.cell(row=i+1, column=1, value="*")

films = ["0078748", "0093773",
         "1375666"]
#          "0090605", "0167261", "0369339",
#          "0372784", "0468569", "0080455",
#          "0167260", "0120737", "0208092",
#          "0253556", "0238380", "0209144"
#          "0278504", "0076740", "0090180",
#          "0147800", "0112864", "0114369",
#          "0113189", "0381061", "0075005",
#          "1853728", "0110912", "0361748",
#          "1392214", "1856101", "3397884",
#          "0119081", "0096256", "0107076",
#          "0364725", "0172495", "0265086",
#          "0190590", "0116282", "0108358",
#          "0058461", "0059578", "1205489",
#          "0103644", "0396269", "0374900",
#          "0083658", "0103064", "0191397",
#          "0118887", "0107290", "0418279",
#          "0066999", "2265171", "0120611",
#          "1745960", "0187078", "0246578",
#          "0110413", "0116483", "0377092",
#          "0166924", "0086190", "1049413"
#  "0129167", "0120363", "0120755",
#  "4912910", "0076729", "0067116",
#  "0348333", "0115759", "0109040",
#  "0113277", "0947810", "0096874",
#  "0320661", "0097733", "0477348",
#  "0082869", "0108399", "0104348",
#  "0120201", "0075784", "0120863",
#  "0118880", "0211915", "0100403",
#  "0084434", "0117998", "0079817",
#  "0469494", "0146838", "0044079",
#  "0102138", "0104684", "0100802",
#  "0120586", "0085636", "10045260",


column = 2
columnThree = 3
columnFour = 4
columnFive = 5
columnSix = 6

# Testing individual films
# ex = im.get_keyword('Exorcist')
# extwo = im.get_movie('10045260')
# extwo = im.get('10045260')
# print(extwo)

movieListYear = []
movieTitles = []

# movie = im.search_movie('apocalypse now')
# moviewww = im.get_movie('0078788')
# print(moviewww['year'])


def render_films():
    for i, v in enumerate(films):
        movie = im.get_movie(v)
        title = movie['original title']
        year = movie['year']
        genre = movie['genres']
        for e in genre:
            for d in movie['directors']:
                ws.cell(row=i+2, column=column, value=title)
                ws.cell(row=i+2, column=columnThree, value=str(d['name']))
                ws.cell(row=i+2, column=columnFour, value=year)
                ws.cell(row=i+2, column=columnSix, value=e)

        movieTitles.append(title)
        movieListYear.append(year)


df = pd.DataFrame({'year': movieListYear, 'movie': movieTitles})
# print(df.head())
# df['movie'] = df['movie'].astype('str')
# df['movie'].str
# print(df.empty)s
df2 = df.pivot_table(index=['year', 'movie'], values=['movie'], aggfunc='size')
print(df2)
with pd.ExcelWriter('films.xlsx') as writer:
    df2.to_excel(writer, sheet_name='pivot')

render_films()
wb.save(filename="films.xlsx")

# print(movieTitles)
# print(movieListYear)

# df = pd.DataFrame({'year': movieListYear, 'movie': movieTitles})
# # print(df.head())
# # df['movie'] = df['movie'].astype('str')
# # df['movie'].str
# # print(df.empty)s
# df2 = df.pivot_table(index=['year', 'movie'], values=['movie'], aggfunc='size')
# print(df2)
# with pd.ExcelWriter('films.xlsx') as writer:
#     df2.to_excel(writer, sheet_name='pivot')
# pd.ExcelWriter(render_films())
# movieListYear.append(year)
# movieListYear.sort()
# # m = max(movieListYear, key=movieListYear.count)
# print('Most movies by year', movieListYear)
# for v in movieListYear:
#     cn = Counter(v)
#     m = max(cn)
#     print(m)
# use max(list,key=list.count) => will get max year/director
# movieListYear.append(year)
# # movieListYear.sort()
# m = max(movieListYear, key=movieListYear.count)
# print(m)
# cn = Counter(movieListYear)

# cn = max(movieListYear, key=movieListYear.count())
# print(f'list {li}')
# print()


# df = pd.DataFrame({'movie': [movie], 'year': [year], 'director': [d['name']]})
# print(df)
# for i in range(year):
# print(dir(i))
# cn = Counter(i)
#     print(cn)
#     total = max(cn)
# ws.cell(row=103, column=column, value=total)

# firstMockData = {'year': [year]}
# firstMockDF = pd.DataFrame(firstMockData)
# # print(firstMockData)
# df_pivot = pd.pivot(df, index='D1', columns='D1', values='D2:D10')
# print(df_pivot)
# firstMockDF.to_excel('test_wb.xlsx', 'sheetA', index=False)


# wb.
# wb = load_workbook("films.xlsx")
# s = pd.read_excel("films.xlsx")
# p1 = pd.pivot_table(s, index='year')

# print(p1)
# movieTest = im.search_movie("Halloween 3: Season of the Witch")
# print(movieTest)
# movie = im.get_movie('0085636').data
# title = movie['original title']
# year = movie['year']
# for i in movie['director']:
#     cn = Counter(i)
#     # total = cn[i]
#     print(f'{title} - {i} - {year}')
# print(f'{total}')

# director = movie['director']

# for i in director:
#     # print(i)
# build an array for year and director
# sort the array to make it easier
# use max(list,key=list.count) => will get max year/director
# movieListYear = []
# movieListDirectors = []

# code = "2690647"

# # searching the Id and getting info set
# person = im.get_person(code, info=['filmography'])

# # making infoset to use as keys
# person.infoset2keys

# # printing filmo graphy
# print(person['filmography']['actor'][0])

#     # for i in movie.infoset2keys['main'][24].name:
#     # for i in movie['director']:
#     #     director = im.search_person(i["name"])[0]
# print(f'{director[0]} - {title} - {year}')
# for i in range(len(director)):
#     d = [director[i] for i in (0)]


# def __repr__(self) -> str:
#     return self

# __repr__(director)
# directorTwo = im.search_person(i["name"][0])
# print(type(director))
# for d in director:
#     print(d)
# for d in director:
#     print(d)
# director = movie['director']
# result = im.search_person(director["name"])[0]
# print(result)
# for i in director:
#     print(str(i))
# title = movie['original title']
# movieListDirectors.append(i)
# movieListDirectors.sort()
# print(director)
# year = movie['year']
# print(f'{title} - {year} - {director}')

# def __init__(director):
#     print(director)
# print(f'{type(title)} -{type(year)} -{str(director)} ')

# movieListDirectors = []
# a =
# ws.cell(column=column, row=i+2).value = title + "" + director
# # ws.cell(column=columnThree, row=i+2).value = im.search_person(i["name"])[0]
# ws.cell(column=columnFour, row=i+2).value = year
# for i in movie['director']:
#     director = im.search_person(i["name"])[0]
# im.update(director)
# print(director)
# print(f'{title} - {director} - {year}')
# movieListDirectors = []
# movieListDirectors.append(str(director))
# movieListDirectors.sort()
# print(movieListDirectors)
# row = 0
# ws.cell(row=row+2, column=columnThree, value=str(director))

# for k in im.search_person(str(i))[:1]:
#     director = im.get_person(k.personID)
#     print(director)
# s = set()
# movieListDirectors.append(str(i))
# movieListDirectors.sort()
# print(movieListDirectors)
# for j in movieListDirectors:
#     print(j)
# direct = s.add(i)
# print(i)

# ws.cell(row=i+2, column=column, value=title)
# ws.cell(row=row+2, column=columnThree, value=str(director))
# ws.cell(row=i+2, column=columnFour, value=year)
# title = movie['original title']
# movieListDirectors.append(i)
# movieListDirectors.sort()
# print(i, v)
# year = movie['year']
# movieListDirectors = []
# ws.cell(row=i+2, column=column, value=title)
# ws.cell(row=i+2, column=columnFour, value=year)
# for i in movie['director']:
# for x in i:
#     print(x)
# ws.cell(row=i+2, column=column, value=title)
# ws.cell(row=i+2, column=columnThree, value=i)
# ws.cell(row=i+2, column=columnFour, value=year)
# movieListDirectors.append(i)
# movieListDirectors.sort()
# cn = Counter(movieListDirectors)
# cn = max(movieListDirectors, key=movieListDirectors.count)
# print(f'count {movieListDirectors}')
# print(f'{title}-{i}')
# for v in movieListDirectors:
#     cn = max(movieListDirectors, key=v.count)
# print(f'list directors {movieListDirectors}')
# print(f'Director of most films {cn}')

#
# list of movies by year, and the the most movies by year
#     year = movie['year']
#     movieList.append(year)
#     movieList.sort()
#     cn = max(movieList, key=movieList.count)
#     print(f'list {li}')
#     print(f'Highest Movie by year {cn}')
#

# for i in movie['']
#     ws.cell(row=i+2, column=column, value=title)
# print(title)
# movie = im.get_movie('0120586').data
# title = movie['original title']
# year = movie['year']
# for i in movie['director']:
#     print(f'{title} - {i} - {year}')
# ws.cell(row=i+2, column=column, value=title)
# ws.cell(row=i+2, column=columnThree, value=i)
# ws.cell(row=i+2, column=columnFour, value=year)
