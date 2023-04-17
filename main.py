from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
import imdb

im = imdb.IMDb()

films = ["0078748", "0093773", "1375666",
         "0090605", "0167261", "0369339",
         "0372784", "0468569", "0080455",
         "0167260", "0120737", "0208092",
         "0253556", "0238380", "0209144",
         "0278504", "0076740", "0090180",
         "0147800", "0112864", "0114369",
         "0113189", "0381061", "0075005",
         "1853728", "0110912", "0361748",
         "1392214", "1856101", "3397884",
         "0119081", "0096256"
         ]

movieTest = im.search_movie("They Live")
print(movieTest)
movie = im.get_movie('0096256').data
title = movie['original title']
year = movie['year']
for i in movie['director']:
    print(f'{title} - {i} - {year}')

wb = Workbook()
ws = wb.active


center_align = Alignment(horizontal='center', vertical='center')

ws['B1'] = "Title"
ws['B1'].font = Font(name='Verdana', size=18, bold=True, color='00FF6600')

for c in ws['A2:A100']:
    c[0].alignment = center_align

for c in ws['B2:B4']:
    c[0].font = Font(size=16, italic=True)

for i in range(1, 100):
    ws.cell(row=i+1, column=1, value="*")

column = 2
for i, v in enumerate(films):
    ws.cell(row=i+2, column=column, value=v)

ws.title = "Films"

wb.save(filename="films.xlsx")
