from bs4 import BeautifulSoup
import requests
import openpyxl

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = "Movie List"
sheet.append(["Movie Rank", "Movie Name", "Year of Relase", "IBDB Rating"])
try:
    source = requests.get("https://www.imdb.com/chart/top/")
    source.raise_for_status()

    soup = BeautifulSoup(source.text, 'html.parser')

    movie = soup.find('tbody', class_='lister-list').find_all('tr')

    for movies in movie:

        name = movies.find('td', class_='titleColumn').a.text

        rank = movies.find('td', class_='titleColumn').get_text(
            strip=True).split('.')[0]

        year = movies.find('td', class_='titleColumn').span.text.strip('()')

        rating = movies.find(
            'td', class_='ratingColumn imdbRating').strong.text

        # print(rank, name, year, rating)

        sheet.append([rank, name, year, rating])

except Exception as e:
    print(e)
finally:
    print("This program is excuted")

excel.save('Movie List.xlsx')
