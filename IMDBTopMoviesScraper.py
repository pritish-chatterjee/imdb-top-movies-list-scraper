import requests
import openpyxl
from bs4 import BeautifulSoup as bs

url = "https://www.imdb.com/chart/top/?ref_=nv_mv_250"

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = "IMDB 250 Top Rated Movies"
sheet.append(['Rank','Name','Year of Release','IMDB Rating'])


try:
    source = requests.get(url)
    source.raise_for_status()

    soup = bs(source.text, 'html.parser')

    movies = soup.find('tbody', class_="lister-list").find_all('tr')
    for movie in movies:
        name = movie.find('td', class_="titleColumn").a.text
        rank = movie.find('td', class_="titleColumn").get_text(
            strip=True).split('.')[0]
        year = movie.find('td', class_="titleColumn").span.text.strip('()')
        rating = movie.find(
            'td', class_="ratingColumn imdbRating").strong.text
        sheet.append([rank,name,year,rating])
except Exception as e:
    print(e)

excel.save('IMDB Movie Ratings.xlsx')