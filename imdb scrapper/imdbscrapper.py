from bs4 import BeautifulSoup
import requests, openpyxl


excel = openpyxl.Workbook()
print(excel.sheetnames)
sheet = excel.active
sheet.title = ('Top rated movies of india')
print(excel.sheetnames)
sheet.append(['Movie Rank', 'Movie name', 'Year of release', 'Imdb Ratings'])

website = 'https://www.imdb.com/india/top-rated-indian-movies/'

#code for scrapping
try:
    source = requests.get(website)
    source.raise_for_status()
    soup=BeautifulSoup(source.text,'html.parser')   #to return html content
    movies=soup.find('tbody', class_="lister-list").find_all('tr')

    for movie in movies:
        name = movie.find('td',class_="titleColumn").a.text
        rank = movie.find('td',class_="titleColumn").get_text(strip=True).split('.')[0]
        year = movie.find('td',class_="titleColumn").span.text.strip('()')
        imdbR = movie.find('td',class_="ratingColumn imdbRating").strong.text 
        print(rank,name,year,imdbR)
        sheet.append([rank, name, year ,imdbR])


except Exception as e:
    print(e)    

excel.save('IMDB Movie Ratings.xlsx')