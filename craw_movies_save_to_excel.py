import requests, bs4, openpyxl

wb = openpyxl.Workbook()
ws = wb.active
ws.title = 'IMDb Top250 Movies'

ws['A1'] = 'Rank'
ws['B1'] = 'Movie Name'
ws['C1'] = 'Stars Rating (Number of ratings)'
ws['D1'] = 'Year'
ws['E1'] = 'Rated'
ws['F1'] = 'Movie Link'

headers={'user-agent':'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_13_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3578.98 Safari/537.36'}
for x in range(10):
    url = 'https://www.imdb.com/chart/top/?ref_=nv_mv_250'
    res = requests.get(url, headers=headers)
    bs = bs4.BeautifulSoup(res.text, 'html.parser')
    bs = bs.find('ul', class_="ipc-metadata-list ipc-metadata-list--dividers-between sc-71ed9118-0 kxsUNk compact-list-view ipc-metadata-list--base")

    for titles in bs.find_all('li'):
        #num = titles.find('h3',class_="ipc-title__text").text
        title = titles.find('h3', class_="ipc-title__text").text
        title_line = title.split('.')
        num = title_line[0]
        title = title_line[1]
        stars_rating = titles.find('span',class_="ipc-rating-star ipc-rating-star--base ipc-rating-star--imdb ratingGroup--imdb-rating").text
        meta_data = titles.find('div',class_="sc-43986a27-7 dBkaPT cli-title-metadata").text
        year = meta_data[:4]
        if 'h' in meta_data and 'm' in meta_data:
            index = meta_data.find('m') + 1
            rated = meta_data[index:]
        else:
            rated = 'Rated Not Found'
        url_movie = titles.find('a')['href']
  
        ws.append([num,title,stars_rating,year,rated,'www.imdb.com'+url_movie])
        print(num + '.' + title + '——' + stars_rating + '\n' + year + ' ' + rated + '\n www.imdb.com' + url_movie + '\n')
            
wb.save('imdb_top250.xlsx') 