import requests, bs4, openpyxl

wb = openpyxl.Workbook()
ws = wb.active
ws.title = '豆瓣Top250'

ws['A1'] = '排名'
ws['B1'] = '片名'
ws['C1'] = '评分'
ws['D1'] = '推荐语'
ws['E1'] = '影片链接'

headers={'user-agent':'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_13_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3578.98 Safari/537.36'}
for x in range(10):
    url = 'https://movie.douban.com/top250?start=' + str(x*25) + '&filter='
    res = requests.get(url, headers=headers)
    bs = bs4.BeautifulSoup(res.text, 'html.parser')
    bs = bs.find('ol', class_="grid_view")
    for titles in bs.find_all('li'):
        num = titles.find('em',class_="").text
        title = titles.find('span', class_="title").text
        comment = titles.find('span',class_="rating_num").text
        url_movie = titles.find('a')['href']

        if titles.find('span',class_="inq") != None:
            tes = titles.find('span',class_="inq").text
            ws.append([num,title,comment,tes,url_movie])
            print(num + '.' + title + '——' + comment + '\n' + '推荐语：' + tes +'\n' + url_movie)
        else:
            ws.append([num,title,comment,'',url_movie])
            print(num + '.' + title + '——' + comment + '\n' +'\n' + url_movie)
            
wb.save('douban.xlsx') 