# -*- encoding: utf-8 -*-
from bs4 import BeautifulSoup as bs
import pandas as pd
from io import BytesIO
from urllib import request
from PIL import Image


class Scraper:
    def __init__(self):
        self.category_name = ''
        self.subcategory_name = ''
        self.db = pd.DataFrame()

    def collect_category_links(self):
        page = request.urlopen('http://a-plus.ua/index.php?route=common/home')
        soup = bs(page, 'html.parser')
        links = soup.find_all('a', class_='dropdown-toggle')
        category_links = [link['href'] for link in links]
        return category_links

    def process_good(self, url):
        page = request.urlopen(url)
        soup = bs(page, 'html.parser')
        temp = pd.DataFrame()
        container = soup.find('div', class_='product-info')
        b = soup.find('div', id='breadcrumb')
        temp['Category'] = [b.find_all('span')[1 - len(b.find_all('span'))].find(text=True)]
        temp['Subcategory'] = [b.find_all('span')[-2].find(text=True)]
        temp['Image'] = [str(container.find('img')['src'])]
        temp['Name'] = [container.find('h1').find(text=True)]
        strings = container.find('div', class_='price-gruop').find_all(text=True)
        temp['Price'] = [strings[-3].replace('\t', '').replace('\n', '') + ' ' + strings[-2]]
        self.db = self.db.append(temp)

    def process_subcategory(self, url):
        try:
            page = request.urlopen(url.replace('у', 'y').replace('с', 'c'))
        except Exception:
            return
        soup = bs(page, 'html.parser')
        catalog_box = soup.find('div', class_='products-block')
        try:
            goods_links = [img['href'] for img in catalog_box.find_all('a', class_='img')]
        except Exception:
            goods_links = []
        for goods_link in goods_links:
            self.process_good(goods_link)

    def process_category(self, url):
        page = request.urlopen(url.replace('с', 'c'))
        soup = bs(page, 'html.parser')
        subcategory_box = soup.find('div', class_='panel-body category-list clearfix box-content')
        subcategory_links = []
        for link in subcategory_box.find_all('a'):
            subcategory_links.append(link['href'])
        for subcategory_link, _ in zip(subcategory_links, range(len(subcategory_links))):
            print('Processing subcategory number', _ + 1, 'out of', len(subcategory_links), ':', subcategory_link)
            self.process_subcategory(subcategory_link)

    def scrape(self):
        category_links = self.collect_category_links()
        _ = 0
        for category_link in category_links:
            print('Processing category number', _ + 1, 'out of', len(category_links), ':', category_link)
            self.process_category(category_link)
            _ += 1
        self.db.to_csv('Database.csv')

'''
s = Scraper()
s.scrape()
'''
df = pd.read_csv('Database (1).csv', encoding='utf-8')
writer = pd.ExcelWriter('output.xlsx')
df.to_excel(writer, 'Sheet1')
wb = writer.book
ws = wb.get_worksheet_by_name('Sheet1')
for _, image_url in zip(range(df['Image'].size), df['Image']):
    print(_ + 1)
    image_data = BytesIO(request.urlopen(image_url).read())
    img = Image.open(image_data)
    scale = 183 / max(img.size)
    offset_x = (185 - img.size[0]*scale) / 2
    offset_y = (185 - img.size[1]*scale) / 2
    ws.set_row(_, 138)
    ws.insert_image(_ + 1, 6, image_url, {'image_data': image_data, 'x_scale': scale, 'y_scale': scale,
                                          'y_offset': offset_y, 'x_offset': offset_x})
ws.set_row(0, 20)
ws.set_column(0, 8, 25.5)
writer.save()
wb.close()

