# -*- coding: utf-8 -*-
from scrapy import Spider, Request
from collections import OrderedDict
from xlrd import open_workbook
import os, requests, re, time

def download(url, destfilename):
    if not os.path.exists(destfilename):

        try:
            r = requests.get(url, stream=True)
            # if r.status_code != 200:
            #     r = requests.get(temp_url, stream=True)
            with open(destfilename, 'wb') as f:
                for chunk in r.iter_content(chunk_size=1024):
                    if chunk:
                        f.write(chunk)
                        f.flush()
        except:
            # print(url)
            pass

def readExcel(path):
    wb = open_workbook(path)
    result = []
    for sheet in wb.sheets():
        number_of_rows = sheet.nrows
        number_of_columns = sheet.ncols
        herders = []
        for row in range(0, number_of_rows):
            values = OrderedDict()
            for col in range(number_of_columns):
                value = (sheet.cell(row,col).value)
                if row == 0:
                    herders.append(value)
                else:

                    values[herders[col]] = value
            if len(values.values()) > 0:
                result.append(values)
        break

    return result


class AngelSpider(Spider):
    name = "manomano_scraper"
    start_urls = 'https://www.manomano.fr/'
    count = 0
    use_selenium = False
    urls = readExcel("Input.xlsx")
    models = []

    headers = ['EAN', 'Nom', 'Liens', 'Prix', 'Catégorie', 'Vendeur','Description', 'Image1', 'Image2', 'Image3', 'Image4', 'Image5']

    def start_requests(self):
        yield Request(self.start_urls, callback=self.parse)

    def parse(self, response):
        for i, val in enumerate(self.urls):
            ern_code = val['URL']
            if ern_code != '':
                yield Request(response.urljoin(ern_code), callback=self.parse1)
                # yield Request('https://www.manomano.fr/manuel-et-livre-de-jardinage-3811', callback=self.parse1)
    def parse1(self, response):
        urls = response.xpath('//ul[@class="product-list__products"]//a[@class="product-card-content"]/@href').extract()
        for url in urls:
            time.sleep(2)
            yield Request(response.urljoin(url), callback=self.final_parse)
            # yield Request('https://www.manomano.fr/jardinage-pour-enfants-3867', callback=self.final_parse)
            # break
        next_tag = response.xpath('//*[@class="pagination__item pagination__item--next"]/a/@href').extract_first()
        if next_tag:
            yield Request(response.urljoin(next_tag), callback=self.parse1)

    def final_parse(self, response):
        try:
            item = OrderedDict()
            for key in self.headers:
                item[key] = None

            categories = response.xpath('//ul[@class="breadcrumbs product__breadcrumbs-top"]//span[@class="breadcrumbs__label"]/text()').extract()

            ean = response.xpath('//script[@data-flix-inpage="flix-inpage"]/@data-flix-ean').extract_first()
            if not ean:
                ean = response.xpath('//meta[@itemprop="gtin13"]/@content').extract_first()
            item['EAN'] = ean

            name = response.xpath('//div[@class="product-info"]/@data-product-title').extract_first()
            # name = ' '.join(name_list).replace('\r\n', '').replace('  ', ' ').strip()
            item['Nom'] = name

            item['Liens'] = response.url

            price = response.xpath('//p[@itemprop="price"]/@content').extract_first()
            # price= prices.replace(',', '.').replace(' ', '').replace('€', '')
            item['Prix'] = price

            item['Catégorie'] = categories[-2].strip()

            item['Vendeur'] = response.xpath('//div[@class="product-seller__info"]/p/a/text()').extract_first()

            descript1_list = response.xpath('//div[@class="description_content"]//text()').extract()
            # if not descript1_list:
            #     descript1_list = response.xpath('//div[@class="product_bloc_content bloc ombre font-2"]//text()').extract()
            descript1 = ''.join(descript1_list).strip()
            item['Description'] = descript1

            tr_tags = response.xpath('//li[@class="list-table__row"]')
            for tr in tr_tags:
                key = tr.xpath('./span[1]/text()').extract_first()
                val = ''.join(tr.xpath('./span[2]//text()').extract())
                item[key] = val
                if not key in self.headers:
                    self.headers.append(key)


            for i in range(5):
                item['Image' + str(i+1)] = ''
            #
            index = 0
            #
            image_urls = response.xpath('//li[contains(@class, "image-inspector__thumbnail-item")]/@data-image').extract()
            # if len(image_urls) < 5:
            #     image_urls1 = response.xpath('//div[contains(@class, "v6horizontal_new_product_page_sizes")]/img/@data-src').extract()
            #     if image_urls1:
            #         image_urls.extend(image_urls1)

            for image_url in image_urls:
                index += 1
                if index > 5:
                    break
                # baseName = image_url.strip().split('/')[-1].replace('.jpg', '')
                temp_name = re.sub('[^A-Za-z0-9 ]+', '', item['Nom'])
                image_name = temp_name.strip().replace(' ', '_').replace('__', '_') + '_' + 'manomano' + '_' + str(index) +".jpg"
                download(image_url.strip(), 'Images/'+image_name)
                item['Image' + str(index)] = image_name

            # specs_tag = response.xpath('//div[@class="product_bloc_content bloc ombre "]//tr')
            # for tag in specs_tag:
            #     key = tag.xpath('./th/text()').extract_first()
            #     val = tag.xpath('./td/text()').extract_first()
            #     if key and val:
            #         if not key.strip() in self.headers:
            #             self.headers.append(key.strip())
            #         item[key.strip()] = val.strip()
            #
            # codic = response.xpath('//*[@class="darty_product_main_content product_column product_left"]/@data-codic').extract_first()
            #
            # yield Request('https://www.darty.com/nav/extra/ajax/zoom_popin?codic={}&flash=disabled'.format(codic), callback=self.getImage, meta={'item':item, 'img_urls':image_urls})
            self.models.append(item)
            self.count += 1
            print(self.count)
            yield item

        except:
            pass

    def getImage(self, response):
        item = response.meta['item']
        index = 0
        img_urls = response.meta['img_urls']
        image_urls = response.xpath('//div[@id="darty_zoom_popin_container"]//img/@src').extract()

        for i, image_url in enumerate(image_urls):
            index += 1
            if index > 5:
                break
            temp_name = re.sub('[^A-Za-z0-9 ]+', '', item['Nom'])
            image_name = temp_name.strip().replace(' ', '_').replace('__', '_') + '_' + 'tropicmarket' + '_' + str(index) +".jpg"
            filename = "Images/" + image_name
            tem_url = img_urls[i]
            download(image_url.strip(), filename, tem_url)
            item['Image' + str(index)] = image_name

        self.models.append(item)
        yield item