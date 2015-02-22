from urllib.request import urlopen

from bs4 import BeautifulSoup
import xlwt

NONE = 'None'
ROOT_WEBSITE = 'http://sklep.gandolf.pl/category.php?id_category=1820'
ROOT_CATEGORY_NAME = 'Gandolf'
XLS_FILENAME = 'sklep_gandolf.xls'
XLS_DESC = 'Opis'
XLS_CODE = 'Kod'
XLS_NAME = 'Nazwa'


class Product:
    name = NONE
    desc = NONE
    code = NONE

    def print(self):
        if self.name:
            print('    Found product ---------------------------- ')
            print('    Product Name: ', self.name)
            print('    Product Desc: ', self.desc)
            print('    Product Code: ', self.code)

    def is_defined(self):
        if self.name and self.desc and self.code:
            return True
        else:
            return False


class Category:

    def __init__(self, name):
        super().__init__()
        self.name = name
        self.products = []

    products = []
    name = NONE

    def add(self, product):
        self.products.append(product)


def parse_category(soup, depth, category_name):
    categories_list = soup('ul', {'class': 'inline_list'})

    i = 0

    if not categories_list:
        parse_products(soup, category_name)
    else:
        for category in categories_list[0].find_all('a'):
            title = category.get('title')
            if title is not None:
                url_to_open = category.get('href')
                if depth is 0:
                    print('Found category: ', title)
                    category_soup = BeautifulSoup(urlopen(url_to_open).read())
                    parse_category(category_soup, depth + 1, title)
                elif depth is 1:
                    print('  Found subcategory: ', title)
                    category_soup = BeautifulSoup(urlopen(url_to_open).read())
                    parse_category(category_soup, 1, title)
            if i is 2:
                break
            i += 1


def parse_code(soup, product):
    code_soup = soup('p', {'id': 'product_reference'})
    if code_soup:
        code = code_soup[0].span.text
        product.code = code


def replace_with_newline(soup, element):
    for e in soup.findAll(element):
        e.replace_with('\n')


def parse_desc(soup, product):
    replace_with_newline(soup, 'p')

    desc_soup = soup('div', {'id': 'idTab1'})
    if desc_soup:
        desc = desc_soup[0].text
        product.desc = desc


def parse_name(soup, product):
    name_soup = soup('div', {'id': 'primary_block'})
    if name_soup:
        name = name_soup[0].h1.text
        product.name = name


def parse_product(soup):
    replace_with_newline(soup, 'br')

    product = Product()
    parse_name(soup, product)
    parse_code(soup, product)
    parse_desc(soup, product)
    if product.is_defined():
        product.print()
        # products.append(product)
    return product


def parse_products(soup, name):
    i = 0

    category = Category(name)

    for link in soup('p', {'class': 'product_desc'}):
        title = link.a.get('title')

        if title is not None:
            url_to_open = link.a.get('href') + '&n=100000'
            product_soup = BeautifulSoup(urlopen(url_to_open).read())
            product = parse_product(product_soup)
            category.add(product)

        if i is 4:
            break
        i += 1

    categories.append(category)


def write_product_to_xls_sheet(index, product, sheet):
    sheet.write(index + 1, 0, product.name)
    sheet.write(index + 1, 1, product.code)
    sheet.write(index + 1, 2, product.desc)


def write_category_to_xls_book(book, category):
    sheet = book.add_sheet(category.name[:31])
    sheet.write(0, 0, XLS_NAME)
    sheet.write(0, 1, XLS_CODE)
    sheet.write(0, 2, XLS_DESC)
    return sheet


def write_to_xls():
    book = xlwt.Workbook()

    for category in categories:
        sheet = write_category_to_xls_book(book, category)

        for index, product in enumerate(category.products):
            write_product_to_xls_sheet(index, product, sheet)

    book.save(XLS_FILENAME)

# --------------------------------------------------------------------------------------

categories = []

soup = BeautifulSoup(urlopen(ROOT_WEBSITE).read())
parse_category(soup, 0, ROOT_CATEGORY_NAME)
write_to_xls()
