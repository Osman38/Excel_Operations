import requests
import xlwings as xw
from bs4 import BeautifulSoup


def cargoUPS(url):
    page = BeautifulSoup(requests.get(url).content, 'html.parser')
    try:
        statusCargo = page.find('span', attrs={'id': 'ctl00_MainContent_Label3'}).text
        return statusCargo
    except AttributeError:
        statusCargo = page.find('span', attrs={'id': 'ctl00_MainContent_Label56'}).text
        return statusCargo


def read_excel(file):
    sheet = xw.Book(file).sheets['sayfa-1']
    cargoUrl = sheet.range('k:k')
    s = 1
    for cat in cargoUrl:
        if cat.value is not None and cat.value != 'KARGO URL':
            newValue = cargoUPS(cat.value)
            x = sheet[f'l{s}'].value = newValue
            s += 1
            print(cat.value)
        elif  cat.value is not None or cat.value == 'KARGO URL':
            s += 1

    return 'finish'


if __name__ == '__main__':
    read_excel('hb_cargo.xlsx')
