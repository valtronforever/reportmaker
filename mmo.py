import csv
import datetime
from decimal import Decimal

delimiter = '	'


def get_mmo_info(path):
    info = {}
    with open(path, encoding='cp1251') as csvfile:
        reader = csv.reader(csvfile, delimiter=delimiter)
        for index, row in enumerate(reader, start=1):
            if index == 2:
                info['number'] = row[0]
                info['date'] = datetime.datetime.strptime(row[1], "%d.%m.%Y").date()
                info['supplier'] = row[3]
                break
    return info


def get_mmo_items(path):
    with open(path, encoding='cp1251') as csvfile:
        reader = csv.reader(csvfile, delimiter=delimiter)
        for index, row in enumerate(reader, start=1):
            if index >= 4:
                yield {
                    'title': row[1],
                    'maker': row[3],
                    'morion_id': row[4],
                    'price': Decimal(row[19].replace(',', '.')),
                    'count': int(row[15]),
                    'units': row[14],
                }
