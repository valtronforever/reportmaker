import pathlib
import xlrd
import datetime
from decimal import Decimal


def parse_int(x):
    temp = []
    for c in str(x).strip():
        if c.isdigit():
            temp.append(c)
        else:
            break
    return int("".join(temp)) if "".join(temp) else None


def list_files(path):
    files = []
    for currentFile in pathlib.Path(path).iterdir():
        if not currentFile.is_dir():
            files.append(str(currentFile))
    return files


def list_files_recursively(path):
    files = []
    for currentFile in pathlib.Path(path).rglob("*"):
        if not currentFile.is_dir():
            files.append(str(currentFile))
    return files


def list_xls(path):
    return [f for f in list_files(path) if f.lower().endswith('.xls')]


def list_xls_recursively(path):
    return [f for f in list_files_recursively(path) if f.lower().endswith('.xls')]


def list_mmo(path):
    return [f for f in list_files_recursively(path) if f.lower().endswith('.mmo')]


def xls_get_number_as_str(c):
    n = None
    if c.ctype == 2:  # Number
        iv = int(c.value)
        fv = c.value
        if iv == fv:
            n = str(iv)
        else:
            n = str(fv)
    else:
        n = str(c.value)
    return n


def xls_get_int(c):
    n = None
    if c.ctype == 2:  # Number
        return int(c.value)
    else:
        n = parse_int(str(c.value))
    return n


def xls_get_date(book, c):
    return datetime.datetime(
        *xlrd.xldate_as_tuple(c.value, book.datemode)
    ).date()


def xls_get_decimal(c):
    n = None
    if c.ctype == 2:  # Number
        n = Decimal(c.value)
    else:
        n = Decimal(str(c.value).replace("'", '').replace(',', '.'))
    return n


def xls_get_price(c):
    n = xls_get_decimal(c)
    return n.quantize(Decimal('1.00'))


def get_codes_from_report(path):
    book = xlrd.open_workbook(path)
    sh = book.sheet_by_index(0)
    codes = []
    for r in range(sh.nrows):
        c = sh.cell(rowx=r, colx=0)
        if c.ctype == 2:  # Number
            iv = int(c.value)
            fv = c.value
            if iv == fv:
                codes.append(str(iv))
            else:
                codes.append(str(fv))
        else:
            codes.append(str(c.value))
    return codes


def get_full_info_from_report(path):
    book = xlrd.open_workbook(path)
    sh = book.sheet_by_index(0)
    info = []
    for r in range(sh.nrows):
        c = sh.cell(rowx=r, colx=0)
        code = None
        if c.ctype == 2:  # Number
            iv = int(c.value)
            fv = c.value
            if iv == fv:
                code = str(iv)
            else:
                code = str(fv)
        else:
            code = str(c.value)
        title = sh.cell(rowx=r, colx=1)
        info.append((code, title,))
    return info


def get_info_from_member_report(path):
    book = xlrd.open_workbook(path)
    sh = book.sheet_by_index(0)

    c = sh.cell(rowx=0, colx=1)
    code = None
    if c.ctype == 2:  # Number
        iv = int(c.value)
        fv = c.value
        if iv == fv:
            code = str(iv)
        else:
            code = str(fv)
    else:
        code = str(c.value)

    rep_d = sh.cell_value(rowx=1, colx=1)
    try:
        rep_d_as_date = datetime.datetime(*xlrd.xldate_as_tuple(rep_d, book.datemode)).date()
    except Exception:
        rep_d_as_date = rep_d

    return {
        'code': code,
        'date': rep_d_as_date,
    }


def get_items_from_member_report(path):
    book = xlrd.open_workbook(path)
    sh = book.sheet_by_index(0)
    for r in range(sh.nrows):
        if r > 1:
            yield {
                'supplier': sh.cell_value(rowx=r, colx=0),
                'invoice_number': xls_get_number_as_str(sh.cell(rowx=r, colx=1)),
                'invoice_date': xls_get_date(book, sh.cell(rowx=r, colx=2)),
                'morion_id': xls_get_number_as_str(sh.cell(rowx=r, colx=3)),
                'title': sh.cell_value(rowx=r, colx=4),
                'maker': sh.cell_value(rowx=r, colx=5),
                'price': xls_get_price(sh.cell(rowx=r, colx=6)),
                'count': xls_get_int(sh.cell(rowx=r, colx=7)),
                'units': sh.cell_value(rowx=r, colx=8),
            }
