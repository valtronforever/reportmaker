import pathlib
import xlrd


def list_files(path):
    files = []
    for currentFile in pathlib.Path(path).iterdir():
        if not currentFile.is_dir():
            files.append(str(currentFile))
    return files


def list_xls(path):
    return [f for f in list_files(path) if f.lower().endswith('.xls')]


def list_mmo(path):
    return [f for f in list_files(path) if f.lower().endswith('.mmo')]


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
