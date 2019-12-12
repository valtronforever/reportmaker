import pathlib


def list_files(path):
    for currentFile in pathlib.Path(path).iterdir():
        if not currentFile.is_dir():
            print(currentFile)


def list_xls(path):
    return [f for f in list_files(path) if f.lower().endswith('.xls')]


def list_mmo(path):
    return [f for f in list_files(path) if f.lower().endswith('.mmo')]
