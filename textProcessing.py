import os

import pandas as pd


def getFiles(path, filenameExtension, recursion=False):
    if not os.path.isdir(path):
        raise ValueError(f'"{path}" 不是一个文件夹或路径不存在。')
    result = []
    for filename in os.listdir(path):
        if filename.endswith(filenameExtension):
            fullpath = os.path.join(path, filename)
            if os.path.isfile(fullpath):
                result.append(fullpath)
            elif recursion:
                getFiles(fullpath, filenameExtension, True)
    return result


def mergeExcel(path):
    if not os.path.isdir(path):
        raise ValueError(f'"{path}" 不是一个文件夹或路径不存在。')
    result = pd.DataFrame()
    for file in getFiles(path, 'xls'):
        dt = pd.read_excel(file)
        result = pd.concat([result, dt])
    return result


def openExcel(file):
    return pd.read_excel(file)


def savetoExcel(data, filename):
    data.to_excel(filename, index=False)


def getAllKeys(data, key):
    result = ''
    for item in data[key]:
        result += str(item)
    return result
