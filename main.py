import os
import traceback

import pandas as pd
from selenium.common.exceptions import WebDriverException

import WOS
import textProcessing as tp
from wordFig import wordFig

# 以下信息可选填
SEARCH_URL = ''  # 搜索结果的网址，填写后可自动打开搜索结果
DOWNLOAD_PATH = ''  # 文件下载存放路径，不填写默认下载到当前运行目录“WOS_Downloads”文件夹内

USER_NAME = ''  # 学号，用于自动登录
PASSWORD = ''  # 密码，用于自动登录

RETRY = 3  # 导出搜索结果的重试次数，若网络条件不佳可适当增大

EXCLUDE = ["Example1", "Example2"]  # 生成词云时需要除去的单词

MENU = [
    '''主菜单）
请选择功能编号：
    1. 自动登录
    2. 打开指定页面
    3. 保存部分搜索结果
    4. 保存全部搜索结果
    5. 合并搜索结果
    6. 生成图云
    0. 退出
''',
    '''自动登录）
    运行中...
''',
    '''打开指定页面）
    运行中...
''',
    '''保存部分搜索结果）
    运行中...
''',
    '''保存全部搜索结果）
    运行中...
''',
    '''合并搜索结果）
    运行中...
''',
    '''生成图云）
    生成中...
''',
]


def getOp(s):
    while True:
        try:
            op = int(input(s))
            return op
        except ValueError:
            print('错误！输入必须为数字')
            continue


def getInts(number):
    while True:
        try:
            lis = list(map(int, input(f'输入 {number} 个数字（使用空格隔开）：').split()))
            if len(lis) != number:
                print(f'错误！输入必须为 {number} 个数字。')
                continue
            return lis
        except ValueError:
            print('错误！输入必须为数字')
            continue


def ENTERtoResume(s=None):
    if s is not None:
        input(s)
    else:
        input('按下回车继续')


def formatURL(url):
    if not url.startswith('http'):
        return 'https://' + url
    else:
        return url


if __name__ == '__main__':
    browser = WOS.WOSDriver(DOWNLOAD_PATH)

    try:
        browser.open_url(formatURL(SEARCH_URL) if SEARCH_URL != '' else 'https://www.webofscience.com')
    except WebDriverException:
        print(f'无效的网址："{formatURL(SEARCH_URL)}"，请确认网址正确性。')
        browser.open_url('https://www.webofscience.com')

    while True:
        op = getOp(MENU[0])
        if op < 0 or op > 6:
            ENTERtoResume('无效的选择！请输入有效选项。')

        # 自动登录
        elif op == 1:
            print(MENU[1])
            try:
                browser.autoLogin(USER_NAME, PASSWORD)
            except Exception as e:
                print(e)
                input('自动登录失败！请重试或尝试手动登录。')
            else:
                input('完成，按回车返回主菜单。')

        # 打开页面
        elif op == 2:
            print(MENU[2])
            try:
                browser.open_url(SEARCH_URL)
            except Exception as e:
                print(e)
                input('打开页面失败！请重试或尝试手动打开。')
            else:
                input('完成，按回车返回主菜单。')

        # 部分保存
        elif op == 3:
            print(MENU[3])
            begin, end = getInts(2)
            flag = False
            try_count = 0
            while True:
                try:
                    browser.exportExcel(begin, end)
                    if end - begin > 999:
                        begin += 1000
                    else:
                        break
                except Exception as e:
                    try_count += 1
                    if try_count >= RETRY:
                        traceback.print_exc()
                        input('保存部分结果失败！请尝试重试。')
                        flag = True
                        break
                    browser.refreshPage()
            if not flag:
                input('完成，按回车返回主菜单。')

        # 全部保存
        elif op == 4:
            print(MENU[4])

            try_count = 0
            begin = 1
            end = 0
            flag = False

            while True:
                try:
                    end = browser.getResultsNumber()
                except Exception as e:
                    try_count += 1
                    if try_count >= RETRY:
                        print(e)
                        input('获取全部文献数量失败！请尝试重试。')
                        flag = True
                        break
                else:
                    break
            if flag:
                continue

            while True:
                try:
                    browser.exportExcel(begin, end)
                    if end - begin > 999:
                        begin += 1000
                    else:
                        break
                except Exception as e:
                    try_count += 1
                    if try_count >= RETRY:
                        print(e)
                        input('保存全部结果失败！请尝试重试。')
                        flag = True
                        break
                    browser.refreshPage()

            if not flag:
                input('完成，按回车返回主菜单。')

        # 合并搜索结果
        elif op == 5:
            print(MENU[5])
            files = tp.getFiles(browser.DOWNLOAD_PATH, 'xls')
            dt = pd.DataFrame()
            if not files:
                input('合并错误！搜索结果为空，请保存搜索结果后重试。')
                continue
            for file in files:
                tmp = pd.read_excel(file)
                dt = pd.concat([dt, tmp])
            dt.to_excel(os.path.join(browser.DOWNLOAD_PATH, 'AllRecords.xlsx'), index=None)
            input('完成，按回车返回主菜单。')

        # 生成词云
        elif op == 6:
            print(MENU[6])
            path = os.path.join(browser.DOWNLOAD_PATH, 'AllRecords.xlsx')
            if not os.path.exists(path):
                input('生成失败！文件不存在，请合并搜索结果后重试。')
                continue
            dt = pd.read_excel(path)
            textT = tp.getAllKeys(dt, 'Article Title')
            textA = tp.getAllKeys(dt, 'Abstract')

            figT = wordFig(textT, EXCLUDE)
            figA = wordFig(textA, EXCLUDE)

            print('当前显示标题词云。')
            figT.showPng('Title')
            print('当前显示摘要词云。')
            figA.showPng('Abstract')

            s = input('输入任意字符保存图片，或直接按下回车跳过：')
            if s:
                figT.savePng('Title.png')
                figA.savePng('Abstract.png')

            input('完成，按回车返回主菜单。')

        elif op == 0:
            break
