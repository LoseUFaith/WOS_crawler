import WOS
from time import sleep

# 下面信息可选填，填写后可提升自动化程度
SEARCH_URL = ''  # 搜索结果的网址，填写后可自动打开搜索结果
DOWNLOAD_PATH = ''  # 下载后存放路径，不填写默认下载到当前运行目录“WOS_Downloads”文件夹内

USER_NAME = ''  # 学号，用于自动登录
PASSWORD = ''  # 密码，用于自动登录

MENU = [
    """主菜单）
请选择功能编号：
    1. 自动登录
    2. 打开指定页面
    3. 保存部分搜索结果
    4. 保存全部搜索结果
    0. 退出
""",
    """自动登录）
    运行中...
""",
    """打开指定页面）
    运行中...
""",
    """保存部分搜索结果）
    运行中...
""",
    """保存全部搜索结果）
    运行中...
""",
]

RETRY = 3  # 导出搜索结果的重试次数，若网络条件不佳可适当增大


def getOp(s):
    while True:
        try:
            op = int(input(s))
            return op
        except ValueError:
            input("错误！输入必须为数字")
            continue


def getInts(number):
    while True:
        try:
            lis = list(map(int, input(f"输入 {number} 个数字（使用空格隔开）：").split()))
            if len(lis) != number:
                input(f"错误！输入必须为 {number} 个数字。")
                continue
            return lis
        except ValueError:
            input("错误！输入必须为数字")
            continue


def ENTERtoResume(s=None):
    if s is not None:
        input(s)
    else:
        input("按下回车继续")


if __name__ == '__main__':
    browser = WOS.WOSDriver(DOWNLOAD_PATH)
    browser.open_url(SEARCH_URL if SEARCH_URL != "" else "https://www.webofscience.com")
    while True:
        op = getOp(MENU[0])
        if op < 0 or op > 4:
            ENTERtoResume("无效的选择！请输入有效选项。")
            continue
        if op == 1:
            print(MENU[1])
            try:
                browser.autoLogin(USER_NAME, PASSWORD)
            except Exception as e:
                print(e)
                input("自动登录失败！请重试或尝试手动登录。")
                continue
            else:
                input("完成，按回车返回主菜单。")
                continue
        if op == 2:
            print(MENU[2])
            try:
                browser.open_url(SEARCH_URL)
            except Exception as e:
                print(e)
                input("打开页面失败！请重试或尝试手动打开。")
                continue
            else:
                input("完成，按回车返回主菜单。")
                continue
        if op == 3:
            print(MENU[3])
            begin, end = getInts(2)
            flag = False
            try_count = 0
            while True:
                try:
                    browser.exportExcel(begin, end)
                except Exception as e:
                    try_count += 1
                    if try_count >= RETRY:
                        print(e)
                        input("保存部分结果失败！请尝试重试。")
                        flag = True
                        break
                    browser.refreshPage()
                if end - begin > 999:
                    begin += 1000
                else:
                    break
            if not flag:
                input("完成，按回车返回主菜单。")
            continue
        if op == 4:
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
                        input("获取全部文献数量失败！请尝试重试。")
                        flag = True
                        break
                else:
                    break
            if flag:
                continue

            while True:
                try:
                    browser.exportExcel(begin, end)
                except Exception as e:
                    try_count += 1
                    if try_count >= RETRY:
                        print(e)
                        input("保存全部结果失败！请尝试重试。")
                        flag = True
                        break
                    browser.refreshPage()

                if end - begin > 999:
                    begin += 1000
                else:
                    break
            if not flag:
                input("完成，按回车返回主菜单。")
            continue
        if op == 0:
            break

    exit(0)
