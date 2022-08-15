import os
import platform
import traceback

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# 若未设置环境变量，此处需手动输入Chrome Driver文件路径
ChromeDriver_PATH = 'PATH'


class WebDriver:
    def __init__(self, download_path=None):
        if os.path.isdir(str(download_path)):
            self.DOWNLOAD_PATH = download_path
        else:
            if str(download_path) != '':
                print(f'文件下载存放路径 "{download_path}" 不存在！')
                print('更改存放路径为当前目录下 "WOS_Downloads" 文件夹内。')
            self.DOWNLOAD_PATH = os.path.join(os.getcwd(), 'WOS_Downloads')
            if not os.path.isdir(self.DOWNLOAD_PATH):
                os.mkdir(self.DOWNLOAD_PATH)

        prefs = {'profile.default_content_settings.popups': 0, 'download.default_directory': self.DOWNLOAD_PATH}
        o = webdriver.ChromeOptions()
        o.add_experimental_option('prefs', prefs)
        o.add_experimental_option('excludeSwitches', ['enable-automation','enable-logging'])
        o.add_argument('--disable-extensions')
        o.add_argument('user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.0.0 Safari/537.36')

        local_path = os.path.join(os.getcwd(), 'chromedriver')

        if platform.system()=='Windows':
            print('Windows系统')
            local_path += '.exe'

        if os.path.exists(local_path):
            print('ChromeDriver路径：当前目录')
            self.driver = webdriver.Chrome(executable_path=local_path, options=o)
        elif os.path.exists(ChromeDriver_PATH):
            print(f'ChromeDriver路径：指定目录（{ChromeDriver_PATH}）')
            self.driver = webdriver.Chrome(executable_path=ChromeDriver_PATH, options=o)
        else:
            print('ChromeDriver路径：系统环境变量')
            self.driver = webdriver.Chrome(options=o)

    def getChrome(self):
        return self.driver

    def open_url(self, url):
        self.driver.get(url)

    def waitForLoaded(self, by, value, limit=10):
        try:
            element = WebDriverWait(self.driver, limit).until(EC.presence_of_element_located((by, value)))
        except Exception as e:
            traceback.print_exc()
            raise RuntimeError(f'Cannot find element: {value}')
        else:
            return element

    def isElementExist(self, by, value):
        try:
            element = self.driver.find_element(by, value)
        except NoSuchElementException:
            return False
        else:
            return True

    def currentUrl(self):
        return self.driver.current_url

    def refreshPage(self):
        self.driver.refresh()

    def __del__(self):
        try:
            self.driver.quit()
        except ImportError:
            pass
