from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException

import traceback
import os

# 若未设置环境变量，此处需手动输入Chrome Driver文件路径
ChromeDriver_PATH = "PATH"


class WebDriver:
    def __init__(self, download_path=None):
        if download_path is None or not os.path.exists(download_path):
            self.DOWNLOAD_PATH = os.getcwd() + "/WOS_Downloads"
            if not os.path.exists(self.DOWNLOAD_PATH):
                os.mkdir(self.DOWNLOAD_PATH)
        else:
            self.DOWNLOAD_PATH = download_path

        prefs = {'profile.default_content_settings.popups': 0, 'download.default_directory': self.DOWNLOAD_PATH}
        o = webdriver.ChromeOptions()
        o.add_experimental_option('prefs', prefs)
        o.add_argument("--disable-extensions")
        o.add_argument("user-agent:{}".format("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.0.0 Safari/537.36"))

        if os.path.exists(ChromeDriver_PATH):
            self.driver = webdriver.Chrome(executable_path=ChromeDriver_PATH, options=o)
        else:
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
            raise RuntimeError(f"Cannot find element: {value}")
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
        self.driver.quit()
