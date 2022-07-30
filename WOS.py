from selenium.webdriver.common.by import By
import driver
from time import sleep
from bs4 import BeautifulSoup
import os


class WOSDriver(driver.WebDriver):
    def __init__(self, download_path=None):
        driver.WebDriver.__init__(self, download_path)

    def clearBanners(self):
        try:
            chrome = self.getChrome()
            page = chrome.page_source.encode('UTF-8')
            page = BeautifulSoup(page, 'lxml')
            if page.find_all('button', class_ = 'cookie-setting-link'):
                chrome.find_element(By.CSS_SELECTOR, '#onetrust-accept-btn-handler').click()
                # sleep(5)

            if page.find_all('button', class_ = 'bb-button _pendo-button-primaryButton _pendo-button'):
                chrome.find_element(By.CSS_SELECTOR, '#pendo-button-59b176ac').click()
        except Exception as e:
            print("无法清除横幅！")
            input("请手动清除横幅后按回车键继续。")

    def autoLogin(self, username, password):
        select_element = self.waitForLoaded(By.CSS_SELECTOR, ".mat-select-arrow").click()
        self.waitForLoaded(By.CSS_SELECTOR, "#mat-option-9 span:nth-child(1)").click()
        self.waitForLoaded(By.CSS_SELECTOR, "button.wui-btn--login:nth-child(4) span:nth-child(1) span:nth-child(1)").click()
        self.waitForLoaded(By.CSS_SELECTOR, "#show").send_keys("上海大学")
        sleep(1)
        self.waitForLoaded(By.CSS_SELECTOR, "#idpForm > div.unit1 > div > div > div.ipt > ul > li:nth-child(1)").click()
        self.waitForLoaded(By.CSS_SELECTOR, "#idpSkipButton").click()
        self.waitForLoaded(By.CSS_SELECTOR, "#username").send_keys(username)
        self.waitForLoaded(By.CSS_SELECTOR, "#password").send_keys(password)
        self.waitForLoaded(By.CSS_SELECTOR, "#submit-button").click()
        sleep(2)

    def exportExcel(self, begin, end=-1):
        if end - begin >= 1000 or end == -1:
            end = begin + 999
        self.clearBanners()
        self.waitForLoaded(By.CSS_SELECTOR, "#snRecListTop > app-export-menu > div > button").click()
        sleep(1)
        self.waitForLoaded(By.CSS_SELECTOR, "#exportToExcelButton").click()
        sleep(0.5)

        try:
            self.waitForLoaded(By.XPATH, "//*[text()='Records from:']").click()
            sleep(0.5)

            self.waitForLoaded(By.XPATH, "//input[@name='markFrom']").clear()
            self.waitForLoaded(By.XPATH, "//input[@name='markFrom']").send_keys(str(begin))
            self.waitForLoaded(By.XPATH, "//input[@name='markTo']").clear()
            self.waitForLoaded(By.XPATH, "//input[@name='markTo']").send_keys(str(end))

            self.waitForLoaded(By.XPATH, "//button[@aria-label=' Author, Title, Source']").click()
            sleep(1)

            self.waitForLoaded(By.XPATH, "//*[text()=' Edit ']").click()
            sleep(0.5)

            self.waitForLoaded(By.XPATH, "//*[text()='Author, Title, Source ']").click()
            self.waitForLoaded(By.XPATH, "//*[text()='Author, Title, Source ']").click()
            self.waitForLoaded(By.XPATH, "//*[text()='Abstract and Other ']").click()
            self.waitForLoaded(By.XPATH, "//*[text()='Abstract and Other ']").click()

            # 选择导出信息（可自定义添加）
            self.waitForLoaded(By.XPATH, "//*[text()='Author(s) ']").click()
            self.waitForLoaded(By.XPATH, "//*[text()='Title ']").click()
            self.waitForLoaded(By.XPATH, "//*[text()='Abstract ']").click()

            self.waitForLoaded(By.XPATH, "//*[text()=' Save selections ']").click()
            sleep(0.1)

            self.waitForLoaded(By.XPATH, "//*[text()='Export']").click()
        except Exception as e:
            self.waitForLoaded(By.XPATH, "//button[@aria-label='Close']")
            raise e

        except_count = 10
        print(f"正在等待下载（{begin}-{end}）   ", end="")
        while any([filename.endswith(".crdownload") for filename in os.listdir(self.DOWNLOAD_PATH)]) or self.isElementExist(By.XPATH, "//button[@disabled='true' and @class='mat-focus-indicator cdx-but-md mat-stroked-button mat-button-base mat-primary mat-button-disabled']"):
            sleep(1)
            print("\b\b\b.  ", end="")
            sleep(1)
            print("\b\b\b.. ", end="")
            sleep(1)
            print("\b\b\b...", end="")
            sleep(1)
            print("\b\b\b   ", end="")
            except_count -= 1
            if except_count < 0:
                raise RuntimeError("下载失败！")
        os.rename(self.DOWNLOAD_PATH + "/savedrecs.xls", self.DOWNLOAD_PATH + f"/Records{begin}-{end}.xls")
        print("\b\b下载完成。")

    def getResultsNumber(self):
        return int(self.waitForLoaded(By.CSS_SELECTOR, "body > app-wos > div > div > main > div > div > div.held > app-input-route > app-base-summary-component > app-search-friendly-display > div.search-display > app-general-search-friendly-display > h1 > span").text.replace(",", ""))
