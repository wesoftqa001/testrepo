import unittest
import sys
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, TimeoutException
import xlrd
import time
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.select import Select
import HTMLTestRunner

XLSPATH = 'D:/sample/Test sheet.xls'
XLSSHEET = 'Test'
PICPATH = 'D:/sample/1299.jpg'
REPORT_PATH = 'D:/report/'
REPORT_TIME = time.strftime("%Y-%m-%d %H%M%S", time.localtime())


''' return the bottom-right element of a n*n excel form'''
def get_excel_data(path, sheet=None):
    wb = xlrd.open_workbook(path)
    if sheet is None:
        worksheet = wb.sheet_by_index(0)
    else:
        worksheet = wb.sheet_by_name(sheet)

    return worksheet.cell(worksheet.nrows - 1, worksheet.ncols - 1).value


class Google(unittest.TestCase):
    def setUp(self):
        self.driver = webdriver.Chrome('chromedriver.exe')
        self.driver.maximize_window()
        #self.driver.get('http://google.com')
        time.sleep(2)

    @unittest.SkipTest
    def test_locate_to_gmail(self):
        try:
            self.driver.get('http://google.com')
            more_apps = self.driver.find_element_by_xpath('//*[@id="gbwabc"]/div[1]/a')
            more_apps.click()
            gmaillink = WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located(
                (By.XPATH, '//*[@id="gb23"]/span[1]')))
            gmaillink.click()

            # Goto gmail
            getstarted = WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located(
                (By.LINK_TEXT, 'GET STARTED')))
            self.driver.execute_script('arguments[0].scrollIntoView(true);', getstarted)
            time.sleep(3)
            getstarted.click()
            assert 'Sign', 'up' in self.driver.title.split(' ')
            self.assertTrue(EC.title_contains('Sign up'))
            print('case 1 passed')

        except (NoSuchElementException, TimeoutException) as e:
            raise e

    @unittest.SkipTest
    def test_read_xls(self):
        try:
            kw = get_excel_data(XLSPATH, XLSSHEET)
            print(kw)
            self.driver.find_element_by_name('q').send_keys(kw, Keys.ENTER)
            self.assertTrue(EC.title_contains('Python'))
            print('case 2 passed')

        except (NoSuchElementException, TimeoutException):
            print('case 2 failed')

    @unittest.SkipTest
    def test_upload_picture(self):
        try:
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="gbw"]/div/div/div[1]/div[2]/a'))).click()
            WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="qbi"]'))).click()
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, '#qbug > div > a'))).click()
            # input label button
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="qbfile"]'))).send_keys(PICPATH)
            self.assertIn('1299', self.driver.find_element_by_xpath('//div[@class="_hUb"]/a').text)
            time.sleep(3)
            print('case 3 passed')

        except BaseException as e:
            print(e)
            print('case 3 failed')

    @unittest.SkipTest
    def test_select_result_links(self):
        try:
            self.driver.find_element_by_name('q').send_keys('python', Keys.ENTER)
            self.driver.implicitly_wait(10)
            result_links = self.driver.find_elements_by_xpath('//h3[@class="r"]/a')
            print(len(result_links))
            for i in result_links:
                ActionChains(self.driver).move_to_element(i).perform()
                time.sleep(1)
            print('case 4 passed')

        except (NoSuchElementException, TimeoutException) as e:
            print('case 4 failed')
            print(e)

    @unittest.SkipTest
    def test_gmail_sign_up(self):
        try:
            self.driver.get(
                'https://accounts.google.com/SignUp?service=mail&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F')
            language = Select(self.driver.find_element_by_id('lang-chooser'))
            language.select_by_value('zh-CN')
            self.driver.implicitly_wait(10)
            self.driver.find_element_by_id('Gender').click()
            WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.ID, ':f'))).click()

            '''get data from excel'''
            self.driver.find_element_by_id('LastName').send_keys(get_excel_data(XLSPATH))
            self.driver.find_element_by_id('FirstName').send_keys(get_excel_data(XLSPATH, XLSSHEET))

            self.driver.find_element_by_id(':i').click()
            WebDriverWait(self.driver, 5).until(
                EC.presence_of_element_located((By.XPATH, '//div[contains(text(), "澳大利亚")]'))).click()
            self.driver.find_element_by_id('submitbutton').click()
            self.assertIsNotNone(self.driver.find_elements_by_class_name('errormsg'))  # assert there is error message in page
            time.sleep(5)

            '''direct to help page'''
            self.driver.find_element_by_link_text('详细了解').click()
            print(self.driver.window_handles)
            self.driver.switch_to.window(self.driver.window_handles[-1])  # switch to last window
            WebDriverWait(self.driver, 5).until(
                EC.presence_of_element_located((By.XPATH, '//span[contains(text(), "中文")]'))).click()
            self.driver.find_element_by_xpath('//li[contains(text(), "English")]').click()
            time.sleep(5)
            print('case 5 passed')

        except (NoSuchElementException, TimeoutException):
            raise

    #@unittest.SkipTest
    def test_jenkins(self):
        try:
            self.driver.get('http://baidu.com')
            self.driver.implicitly_wait(10)
            self.driver.find_element_by_xpath('//*[@id="u1"]/a[1]').click()
            time.sleep(3)
            #assert '123' in self.driver.title
            self.assertIn('新闻', self.driver.title)
        except (NoSuchElementException, TimeoutException) as e:
            raise e

    def test_jenkins_2(self):
        try:
            self.driver.get('http://baidu.com')
            self.driver.implicitly_wait(10)
            self.driver.find_element_by_xpath('//*[@id="u1"]/a[1]').click()
            time.sleep(1)
            self.driver.save_screenshot('scr/result.png')
            #assert '123' in self.driver.title
            self.assertIn('123', self.driver.title)
        except (NoSuchElementException, TimeoutException) as e:
            self.driver.save_screenshot('scr/result.png')
            raise e

    def tearDown(self):
        self.driver.quit()


if __name__ == "__main__":
    unittest.main()
'''
    file_path = REPORT_PATH + REPORT_TIME + '.html'
    suite = unittest.TestSuite()
    suite.addTest(Google('test_jenkins'))
    fp = open(file_path, 'wb')
    runner = HTMLTestRunner.HTMLTestRunner(stream=fp, title='report', description='demo')
    runner.run(suite)
    fp.close()'''

