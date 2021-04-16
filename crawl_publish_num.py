#!/usr/bin/env python3
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
import pandas as pd
import time
import random
from collections import OrderedDict
import sys
import os
from pathlib import Path
from multiprocessing import Pool
import utils

# 用于记录屏幕输出
class Logger(object):
    def __init__(self, filename="crawl_publish_num.log"):
        self.terminal = sys.stdout
        self.log = open(filename, "a")
 
    def write(self, message):
        self.terminal.write(message)
        self.log.write(message)
 
    def flush(self):
        pass
sys.stdout = Logger()

# 将每三位以逗号分隔的字符串表示的数字转换成阿拉伯数字
def str2int(s):
    nums = [int(i) for i in s.split(',')]
    res = 0
    for num in nums:
        res = res * 1000 + num
    return res

# 检查年份y是否为在start与end之间
def check_date(y, start, end):
    for year in range(start, end + 1):
        if y == year:
            return True
    return False

url = 'https://chn.oversea.cnki.net/'
WAIT_SECONDS = 15
MAX_NUM_PAPERS = 1500

type_browser = 'chrome'  # 浏览器类型(目前仅支持chrome和firefox)
# 浏览器驱动路径
path_firefox_driver = '/home/panda/Downloads/geckodriver'
path_chrome_driver = '/home/panda/Downloads/chromedriver'

def get_browser(type_browser, headless=False):
    global path_firefox_driver, path_chrome_driver
    browser = None

    if type_browser == 'firefox':
        from selenium.webdriver.firefox.options import Options
        options = Options()
        if headless:
            options.add_argument('--headless')  
        browser = webdriver.Firefox(options=options, executable_path=path_firefox_driver) 
    elif type_browser == 'chrome':
        from selenium.webdriver.chrome.options import Options
        options = Options()
        if headless:
            options.add_argument('--headless')
        options.add_argument('--no-sandbox')
        options.add_argument('start-maximized')
        options.add_argument('--disable-infobars')
        options.add_argument('--disable-extentions')
        browser = webdriver.Chrome(options=options, executable_path=path_chrome_driver)
        
    return browser

def start_crawl(journal, start_year, end_year, output_file):
    global type_browser, url, WAIT_SECONDS, MAX_NUM_PAPERS
    print('Start crawling publish number of journal {} during {} - {}'.format(journal, start_year, end_year))

    browser = get_browser(type_browser, headless=False)
    browser.get(url)

    # 等待高级检索按键加载完成并点击
    try:
        a_high_search = WebDriverWait(browser, WAIT_SECONDS).until(
            EC.element_to_be_clickable(
                (By.LINK_TEXT, '高级检索')
            )
        )
    except TimeoutException as e:
        print('Timeout during waiting for loading of high search buttom.')
        browser.quit()
        return False
    a_high_search.click()
    time.sleep(1)

    # 切换到最新打开的窗口
    windows = browser.window_handles
    browser.switch_to.window(windows[-1])

    # 找到文献来源标签，用于后面判断点击刷新完成
    try:
        span_src = browser.find_element_by_xpath('//*[@id="gradetxt"]/dd[3]/div[2]/div[1]/div[1]/span')
    except NoSuchElementException as e:
        print(str(e))
        browser.quit()
        return False

    # 等待学术期刊按键加载完成并点击
    try:
        span_journal = WebDriverWait(browser, WAIT_SECONDS).until(
            EC.element_to_be_clickable(
                (By.XPATH, '//ul[@class="doctype-menus keji"]/li[@data-id="xsqk"]/a/span')
            )
        )
    except TimeoutException as e:
        print('Timeout during waiting for loading of journal buttom.')
        browser.quit()
        return False
    browser.execute_script('arguments[0].click();', span_journal)

    # 通过文献来源标签的过期判断刷新完成
    try:
        WebDriverWait(browser, WAIT_SECONDS).until(
            EC.staleness_of(span_src)
        )
    except TimeoutException as e:
        print('Timeout during refresh after clicking journal buttom')
        browser.quit()
        return False

    # 等待期刊名称输入框加载完成
    try:
        input_journal = WebDriverWait(browser, WAIT_SECONDS).until(
            EC.presence_of_element_located(
                (By.XPATH, '//*[@id="gradetxt"]/dd[3]/div[2]/input')
            )
        )
    except TimeoutException as e:
        print('Timeout during waiting for input box for jounral name.')
        browser.quit()
        return False
    input_journal.send_keys(journal)  # 输入期刊名
    time.sleep(1)

    # 找到起始年输入框并输入起始年
    try:
        input_start_year = browser.find_element_by_xpath('//input[@placeholder="起始年"]')
    except NoSuchElementException as e:
        print(str(e))
        browser.quit()
        return False
    input_start_year.send_keys(start_year)

    # 找到结束输入框并输入结束年
    try:
        input_end_year = browser.find_element_by_xpath('//input[@placeholder="结束年"]')
    except NoSuchElementException as e:
        print(str(e))
        browser.quit()
        return False
    input_end_year.send_keys(end_year)

    # 找到检索键并点击
    try:
        input_search = browser.find_element_by_xpath('//input[@value="检索"]')
    except NoSuchElementException as e:
        print(str(e))
        browser.quit()
        return False
    input_search.click()

    # 根据搜索结果总数标签的出现判断页面是否刷新完成
    try:
        em_total = WebDriverWait(browser, WAIT_SECONDS).until(
            EC.visibility_of_element_located(
                (By.XPATH, '//*[@id="countPageDiv"]/span[1]/em')
            )
        )
    except TimeoutException as e:
        print('Timeout during waiting for paper number cell.')
        browser.quit()
        return False
    expected_num = str2int(em_total.text)

    # 找到发表年度
    try:
        dt_year = WebDriverWait(browser, WAIT_SECONDS).until(
            EC.element_to_be_clickable(
                (By.XPATH, '//dt[@groupitem="发表年度"]')
            )
        )
    except TimeoutException as e:
        print('Timeout during waiting for loading of publish years list. Journal: {}, year: {} - {}.'.format(journal, start_year, end_year))
        browser.quit()
        return False
    
    # 将鼠标移动到发表年度，若发表年度被折叠，则点击之
    action = ActionChains(browser).move_to_element(dt_year)
    dt_year_parent = browser.find_element_by_xpath('//dt[@groupitem="发表年度"]/..')
    if dt_year_parent.get_attribute('class') == 'is-up-fold off':
        action.click().perform()  # 点击发表年度

    # 让鼠标悬浮在发表年度的部分展开列表上，使其完全展开
    try:
        div_lst = WebDriverWait(browser, WAIT_SECONDS).until(
            EC.visibility_of_element_located(
                (By.XPATH, '//dd[@tit="发表年度"]/div')
            )
        )
    except TimeoutException as e:
        print('Timeout during waiting for display of complete publish years. Journal: {}, year: {} - {}.'.format(journal, start_year, end_year))
        browser.quit()
        return False
    ActionChains(browser).move_to_element(div_lst).perform()      

    info = OrderedDict()
    info['期刊名称'] = journal
    actual_num = 0
    timeout = False
    # 找到目标范围内的年份的发表数量
    for year in range(start_year, end_year + 1):
        if not timeout or year == 2014:
            try:
                span = WebDriverWait(browser, WAIT_SECONDS).until(
                    EC.visibility_of_element_located(
                        (By.XPATH, '//input[@type="checkbox" and @text="{0}" and @value="{0}"]/following-sibling::span'.format(str(year)))
                    )
                )
            except TimeoutException as e:
                print('No paper found for journal {0} in year {1}. Journal: {0}, year: {2} - {3}.'.format(journal, year, start_year, end_year))
                info[str(year)] = 0
                timeout = True
                continue
        else:
            try:
                span = browser.find_element_by_xpath('//input[@type="checkbox" and @text="{0}" and @value="{0}"]/following-sibling::span'.format(str(year)))
            except NoSuchElementException as e:
                print(str(e))
                info[str(year)] = 0
                continue
        #print('year: {}, num: {}'.format(year, span.text))
        num = int(span.text[1:-1])
        info[str(year)] = num   
        actual_num += num

    # 检验九年发表数量总和是否等于筛选结果数量
    if actual_num != expected_num:
        print('Total number of papers conflict, expected number: {}, actual number: {}. Journal: {}, year: {} - {}.'.format(expected_num, actual_num, journal, start_year, end_year))
        browser.quit()
        return False
    
    info['总数'] = expected_num
    # 将爬取结果保存到excel中
    df = pd.DataFrame(data=[info])
    df.to_excel(output_file) 
    
    browser.quit()
    print('Finish crawling publish number of journal {} during {} - {}, expected number of papers: {}'.format(journal, start_year, end_year, expected_num))
    return True

def main():
    if (len(sys.argv)) != 3:
        print('Usage: python3 crawl_cnki.py start end')
        return
    else:
        start, end = int(sys.argv[1]), int(sys.argv[2])

    start_time = time.time()
    print(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    
    # 将excel文件中的期刊名转存为txt文件，并读取出范围为start到end的期刊名
    ori_src, new_src = './待爬取数据.xlsx', './journals.txt'
    utils.excel2txt(ori_src, new_src)
    journals = utils.read_txt(new_src, start, end)

    # 创建用于保存输出结果的目录
    output_dir = './publish_numbers'
    if not os.path.exists(output_dir):
        os.mkdir(output_dir)
    
    start_year, end_year = 2012, 2020
    cnt, total = 0, end - start + 1
    succeed = failed = skipped = 0
    
    for i in range(end - start + 1):
        cnt += 1
        output_file = output_dir + '/' + journals[i] + '.xlsx'
        if not Path(output_file).is_file():
            if start_crawl(journals[i], start_year, end_year, output_file):
                succeed += 1
            else:
                failed += 1
            print('Progress: {}/{}, succeed: {}, failed: {}, skipped: {}, used time: {}'.format(cnt, total, succeed, failed, skipped, time.time() - start_time))
        else:
            skipped += 1
            
    print('Finished crawl. Total succeed: {}, total failed: {}, total skipped: {}, total used time: {}'.format(succeed, failed, skipped, time.time() - start_time))                
    
if __name__ == '__main__':
    main()
