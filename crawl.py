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
import utils
import requests

# 用于记录屏幕输出
class Logger(object):
    def __init__(self, filename="crawl.log"):
        self.terminal = sys.stdout
        self.log = open(filename, "a")
 
    def write(self, message):
        self.terminal.write(message)
        self.log.write(message)
 
    def flush(self):
        pass
sys.stdout = Logger()

def get_proxy():
    return requests.get("http://127.0.0.1:5010/get/").json()

def delete_proxy(proxy):
    requests.get("http://127.0.0.1:5010/delete/?proxy={}".format(proxy))

def check_proxy(proxy):
    proxy = get_proxy().get("proxy")
    ok = True
    try:
        requests.get('http://httpbin.org/ip', proxies={"http": "http://{}".format(proxy)})
    except Exception:
        ok = False
    # 删除代理池中代理
    delete_proxy(proxy)
    return ok

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
path_chrome_driver = r'C:\Program Files (x86)\Google\Chrome\Application\chromedriver.exe'

def get_browser(type_browser, headless=False, use_proxy=False):
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
        options.add_argument('--disable-gpu')
        options.add_argument('--disable-infobars')
        options.add_argument('--disable-extentions')
        if use_proxy:
            proxy = get_proxy().get('proxy')
            while not check_proxy(proxy):
                proxy = get_proxy().get('proxy')
            print('Using proxy: {}'.format(proxy))
            options.add_argument('--proxy-server=%s' % proxy)
        browser = webdriver.Chrome(options=options, executable_path=path_chrome_driver)
        
    return browser

def start_crawl(journal, start_year, end_year, output_file):
    global type_browser, url, WAIT_SECONDS, MAX_NUM_PAPERS
    print('Start crawling papers from {} published during {} - {}'.format(journal, start_year, end_year))

    browser = get_browser(type_browser, headless=False, use_proxy=False)
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
        
    '''
    # 检查筛选后的文献总数是否超过1500条，超过1500条需要输入验证码，这里直接放弃
    total = str2int(em_total.text)
    if total > MAX_NUM_PAPERS:
        print('Number of papers: {} > {}. Journal: {}, year: {} - {}.'.format(total, MAX_NUM_PAPERS, journal, start_year, end_year))
        browser.quit()
        return False
    '''
    
    # 找到每页文献数标签并点击使其展开
    try:
        div_perpage = browser.find_element_by_id('perPageDiv').find_element_by_tag_name('div')
    except NoSuchElementException as e:
        print(str(e))
        browser.quit()
        return False
    browser.execute_script("arguments[0].scrollIntoView();", div_perpage) 
    div_perpage.click()
    
    # 等待文本为50的标签加载完成并点击
    try:
        li_50 = WebDriverWait(browser, WAIT_SECONDS).until(
            EC.element_to_be_clickable(
                (By.XPATH, '//div[@id="perPageDiv"]/ul/li[@data-val="50"]')
            )
        )
    except TimeoutException as e:
        print('Timeout during waiting for loading of buttom perPageDiv. Journal: {}, year: {} - {}.'.format(journal, start_year, end_year))
        browser.quit()
        return False 
            
    # 寻找每页文献数量标签
    try:
        span = browser.find_element_by_id('perPageDiv').find_element_by_tag_name('span') 
    except NoSuchElementException as e:
        print(str(e))
        browser.quit()
        return False
        
    li_50.click()

    # 通过每页文献数量标签的过期，判断文献列表刷新完成
    try:
        WebDriverWait(browser, WAIT_SECONDS).until(EC.staleness_of(span))
    except TimeoutException as e:
        print('Timeout during waiting for refresh of search results after clciking buttom perPageDiv. Journal: {}, year: {} - {}.'.format(journal, start_year, end_year))
        browser.quit()
        return False

    cols = ['篇名', '作者', '期刊名称', '发表时间', '被引次数', '被下载次数'] # 需要保存的信息种类
    col2index = {}
    for i, col in enumerate(cols):
        col2index[col] = i + 1
    result = []
    page_cnt = 0
    # 保存所有页的文献信息
    while True:
        page_cnt +=1
        # 文献信息保存在表格中，表格的一行(tr)对应一篇文献信息
        try:
            trs = browser.find_elements_by_xpath('//*[@id="gridTable"]/table/tbody/tr')
        except NoSuchElementException as e:                         
            print(str(e))
            browser.quit()
            return False
        # 保存当前页的所有文献信息
        for tr in trs:
            # 表格的一列(td)对应一篇文献的特定信息，如篇名、作者
            try: 
                tds = tr.find_elements_by_tag_name('td')
            except NoSuchElementException as e:
                print(str(e))
                browser.quit()
                return False
            '''
            # 检查期刊名是否符合预期(可能模糊匹配)
            if not(tds[3].text == journal or tds[3].text.startswith(journal + '(')):
                print('Actual source: {}, expected source: {}'.format(tds[3].text, journal))
                continue
            # 检查发表时间是否符合预期
            y = int(tds[4].text.split('-')[0])
            if not check_date(y, start_year, end_year):
                print('Actual year: {}, expected year: {} - {}. Journal: {}, year: {} - {}.'.format(y, start_year, end_year, journal, start_year, end_year))
                continue
            '''
            info = OrderedDict()
            info['时间段'] = '2012-2020'
            for col in cols:
                if col == '发表时间':
                    info[col] = tds[col2index[col]].text
                else:
                    try:
                        a = tds[col2index[col]].find_element_by_tag_name('a')
                    except NoSuchElementException as e:
                        info[col] = ''
                        continue
                    info[col] = a.text
            result.append(info)

        # 寻找下一页按键
        try:
            next_page = WebDriverWait(browser, WAIT_SECONDS).until(
            EC.element_to_be_clickable(
                    (By.ID, 'PageNext')
                )
            )
        except TimeoutException:
            break

        # 将鼠标拖动到下一页按键附近并点击
        browser.execute_script("arguments[0].scrollIntoView();", next_page) 
        span = browser.find_element_by_xpath('//span[@class="cur"]')
        ActionChains(browser).move_to_element(next_page).click().perform()
        
        # 30页必出验证码
        if (page_cnt % 30)== 0:
            time.sleep(15)
        
        # 通过当前页码标签判断页面是否刷新
        try:
            WebDriverWait(browser, WAIT_SECONDS).until(EC.staleness_of(span))  
        except TimeoutException as e:
            print('Timeout during waiting for refresh of current page after clicking next page. Journal: {}, year: {} - {}.'.format(journal, start_year, end_year))
            browser.quit()
            return False        
    
    # 将爬取结果保存到excel中
    if len(result):
        df = pd.DataFrame(data=result)
        df.to_excel(output_file) 

    browser.quit()
    print('Finish crawling papers from {} published in year: {} - {}, number of papers: {}'.format(journal, start_year, end_year, len(result)))
    return len(result) > 0

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
    output_dir = './output'
    if not os.path.exists(output_dir):
        os.mkdir(output_dir)
    
    start_years = [2012, 2015, 2018]
    end_years = [2014, 2017, 2020]
    
    cnt, total = 0, (end - start + 1) * len(start_years)
    succeed = failed = skipped = 0
    for i in range(end - start + 1):
        for j in range(len(start_years)):
            cnt += 1
            output_file = output_dir + '/' + journals[i] + str(j + 1) + '.xlsx'
            if not Path(output_file).is_file():
                if start_crawl(journals[i], start_years[j], end_years[j], output_file):
                    succeed += 1
                else:
                    failed += 1
                print('Progress: {}/{}, succeed: {}, failed: {}, skipped: {}, used time: {}'.format(cnt, total, succeed, failed, skipped, time.time() - start_time))
            else:
                skipped += 1
    '''
    # 找出缺失的(未能爬下来的)文件，并把多出来的文件移动到其他目录
    src = './output' + '_' + str(start) + '_' + str(end)
    src = './output'
    _, missing_files = utils.find_extra_missing(src, journals, './others')
    cnt, total = 0, len(missing_files) * len(start_years)
    succeed = failed = skipped = 0
    '''
    '''
    # 按每三年一个文件爬取
    for i in range(len(missing_files)):
        cnt += 1
        output_file = output_dir + '/' + journals[i]
        journal, j = missing_files[i][:-6], int(missing_files[i][-6]) - 1
        if not Path(output_file).is_file():
            if start_crawl(journal, start_years[j], end_years[j], output_file):
                succeed += 1
            else:
                failed += 1
            print('Progress: {}/{}, succeed: {}, failed: {}, skipped: {}, used time: {}'.format(cnt, total, succeed, failed, skipped, time.time() - start_time))
        else:
            skipped += 1 
    '''
    '''
    # 按每一年一个文件爬取
    for i in range(len(missing_files)):
        lst = missing_files[i].split('.')
        journal, j = missing_files[i][:-6], int(missing_files[i][-6]) - 1
        for year in range(start_years[j], end_years[j] + 1):
            cnt += 1
            output_file = output_dir + '/' + journal + str(year) + '.xlsx'
            if not Path(output_file).is_file():
                if start_crawl(journal, year, year, output_file):
                    succeed += 1
                else:
                    failed += 1
            else:
                skipped += 1
            print('Progress: {}/{}, succeed: {}, failed: {}, skipped: {}, used time: {}'.format(cnt, total, succeed, failed, skipped, time.time() - start_time))
    '''
    print('Finished crawl. Total succeed: {}, total failed: {}, total skipped: {}, total used time: {}'.format(succeed, failed, skipped, time.time() - start_time))                
    
if __name__ == '__main__':
    main()
