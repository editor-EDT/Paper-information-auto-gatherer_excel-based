# -*- coding: utf-8 -*-
"""
Created on Sat Feb 13 17:33:13 2021
http://www.fenqubiao.com/
hfyjy1
hfyjy1

@authors: 童丹佛，童伟
@投资人：童伟
"""

from bs4 import BeautifulSoup
from selenium.webdriver import *
import time

def fengqubiao_login(username,password,browser):
    browser.get('http://www.fenqubiao.com')
    for x in username:
        browser.find_element_by_id("Username").send_keys(x)
        time.sleep(0.2)
    for x in password:
        browser.find_element_by_id("Password").send_keys(x)
        time.sleep(0.2)
    browser.find_element_by_id("login_button").click()
    # time.sleep(5)
    # browser.close()
def fengqubiao_parse(journal,browser,issn='0000-0000'):
    browser.maximize_window()
    browser.implicitly_wait(3)
    browser.find_element_by_link_text('检索').click()
    browser.implicitly_wait(3)
    browser.find_element_by_id("ContentPlaceHolder1_tbxTitleorIssn").send_keys(journal)
    browser.implicitly_wait(3)
    browser.find_element_by_id("ContentPlaceHolder1_btnSearch").click()
    time.sleep(3)
    # print(browser.page_source)
    # soup = BeautifulSoup(browser.page_source,'html5lib')
    try :
        soup = BeautifulSoup(browser.page_source,'html5lib')
        a = soup.find('table',attrs={'id':"detailJournal",'class':"detail table table-bordered"}).find_all('tr')
    except:
        # a = []
        table = []
        b = soup.find_all('tr')
        for x in range(len(b)):
            table.append([])
            for y in b[x].find_all('td'):
                table[x].append(y)
        table = table[1:]
        # print(table)
        for x in range(len(table)):
            # print(table[x][2].text)
            if table[x][2].text == issn:
                # print("Yes, Found it:", table[x][2].text)
                browser.find_element_by_xpath('//html//body//form//div[4]//div[2]//div//div//table//tbody//tr[%d]//td[2]//a'%(x+1)).click()
                break
        browser.implicitly_wait(3)
        current_window = browser.current_window_handle
        all_window = browser.window_handles
        for window in all_window:
            if window != current_window:
                browser.switch_to.window(window)
        browser.implicitly_wait(3)
        soup = BeautifulSoup(browser.page_source,'html5lib')
        # print('url: ', browser.current_url)
        a = soup.find('table',attrs={'id':"detailJournal",'class':"detail table table-bordered"}).find_all('tr')
        browser.close()
        browser.switch_to.window(all_window[0])
    table = []
    subject_name = ''
    fengqu = ''
    impact_factor = ''
    for x in range(len(a)):
        table.append([])
        for y in a[x].find_all('td'):
            table[x].append(y)
    for x in range(len(table)):
        for y in range(len(table[x])):
            if '大类' == table[x][y].text:
                subject_name = table[x][y+1].text
                fengqu = str(int(table[x][y+2].text))
            if '期刊影响因子' == table[x][y].text:
                impact_factor = table[x+2][y+2].text
    # if len(a) != 0:
    #     fengqu = fengqu[41]
    return {'subject_name':subject_name,'fengqu':fengqu,'impact_factor':impact_factor}
            
# print(111111111111111)
# options = ChromeOptions()
# options.add_argument("--auto-open-devtools-for-tabs")
# browser = Chrome(options=options)

# browser = Chrome()
# fengqubiao_login('hfyjy1','hfyjy1', browser)
# # a = input('请在继续前在Chrome上登录“中国科学院文献情报中心期刊分区表”（点击Enter继续）')
# print(fengqubiao_parse('JOURNAL OF MATERIALS SCIENCE', browser, '0022-2461'))
# print(fengqubiao_parse('Nature', browser))