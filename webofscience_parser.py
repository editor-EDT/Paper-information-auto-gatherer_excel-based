# -*- coding: utf-8 -*-
"""
Created on Sat Feb 13 17:33:13 2021
@authors: 童丹佛，童伟
@投资人：童伟
1. 检索选择“标题”
2. 导出文本(plain text)savedrecs.txt
3. Affiliation: HMFL, High Magnetic Field Lab,  High Magnetic Field Laboratory, High Magnet Field Lab, High Field Magnet Lab
待修订：
1. 将第一单位信息提取出来，填到表格中

***
The method "tabled_dict" returns a dictionary with keys: 'authors', 'wos', 'ISSN', 'pub_date','pub_type', 'journal_name', 
'volume_issue', 'subject_type', 'corresponding_author', 'SCI_EI', 'rank_HMFL', 'page', 'url'.
***
"""

from selenium.webdriver import *
from selenium.common.exceptions import TimeoutException
import rebuild_title
from bs4 import BeautifulSoup
import time
import Wos_tex_parser
import os,sys




def wos_parse(title,browser):
    dirname, filename = os.path.split(os.path.abspath(sys.argv[0]))
    try:
        os.remove(dirname+"\\savedrecs.txt")
    except :
        pass
    title = rebuild_title.title_parse(title)
    browser.set_page_load_timeout(200)
    browser.get('http://apps.webofknowledge.com')
    #browser.get('https://www.webofscience.com/wos/alldb/basic-search')
    '''try:
        time.sleep(3)
        browser.find_element_by_id('pendo-close-guide-8fdced48').click()
    except:
        pass'''
    time.sleep(2)
    browser.find_element_by_id('onetrust-accept-btn-handler').click()
    browser.find_element_by_name('search-main-box').clear()
    browser.find_element_by_name('search-main-box').send_keys(title)
    browser.find_element_by_xpath('//button[@aria-label = "选择检索字段 主题"]').click()
    browser.find_element_by_xpath('//div[@aria-label = "标题"]').click()
    browser.find_element_by_xpath('//div[@class = "button-row"]//button[@data-ta = "run-search"]').click()
    time.sleep(3)
    soup = BeautifulSoup(browser.page_source,'html5lib')
    print(soup.find('span',attrs={'class':'mat-checkbox-label'}).text)
    try:
        
        if soup.find('span',attrs={'id':'footer_formatted_count'}).text != '1':
            for x in range(int(soup.find('span',attrs={'id':'footer_formatted_count'}).text)):
                a = soup.find('div',attrs = {'id':'RECORD_%d'%(x+1)})
                b = 0
                for y in a.find_all('span'):
                    # print(y.text)
                    if y.text == '卷:  \u200f' and 'Correction' not in a.find('value',attrs = {'lang_id':''}).text:
                        id_num = x+1
                        b = 1
                        browser.find_element_by_xpath('//div[@id = "RECORD_%d"]//a[@class = "smallV110 snowplow-full-record"]'%(x+1)).click()
                        break
                if b == 1:
                    break
        else:
            # print(1)
            browser.implicitly_wait(5)
            id_num = 1
            browser.find_element_by_xpath('//div[@id = "RECORD_1"]//a[@class = "smallV110 snowplow-full-record"]').click()
        browser.find_element_by_id("exportTypeName").click()
        try:
            browser.find_element_by_xpath('//li[@onclick = "trackSnowplowEventSE(\'save-to-other-file-formats-click\');"]//a').click()
        except :
            pass

        browser.find_element_by_id("exportButton").click()
        time.sleep(5)
        tex_parser = Wos_tex_parser.FileWosParser()
        
        info_temp = tex_parser.tabled_dict(tex_parser.parse(dirname+"\\savedrecs.txt")[0])
        os.remove(dirname+"\\savedrecs.txt")
        soup = BeautifulSoup(browser.page_source,'html5lib')
        if info_temp['subject_type'] == '':
            for x in soup.find_all('p'):
                if '研究方向:' in x.text:
                    info_temp['subject_type'] = x.text.replace('研究方向:','')[1:] #这里是否应该加个break
        if info_temp['corresponding_author'] == '':
            for x in soup.find_all('p'):
                if '通讯作者地址:' in x.text and '通讯作者地址:' != x.text:
                    # print([x.text])
                    corresponding_authors = []
                    for y in x.text[21:-8].split('; '):
                        # print(y)
                        if y not in corresponding_authors:
                            corresponding_authors.append(y)
                            info_temp['corresponding_author'] += y.replace(', ',' ')+', '
            info_temp['corresponding_author'] = info_temp['corresponding_author'][2:-2]
        if info_temp['rank_HMFL'] == '':
            info_temp['rank_HMFL'] = '0'
            affiliations = []
            for x in soup.find_all('table'):
                if '增强组织信息的名称' in x.text:
                    for y in x.find_all('a'):
                        # print([y.text])
                        if '[ ' in y.text:
                            affiliations.append(y.text)
            for x in range(len(affiliations)):
                if 'hmfl' in affiliations[x].lower() or 'high magnetic field lab' in affiliations[x].lower() or 'high magnet field lab' in affiliations[x].lower() or 'high field magnet lab' in affiliations[x].lower():
                    info_temp['rank_HMFL'] = affiliations[x][2]
                    break
        try:
            browser.find_element_by_id("full_text_%d"%id_num).click()
        except :
            pass
        current_window = browser.current_window_handle
        browser.find_element_by_xpath('//a[@title = "查看来自出版商的全文"]').click()
    #        time.sleep(5)
        all_window = browser.window_handles
        for window in all_window: #遇到长时间加载不完的网页，如何终止
            if window != current_window:
                browser.switch_to.window(window)
        
        browser.set_page_load_timeout(60)
        try:
            info_temp['url'] = browser.current_url
            #driver.get('http://www.baidu.com')
        except TimeoutException:
            browser.execute_script("window.stop();")
        info_temp['url'] = browser.current_url
        browser.close()
        browser.switch_to.window(all_window[0])
    except :
        #不能用'none'或空'',否则excel读取时会截止到'none', ''填写excel时会自动变为none
        info_temp = {'authors': '无', 'wos': '无', 'ISSN': '无', 'pub_date': '无','pub_type': '无', 'journal_name': '无','volume_issue': '无', 'subject_type': '无', 'corresponding_author': '无', 'SCI_EI': '无', 'rank_HMFL': '无', 'page': '无', 'url': '无'}
    return info_temp
browser = Chrome()
title = 'Ultra-fast synthesis of water soluble MoO3x quantum dots with controlled oxygen vacancies and their near infrared fluorescence sensing to detect H2O2'
print(wos_parse(title,browser))
# title = rebuild_title.title_parse('Research on radiation noise field reconstruction of underwater vehicle shell structure')
# print(wos_parse(title,browser))