# -*- coding: utf-8 -*-
"""
Created on Sun Dec 27 10:47:20 2020
@authors: 童丹佛，童伟
@投资人：童伟
待修订
1. 将第一单位信息提取出来，填到表格中


"""

import xlwings,os,sys
import webofscience_parser
import fengqubiao_web_parser
def filling_sheet(location, sheet, restart = '', start_at = 1):
    print('starting...')
    app = xlwings.App(visible=True,add_book=False)
    book = app.books.open(location)
    sht = book.sheets[sheet]
    title = sht.range(2,3).expand('down').value
    #提取title
    a = restart
    #提取上次的记录，并用于还原
    if a != '':
        f = open("progress.txt",'w')
        f.write('%s:%s'%(a,start_at))
        f.close()
    f = open("progress.txt",'r')
    f.seek(0)
    last_break = f.read().split(':')
    f.close()
    print(last_break)
    if last_break[0] == 'wos':
        wos_start = int(last_break[1])
        fengqubiao_start = 1
    else:
        wos_start = len(title)+1
        fengqubiao_start = int(last_break[1])

    dirname, filename = os.path.split(os.path.abspath(sys.argv[0]))
    option = webofscience_parser.ChromeOptions()
    option.add_argument('--ignore-certificate-error')
    option.add_argument('--ignore-ssl-error')
    prefs = {"download.default_directory" : dirname}
    option.add_experimental_option("prefs",prefs)
    #option.add_argument('download.default_directory=E:/bank_py')
    browser = webofscience_parser.Chrome(executable_path='chromedriver.exe',options=option)
    f_NI = open('nature_index.txt','r')
    all_NI = [s.lower() for s in f_NI.readlines()]
    f_NI.close()
    print('start done')
    # '''提取title'''-----------------------------------------------------^
    print('searching...0/3')

    inf = []
    for x in range(wos_start-1,len(title)):
        f = open("progress.txt",'w')
        f.write('wos:'+str(x+1))
        f.close()
        inf.append(webofscience_parser.wos_parse(title[x],browser))
        sht.range(x+2,4).value = inf[x-wos_start+1]['authors']
        sht.range(x+2,5).value = inf[x-wos_start+1]['url']
        sht.range(x+2,6).value = inf[x-wos_start+1]['wos']
        sht.range(x+2,7).value = inf[x-wos_start+1]['ISSN']
        sht.range(x+2,8).value = inf[x-wos_start+1]['pub_date']
        sht.range(x+2,9).value = inf[x-wos_start+1]['journal_name']
        sht.range(x+2,10).value = inf[x-wos_start+1]['volume_issue']
        sht.range(x+2,12).value = inf[x-wos_start+1]['pub_type']
        sht.range(x+2,13).value = inf[x-wos_start+1]['subject_type']
        sht.range(x+2,14).value = inf[x-wos_start+1]['SCI_EI']
        sht.range(x+2,15).value = inf[x-wos_start+1]['corresponding_author']
        sht.range(x+2,16).value = inf[x-wos_start+1]['rank_HMFL']
        sht.range(x+2,17).value = inf[x-wos_start+1]['page']
        book.save()
        yield x-wos_start+2,len(title)-wos_start+1,'wos'
        print('No. %d/Total %d'%(x+1,len(title)))
    print('searching...1/3')
    inf2 = []
    journal_name = sht.range(2,9).expand('down').value
    ISSN = sht.range(2,7).expand('down').value
    pub_type = sht.range(2,12).expand('down').value
    fengqubiao_web_parser.fengqubiao_login('****','****',browser)
    for x in range(fengqubiao_start-1,len(title)):
        f = open("progress.txt",'w')
        f.write('fenqubiao:'+str(x+1))
        f.close()
        if pub_type[x] == '期刊论文':
            inf2.append(fengqubiao_web_parser.fengqubiao_parse(journal_name[x],browser,ISSN[x]))
        else:
            inf2.append({'subject_name':'无','fengqu':'无','impact_factor':'无'}) #不能用'none'或空'',否则excel读取时会截止到'none'
        if journal_name[x].lower()+'\n' in all_NI:
            inf2[x-fengqubiao_start+1]['NI'] = '是'
        else:
            print(x-fengqubiao_start+1)
            print(inf2) 
            inf2[x-fengqubiao_start+1]['NI'] = '否'
        sht.range(x+2,20).value = inf2[x-fengqubiao_start+1]['subject_name']
        sht.range(x+2,19).value = inf2[x-fengqubiao_start+1]['NI']
        sht.range(x+2,18).value = inf2[x-fengqubiao_start+1]['fengqu']
        sht.range(x+2,11).value = inf2[x-fengqubiao_start+1]['impact_factor']
        book.save()
        yield x-fengqubiao_start+2,len(title)-fengqubiao_start+1,'fengqu'

        print('No. %d/Total %d'%(x+1,len(title)))
    print('searching...2/3')
    f = open("progress.txt",'w')
    f.write('wos:1')
    f.close()
    book.save()
    book.app.quit()
    print('search done')
