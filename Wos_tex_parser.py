# -*- coding: utf-8 -*-
"""
Created on Sat Feb  6 10:42:08 2021

This module contains a Class "FileWosParser" with two methods "parse" and "tabled_dict".

The method "tabled_dict" returns a dictionary with keys: 'authors', 'wos', 'ISSN', 'pub_date', 'journal_name', 'pub_type', 
'volume_issue', 'subject_type', 'corresponding_author', 'SCI_EI', 'rank_HMFL', 'page'.

@authors: 童丹佛，童伟
@投资人：童伟
"""

class FileWosParser:
    def parse(self,*args):
        all_items = []
        for x in args:
            f = open(x,'r',encoding='utf8')
            inf = {}
            tex = f.readlines()
            #print(tex)
            tex[0] = tex[0][1:]
            key = ''
            value = []
            for x in tex:
                if x[0] != ' ':
                    inf[key] = value
                    key = x[:2]
                    value = [x[3:-1]]
                else:
                    value.append(x[3:-1])
            all_items.append(inf)
            f.close()
        return all_items
    def tabled_dict(self,dic):
        out_dic = {}
        
        
        out_dic['authors'] = ''
        for x in dic['AU']:
            out_dic['authors'] += x.replace(', ',' ')+', '
        out_dic['authors'] = out_dic['authors'][:-2]
        # print(out_dic['AU'])
        
        out_dic['wos'] = dic['UT'][0][4:]
        # print(out_dic['wos'])

        try:
            out_dic['journal_name'] = dic['SE'][0]
            out_dic['pub_type'] = '会议论文'
        except :
            out_dic['journal_name'] = dic['SO'][0]
            out_dic['pub_type'] = '期刊论文'

        try:
            out_dic['ISSN'] = dic['SN'][0]
        except :
            out_dic['ISSN'] = dic['EI'][0]
        # print(out_dic['ISSN'])
        
        if  out_dic['pub_type'] == '期刊论文':
            try:
                out_dic['pub_date'] = dic['PD'][0]
            except :
                out_dic['pub_date'] = dic['EA'][0]
            if len(out_dic['pub_date']) < 11:
                if len(out_dic['pub_date']) == 3:
                    out_dic['pub_date'] += ' 1 '+dic['PY'][0]
                elif len(out_dic['pub_date']) == 6 or len(out_dic['pub_date']) == 5:
                    out_dic['pub_date'] += ' '+dic['PY'][0]
                elif len(out_dic['pub_date']) == 8:
                    out_dic['pub_date'] = out_dic['pub_date'].replace(' ',' 1 ')
                    
            pub_date = out_dic['pub_date'].split(' ')
            mon_replace = {'JAN':'01','FEB':'02','MAR':'03','APR':'04','MAY':'05','JUN':'06','JUL':'07','AUG':'08','SEP':'09','OCT':'10','NOV':'11','DEC':'12'}
            out_dic['pub_date'] = pub_date[2]+'/'+mon_replace[pub_date[0]]+'/'+pub_date[1]

            try:
                out_dic['subject_type'] = dic['WC'][0]
                # print(out_dic['subject_type'])
            except :
                out_dic['subject_type'] = ''
        else:
            out_dic['pub_date'] = dic['PY'][0]+'/01/01'
            try:
                out_dic['subject_type'] = dic['WC'][0]
                # print(out_dic['subject_type'])
            except :
                out_dic['subject_type'] = dic['SO'][0]            
        # print(out_dic['pub_date'])
        
        # print(out_dic['journal_name'])
        # try:
        #     out_dic['volume_issue'] = '%s, %s (%s)'%(dic['VL'][0],dic['AR'][0],dic['PY'][0])
        #     out_dic['page'] = dic['AR'][0]
        #     # print(out_dic['volume_issue'])
        # except Exception as e:
        #     e = str(e)[1:-1]
        #     print([e])
        #     if e == 'VL':
        #         try:
        #             print(e)
        #             out_dic['volume_issue'] = ' , %s (%s)'%(dic['AR'][0],dic['PY'][0])
        #             out_dic['page'] = dic['AR'][0]
        #         except Exception as e:
        #             e = str(e)[1:-1]
        #             if e == 'AR':
        #                 out_dic['volume_issue'] = ' , %s-%s (%s)'%(dic['BP'][0],dic['EP'][0],dic['PY'][0])
        #                 out_dic['page'] = '%s-%s'%(dic['BP'][0],dic['EP'][0])
        #     elif e == 'AR':
        #         print(e)
        #         out_dic['volume_issue'] = '%s, %s-%s (%s)'%(dic['VL'][0],dic['BP'][0],dic['EP'][0],dic['PY'][0])
        #         out_dic['page'] = '%s-%s'%(dic['BP'][0],dic['EP'][0])
        try:
            VL = dic['VL'][0]
        except :
            VL = ' '
        try:
            pages = dic['AR'][0]
        except :
            try:
                pages = '%s-%s'%(dic['BP'][0],dic['EP'][0])
            except :
                pages = 'xxxx'
        
        try:
            year = dic['PY'][0]
        except :
            year = dic['EA'][0][4:]
        
        out_dic['volume_issue'] = '%s, %s (%s)'%(VL,pages,year)
        out_dic['page'] = pages
        
        

            
        try:
            out_dic['corresponding_author'] = dic['RP'][0].split(' (corresponding author)')
            slow_save = []
            for x in out_dic['corresponding_author']:
                a = x.split('; ')
                # print(a)
                for y in a:
                    if y not in slow_save and len(y)<10:
                        slow_save.append(y)
            out_dic['corresponding_author'] = ''
            for x in slow_save:
                out_dic['corresponding_author'] += x.replace(', ',' ')+', '
            out_dic['corresponding_author'] = out_dic['corresponding_author'][:-2]
        except :
            out_dic['corresponding_author'] = ''
        
        out_dic['rank_HMFL'] = '0'
        try:
            for x in range(len(dic['C1'])):
                if 'hmfl' in dic['C1'][x].lower() or 'high magnetic field lab' in dic['C1'][x].lower() or 'high magnet field lab' in dic['C1'][x].lower() or 'high field magnet lab' in dic['C1'][x].lower():
                    out_dic['rank_HMFL'] = '%d'%(x+1)
                    break
        except :
            out_dic['rank_HMFL'] = ''
        
        
        
        out_dic['SCI_EI'] = 'SCI'
        
        return out_dic
        

a = FileWosParser()
# print(a.parse('C://Users//Administrator//Downloads//savedrecs-会议1.txt')[0])
# print(a.tabled_dict(a.parse('C://Users//Administrator//Downloads//savedrecs-会议1.txt')[0]))
# print(a.parse('E://bank_py//savedrecs.txt')[0])
# print(a.tabled_dict(a.parse('E://bank_py//savedrecs.txt')[0]))                   
            
        
    