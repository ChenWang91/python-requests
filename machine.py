#!/usr/bin/env python
# -*- coding: utf-8 -*-

import requests
from lxml import etree
import xlwt
import sys
import time

def set_style(table):
    table.col(0).width = (30*100)
    for i in range(1,8):
        table.col(i).width = (30*200)


def get_list(html, xp, num):
    if num != 8:
        lis = html.xpath(xp + 'td[%d]/text()' %num)
        if len(lis) == 0:
            lis.append(' ')
    else:
        lis = html.xpath(xp + 'div/td[1]/text()')
        if len(lis) == 0:
            lis.append(' ')
    return lis

def get_message():
    total_machines = []
    mes = requests.get("http://xxxx.com")
    html = etree.HTML(mes.text)
    #print html.xpath('/html/body/div/div/div[2]/table/tr/td/text()')
    for i in range(1,45):
        dic = {}
        xp = '/html/body/div/div/div[2]/table/tr[%d]/' %i
        get_list(html,xp,1)[0]
        dic['Num'] = get_list(html,xp,1)[0]
        dic['Ip'] = get_list(html,xp,2)[0]
        dic['Machine'] = get_list(html,xp,3)[0]
        dic['Hard_message'] = get_list(html,xp,4)[0]
        dic['Os'] = get_list(html,xp,5)[0]
        dic['Rank'] = get_list(html,xp,6)[0]
        dic['Used'] = get_list(html,xp,8)[0]
        dic['Other'] = get_list(html,xp,7)[0]
        total_machines.append(dic)
    return total_machines
    
def write_excel(total_machines):
    title = ['ID','IP地址','主机名','硬件信息','系统型号','机柜编号','使用状态','备注']
    f = xlwt.Workbook()
    table = f.add_sheet('Machine Summary')
    
    print len(total_machines)

    set_style(table)
    #Write title
    for i in range(8):
        table.write(0,i,title[i].decode('utf-8'))
    
    for i in range(1,45):
        flag = 0
        for j in ['Num','Ip','Machine','Hard_message','Os','Rank','Used','Other']:
            table.write(i,flag,total_machines[i-1][j].decode('utf-8'))
            flag = flag+1

    fname = time.strftime('%Y-%m-%d')+'.xls'
    f.save(fname)
    print 'Save Status is OK!'

if __name__ == '__main__':
    reload(sys)
    sys.setdefaultencoding('utf-8')
    total_machines = get_message()
    write_excel(total_machines)
