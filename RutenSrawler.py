#!/usr/bin/python
# -*- coding: utf-8 -*-

import requests
import re
import sys
import json
import xlsxwriter
from bs4 import BeautifulSoup

head={
    'User-Agent':'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/47.0.2526.80 Safari/537.36',
    'Cookie':'_ts_id=34043601330336053502; _ga=GA1.3.2144250439.1440184992',
    'Referer':'http://mybid.ruten.com.tw/credit/point?pigfish0119&sell&all&3'
}

##檢查帳號是否存在##
def userIDcheck(userID, header):
    checkURL = requests.get("http://mybid.ruten.com.tw/credit/point?"+userID+"&sell&all" , headers = head)
    checkSoup = BeautifulSoup(checkURL.text)
    check=True
    for item in checkSoup.findAll("input", {'name': 'ms'}):
        if item:
            check = False
            print u"使用者帳號不存在".encode(type)
            print u"######################################".encode(type)
            end = raw_input(u'按任意鍵結束'.encode(type))
    return check

##檢查帳號是否有評價(評價幾筆)##
def pagecount(userID, header):
    res = requests.get("http://mybid.ruten.com.tw/credit/point?"+userID+"&sell&all" , headers = head)
    soup = BeautifulSoup(res.text)
    for item in soup.select('#table69'):
        if int(item.select('td')[4].text)<=0:
            print u"該帳號沒有賣出商品的評價".encode(type)
            print u"######################################".encode(type)
            end = raw_input(u'按任意鍵結束'.encode(type))
            return False
        else:
            return (int(item.select('td')[4].text)/20)+1
########################################################################
type = sys.getfilesystemencoding()
print u"露天拍賣爬蟲 Python x DennyChen".encode(type)
userID = raw_input(u"請輸入賣家帳號:".encode(type))
print u"######################################".encode(type)
if userIDcheck(userID,head):
    if pagecount(userID,head):
        print u"ing...".encode(type)
        print u"######################################".encode(type)
        workbook = xlsxwriter.Workbook(userID+'.csv')
        worksheet = workbook.add_worksheet('Hyperlinks')
        n=0
        for i in range(pagecount(userID,head)):
            url = "http://mybid.ruten.com.tw/credit/point?"+userID+"&sell&all&" + unicode(i)
            res = requests.get(url, headers = head)
            result = re.search('var f_list={"OrderList":(.*)?};',res.text)
            #print m.group(1)
            regex = re.compile(r'\\(?![/u"])')
            result_data = regex.sub(r"\\\\", result.group(1))
            data = json.loads(result_data)
            #print data
            for jdata in data:
                if not jdata['user']:
                    jdata['user'] = ("不公開").decode('utf-8')
                str = jdata['user'] + ',' + jdata['date'] + ',' + jdata['name'].encode('latin1', 'ignore').decode('big5') + ',' + jdata['money'].encode('latin1', 'ignore').decode('big5')
                #print str
                worksheet.write(n,0,str)
                n+=1
        workbook.close()
        print u"爬蟲完畢....請查看目錄底下".encode(type)
        print u"######################################".encode(type)
        end = raw_input(u'按任意鍵結束'.encode(type))