#!/usr/bin/python
# -*- coding: UTF-8 -*-

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
            print u"使用者帳號不存在"
            print u"######################################"
            end = raw_input(u'按任意鍵結束')
    return check

##檢查帳號是否有評價(評價幾筆)##
def pagecount(userID, header):
    res = requests.get("http://mybid.ruten.com.tw/credit/point?"+userID+"&sell&all" , headers = head)
    soup = BeautifulSoup(res.text)
    for item in soup.select('#table69'):
        if int(item.select('td')[4].text)<=0:
            print u"該帳號沒有賣出商品的評價"
            print u"######################################"
            end = raw_input(u'按任意鍵結束')
            return False
        else:
            return (int(item.select('td')[4].text)/20)+1
########################################################################

print u"露天拍賣爬蟲 Python x DennyChen"
userID = sys.argv[1].strip()
#userID = "norns"
print u"Srawler "+ userID + " ing"
print u"######################################"
if userIDcheck(userID,head):
    if pagecount(userID,head):
		workbook = xlsxwriter.Workbook('Result_'+userID+'.csv')
		worksheet = workbook.add_worksheet('Hyperlinks')
		n=0
		TotalPage = pagecount(userID,head)
		for i in range(TotalPage):
			url = "http://mybid.ruten.com.tw/credit/point?"+userID+"&sell&all&" + unicode(i)
			res = requests.get(url, headers = head)
			res.encoding = "big5"      # 安安 這句很重要
			result = re.search('var f_list={"OrderList":(.*)?};',res.text)
			regex = re.compile(r'\\(?![/u"])')
			result_data = regex.sub(r"\\\\", result.group(1))
			data = json.loads(result_data, strict=False)
			print u"目前進度 "+ str(i+1) + " / " + str(TotalPage)
			for jdata in data:
				if not jdata['user']:
					jdata['user'] = ("不公開").decode('utf-8')
				_str = jdata['user'] + ',' + jdata['date'] + ',' + jdata['name'] + ',' + jdata['money']
				worksheet.write(n,0,_str)
				n+=1                
		workbook.close()
		print u"######################################"
		print u"爬蟲完畢....請查看目錄底下"