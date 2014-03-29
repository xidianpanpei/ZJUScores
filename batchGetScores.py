#!/usr/bin/env python
#-*-coding:utf-8-*-

import urllib
import urllib2
import cookielib
import re
import string
import xlrd

headers = {
    'User-Agent': 'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:21.0 Gecko/20100101 Firefox/21.0',
    'Accept-Language':'en-US,en;q=0.5',
    'Accept-Encoding': 'gzip, deflate',
    'Content-Type': 'application/x-www-form-urlencoded',
    'Host': '112.15.173.232'
    }

def getPersonScore(name = '', passwd = ''):
    params = {
        "__VIEWSTATE":"/wEPDwUKLTc4MzY5NjY5Nw9kFgICAw9kFgICFw8PFgIeBFRleHRlZGQYAQUeX19Db250cm9sc1JlcXVpcmVQb3N0QmFja0tleV9fFgUFCHJidG5EZXB0BQtyYnRuVGVhY2hlcgULcmJ0blRlYWNoZXIFC3JidG5TdHVkZW50BQtyYnRuU3R1ZGVudA1gZKqG2kk6EvC2qWLlC7HHx7io",
        "__EVENTVALIDATION":"/wEWCAKR+7KNDQKvruq2CALSxeCRDwLI0oXWBQKjndD0AwKViOj6BgL+jNCfDwKT+PmaCMChSLI3uc0DzHYNhu/SeMa6TuKo",
        "UserName":"",
        "Password":"",
        "userType":"rbtnStudent",
        "LoginButton":"登 陆"
        }

    params["UserName"] = name
    params["Password"] = passwd
    scores = {}

    params = urllib.urlencode(params)
    cookieJar = cookielib.CookieJar()
    opener = urllib2.build_opener(urllib2.HTTPCookieProcessor(cookieJar))
    urllib2.install_opener(opener)
    request = urllib2.Request('http://112.15.173.232/Login.aspx', params)
    response = opener.open(request)
    resHtml = response.read()
    if '帐号或密码错误!' in resHtml:
        scores['error'] = 'error'
        return scores

    infoPattern = re.compile(r'<td.*</td>')
    allInfo = infoPattern.findall(resHtml)
    for item in allInfo:
        itemPattern = re.compile(r'<font color="#333333">.*</font>')
        itemStr = str(item)
        itemInfo = itemPattern.findall(itemStr)
        single = itemInfo[0]
        single = single.replace('</td><td>', '').replace('</font>', '')
        result = single.split('<font color="#333333">')
        scores[result[2].strip()] = result[10].strip()
    return scores

if __name__ == '__main__':
    fetchStr = '-------------------------------'

    data = xlrd.open_workbook('names.xls')
    output = open('result.txt', 'a')
    table = data.sheets()[0]
    nrows = table.nrows
    for i in range(1, nrows):
        name = table.cell(i, 1).value
        id = table.cell(i, 0).value
        output.write(fetchStr + '\n')
        output.write (str(id) + '\n')
        score = getPersonScore(str(id), str(id))
        if 'error' in score:
            continue
        for s in score:
            output.write(s + '\t' + score[s] + '\n')
        output.write(fetchStr)
    output.close()
