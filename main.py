#coding:utf-8

import urllib2
import re

baseUrl = 'http://image.baidu.com/search/index?tn=baiduimage&ct=201326592&lm=-1&cl=2&ie=gbk&word=%CD%BC%C6%AC&fr=ala&ala=1&alatpl=others&pos=0'

def downloadHtml(url):

    try:

        reponse = urllib2.urlopen(url).read()

    except urllib2.URLError as e:

        reponse = None
        print e.reason
        print e.code


    return reponse


if __name__ == '__main__':

    html = downloadHtml(baseUrl)

    patern = r'"objURL":"(.*?)",'

    match = re.findall(patern,html,re.S)

    for a in  match:
        print (a)


