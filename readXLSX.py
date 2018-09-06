#coding:utf-8


import os
import xlrd
import xlwt
from xlutils.copy import copy
import re
import webbrowser

from urllib import request
from urllib import error
from bs4 import BeautifulSoup
from selenium import webdriver
import time
from selenium.webdriver import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import requests
import  cv2
import openpyxl


class readHtml(object):

    def __init__(self):
        pass

    def loadHtml(self,url):

        if url == '':
            return ''

        driver = webdriver.Chrome()
        driver.get(url)

        # 让程序休眠,等待网页彻底加载完成
        time.sleep(sleepTime)

        page = driver.page_source

        try:
            driver.close()
        except:
            pass
        return page

    def loadHtmlByRe(self,url):

        try:
            # if url == '':
            #
            #     return  ''

            Request = request.urlopen(url)
            html = Request.read()
            html = html.decode('utf-8')
            return html

        except error.URLError as e:

            try:
                print(e.code)
                return 'ErrorCode' + str(e.code)
            except:
                return 'ErrorCode' + str('链接不完整')

        except error.HTTPError as e:

            print(e.reason)

    def matchStr(self,html,url):

        videoUrl = ''
        soup = BeautifulSoup(html, 'html5lib')

        if url.startswith('http://juxian.juyun.tv'):

            tab = soup.find('title')
            for u in tab:

                if u == '活动已关闭':

                    videoUrl = 'isOver'
                    return videoUrl

                else:
                    pass

            tab = soup.find_all('video')


            for u in tab:
                videoUrl = u.get('src')
                print('内部链接:' + videoUrl)

        elif url.startswith(r'http://720yun.com/'):

            tab = soup.find_all('iframe')

            for u in tab:

                videoUrl = u.get('src')

        elif url.startswith(r'http://zt.gxtv.cn/'):

            tab = soup.find_all('video')
            for u in tab:
                videoUrl = u.get('src')
                print('内部链接:' + videoUrl)
        else:

            tab = soup.find_all('source')

            for u in tab:
                videoUrl = u.get('src')
                print('内部链接:' + videoUrl)

        return videoUrl


    def getUrl(self,url,params):

        headers = {'Connection':'close'}

        r = requests.get(url = url,params = params)
        print(r.text)

    def gerWebDriver(self):

        return self.driver



class Excel(object):


    def __init__(self,excelPath = '/Users/liangTV/Desktop/check.xlsx',isGoON = '1'):

        # 网页异常的原因
        self.reason = ''
        # 待监测的Excel路径
        self.excelPath = excelPath
        self.readExcel(self.excelPath)
        # 异常的条数
        self.exceptCount = 0
        self.sheetRowCount = 0
        self.isNormal = True

        global newFile
        # 待保存的表单
        self.cacheWriteSheet = None
        # 读取到的单个Excel表单数据
        # self.cacheReadSheet = None
        # 读取到的Excel整个表单的数据
        # self.cacheReadExcel = None

        if isGoON is '2':

            oldPath = input('请输入上次储存的路径:')
            if str(oldPath) is '':

                oldPath = os.path.join(os.path.expanduser('~'),'desktop')
                oldPath = oldPath + '/test.xls'

            cacheFile = xlrd.open_workbook(oldPath)
            newFile = copy(cacheFile)
            self.cacheWriteSheet = newFile.get_sheet(0)

        else:

            newFile = xlwt.Workbook()
            self.cacheWriteSheet = newFile.add_sheet(U'1')

        self.readhtml = readHtml()

    # 读取表格信息
    def readExcel(self,path = ''):

        if path is '':

            path = input('请输入有效的路径:')
            self.readExcel(path)

        else:

            self.cacheReadExcel = xlrd.open_workbook(path)
            self.cacheReadSheet = self.cacheReadExcel.sheets()[0]

    # 输出文件路径
    def printPath(self):

        print(self.excelPath)

    #输出表格信息
    def printData(self):

        table = self.cacheReadExcel.sheets()[1]

        nrows = table.nrows

        for dataRow in range(nrows):

            if dataRow == 0:

                continue

            print(self.cacheReadSheet.row_values[dataRow][2])
    #检查网页内链接
    def checkVideoUrl(self,videoUrl):

        webDriver = webdriver.Chrome()
        webDriver.set_page_load_timeout(5)

        checkInfo = []

        try:
            webDriver.get(videoUrl)
            print('没有超时')
            checkInfo.append('视频异常')
            checkInfo.append('视频链接不正常')

            self.isNormal = False
            self.exceptCount = self.exceptCount + 1

            try:
                webDriver.close()
            except:
                pass

        except:

            print('已超时')

            checkInfo.append('正常')

            try:
                webDriver.close()
            except:
                pass

        return checkInfo

    # 检查表格中的URL
    def checkUrl(self,beginRow = 0,stopRow = ''):

        nrows = self.cacheReadSheet.nrows

        self.sheetRowCount = nrows

        if str(stopRow) is '':
            pass

        elif nrows >= int(stopRow):

            nrows = int(stopRow)

        else:
            pass

        for row in range(beginRow,nrows):

            print('当前行数:' + str(row))

            if row == 0:
                continue

            dataRow = self.cacheReadSheet.row_values(int(row))

            url = dataRow[1]

            print(url)

            # 360 度视频
            if url.startswith(r'http://720yun.com/'):

                html = self.readhtml.loadHtmlByRe(url)
                if html.startswith('ErrorCode'):

                    print('异常')
                    dataRow.append('网页异常')
                    dataRow.append(html)

                    self.isNormal = False
                    self.exceptCount = self.exceptCount + 1
                else:
                    print('正常 ')
                    dataRow.append('正常')
            else:
                # 普通网页
                html = self.readhtml.loadHtmlByRe(url)

                if html.startswith('ErrorCode'):

                    dataRow.append('网页异常')
                    dataRow.append(html)

                    self.isNormal = False
                    self.exceptCount = self.exceptCount + 1

                else:

                    html = self.readhtml.loadHtml(url)
                    videoUrl = self.readhtml.matchStr(html, url);

                    # print('视频链接:' + videoUrl)

                    if videoUrl is '':

                        dataRow.append('视频异常')
                        dataRow.append('获取不到视频链接')

                        self.isNormal = False
                        self.exceptCount = self.exceptCount + 1

                    elif videoUrl is 'isOver':

                        dataRow.append('网页异常')
                        dataRow.append('活动结束')

                        self.isNormal = False
                        self.exceptCount = self.exceptCount + 1

                    else:

                        checkInfo = self.checkVideoUrl(videoUrl)
                        for oneCol in checkInfo:
                            dataRow.append(oneCol)

            self.writeData(dataRow, row)

            continue

        # 添加总异常行数
        dic = []
        dic.append('异常总数:')
        dic.append(self.exceptCount)
        self.isNormal = False
        self.writeData(dic, self.sheetRowCount)

        # 保存文件
        global newFile
        self.saveFile(newFile)


    # 写入每一行的数据
    def writeData(self,dataRow,row):

        i = 0

        global newSheet

        for dataOne in dataRow:

            if self.isNormal == True:

                self.cacheWriteSheet.write(row, i, dataOne, self.setExcleNormalStyle())
            else:
                self.cacheWriteSheet.write(row, i, dataOne, self.setExcleExceptStyle())

            i = i + 1
        self.isNormal = True

    # 保存文件
    def saveFile(self,file):

        path = input('请输入新文件存储的路径(默认保存在桌面test.xls文件):')

        if str(path) is '':

            path = os.path.join(os.path.expanduser('~'),'desktop')

            path = path + '/test.xls'

        print(path)

        try:
            file.save(path)
            print('save successfull!')

        except:

            print('保存异常!请检查')

    # 设置异常单元格格式
    def setExcleExceptStyle(self):

        style = xlwt.XFStyle()
        font = xlwt.Font()  # 为样式创建字体
        font.name = 'Times New Roman'
        font.bold = False
        font.color_index = 2
        font.height = 200
        style.font = font

        return style

    # 设置普通单元格格式
    def setExcleNormalStyle(self):

        style = xlwt.XFStyle()
        font = xlwt.Font()  # 为样式创建字体
        font.name = 'Times New Roman'
        font.bold = False
        font.color_index = 4
        font.height = 200

        style.font = font

        return style

if __name__ == '__main__':

    # 用来保存结果的文件
    newFile = None
    excel = None

    path = input('输入表格的路径:')
    isGoOn = input('是为是继上次检查:1 = 不是，2 = 是:')

    startRow = input('开始检查的行数:')
    stopRow = input('停止检查的行数:')

    sleepTimeString = input('请输入加载网页的等待时间:')
    if sleepTimeString is '':
        sleepTimeString = '0'

    sleepTime = float(sleepTimeString)

    try:

        excel = Excel(isGoON=str(isGoOn))

        if str(startRow) is '':

            startRow = 0

        excel.checkUrl(int(startRow),stopRow)
    except:

        print('卧槽,报错了')

    finally:

        print('先保存已经验证过的')
        excel.saveFile(newFile)


