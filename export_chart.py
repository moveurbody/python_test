#!/usr/bin/python
# coding=UTF-8

import logging
import os
import time

from win32com.client import Dispatch

# Setup log format,level and path
logging.basicConfig(level=logging.DEBUG,
                    filename='debug.log',
                    format='%(asctime)s %(levelname)s %(message)s',
                    datefmt='%y-%m-%d %H:%M:%S')
# Time
localtime = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
localtime2 = time.strftime('%Y%m%d_%H%M', time.localtime(time.time()))
localtime3 = time.strftime('%Y%m%d', time.localtime(time.time() - 24 * 60 * 60))
# Local path
current_exec_path = os.getcwd()

try:
    xlsApp = Dispatch("Excel.Application")
    xlsWB = xlsApp.Workbooks.Open(r'C:\Users\Yuhsuan_chen\PycharmProjects\untitled\Detail_20160324_1148.xlsx')
    xlsSheet = xlsWB.Sheets("Data")
    i = 0
    for chart in xlsSheet.ChartObjects():
        print chart.Name
        chart.Chart.Export("C:\Users\Yuhsuan_chen\PycharmProjects\untitled\chart" + str(i) + ".png")
        i = i+1
except Exception, e:
    print "ex"
    print e
    logging.error(e)
finally:
    xlsWB.Close()
