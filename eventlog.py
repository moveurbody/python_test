#!/usr/bin/python
# coding=UTF-8

import os
import time


def write(__str__):
    current_exec_path = os.getcwd()
    localtime = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
    fo = open(current_exec_path + '\\Eventlog.txt', 'ab+')
    fo.write(localtime + " " + __str__ + "\n")
    fo.close()
