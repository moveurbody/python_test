#!/usr/bin/python
# coding=UTF-8
import os

print ("start")
for root, dirs, files in os.walk("C:\Users\Yuhsuan_chen\PycharmProjects\untitled"):
    print root
    for f in files:
        print os.path.join(root, f)
