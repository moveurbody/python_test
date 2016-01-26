#!/usr/bin/python
# coding=UTF-8

fo = open("C:\\Users\\Yuhsuan_chen\\Desktop\\test.py", 'r')
print "file name is: ", fo.name
fo.close()


def wf(path):
    new_file = open(path, "wb")
    new_file.write("hello! moto!\n")
    new_file.close()


def rf(path):
    new_file = open(path, "r")
    content = new_file.read(10)
    new_file.close()
    print content


def wr(path):
    new_file = open(path, "rb+")
    content = new_file.read(10)
    print "The 10 content is: ", content

    position = new_file.tell()
    print "File position is: ", position

    #shift the position, from position 10
    #new_file.seek(10, 0)
    #shift the position, from position 20
    new_file.seek(10, 1)
    position = new_file.tell()
    content = new_file.read(10)
    print "new position is: ", position
    print "file is :" , content
    new_file.close()


# wf("test.txt")
rf("test.txt")
print("----------------")
wr("test.txt")