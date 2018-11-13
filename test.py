#!/usr/bin/python
import sys
import json
import re
from openpyxl import load_workbook, Workbook

def i_to_xy(index):
    num = re.compile(r"[0-9]+")
    let = re.compile(r"[A-Z]+")
    letters = let.search(index).group()
    x = 0
    for i in range(len(letters)):
        x += 27**i*(ord(letters[::-1][i])-ord('A')+1)
    y = int(num.search(index).group())
    return x,y


def xy_to_i(x, y):
    pow = 0
    while 27**(pow) < x:
        pow += 1
    tmpx = x
    letters = ""
    for p in range(pow)[::-1]:
        i = 1
        while 27**p*i <= tmpx-27**p:
            i += 1
        tmpx -= 27**p*i
        letters += chr(64+i)
    numbers = str(y)
    return letters+numbers

zob = i_to_xy("AZZ91")
print("x ="+str(zob[0])+" y="+str(zob[1]))
print("");print("");print("")
x,y = i_to_xy("AZZ91")
zob = xy_to_i(x, y)
print("i ="+str(zob))

exit(0)

letters = ['A','B','C','B','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']
for a in letters:
    for b in letters:
        for c in letters:
            for x in range(1,10):
                for y in range(1,10):
                    for z in range(1,10):
                        test = a+b+c+str(x)+str(y)+str(z)
                        if ( test == xy_to_i(i_to_xy(test)[0],i_to_xy(test)[1])):
                            print(test+" ok")
                            lol = 1
                        else:
                            print(test+" FALSE")
