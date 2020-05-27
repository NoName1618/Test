# -*- coding: utf-8 -*-
"""
Created on Fri May  1 13:18:42 2020

@author: SyunT
"""

import openpyxl as xl
from PIL import Image

inputFile = 'toaru4.xlsx'

def toColNumA1(n):
    s =""
    for i in range(10):
        if n <= 0:
            break
        s+= (chr((n-1)%26 + ord('A')))
        n = (n-1)//26
        
    return s[::-1]
        

wb1 = xl.load_workbook(filename = inputFile)
ws1 = wb1.worksheets[0]

def sizenize(y, x):
    for i in range(1,y+1):
        ws1.row_dimensions[i].height = 5
    for i in range(1,x + 1):
        ws1.column_dimensions[toColNumA1(i)].width = 1


pic = "toaru2.png"

im = Image.open(pic)
rgb_im = im.convert('RGB')

size = rgb_im.size
# (y , x)
sizenize(size[1], size[0])
# print(size[1],size[0])


def rgbToHex(rgb):
    return '%02x%02x%02x' % rgb

# sizeが大きすぎるとexcelでエラーが発生する
try:
    for x in range(1, size[0] + 1 ):
        for y in range(1, size[1] + 1 ):
            rgb = rgb_im.getpixel((x-1,y-1))
            rgbHex = 'FF' + str(rgbToHex(rgb)).upper()
            fill = xl.styles.PatternFill(patternType='solid',
                                         fgColor=rgbHex, bgColor=rgbHex)
            s = toColNumA1(x) + str(y)
            #print(s)
            ws1[s].fill = fill
except Exception as e:
    print(e)

wb1.save(inputFile)