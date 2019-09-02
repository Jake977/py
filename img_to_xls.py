#!/usr/bin/env python
# coding: utf-8

import os
import cv2  #opencv lib
import xlsxwriter #xlsxwriter lib

from fnmatch import fnmatch

ext = '*.jpg'
results = []
testdir = "img" #dir with images
cell_height = 150 #cell height in xls
i = 1

workbook = xlsxwriter.Workbook('myworkbook1.xlsx')
worksheet1 = workbook.add_worksheet ('imgs')

for f in os.listdir(testdir):
    if fnmatch(f, ext):
        imgpath = (os.path.join(testdir, f))
        results.append(imgpath)

while i <= len(results):
    img = cv2.imread(results[i-1], cv2.IMREAD_UNCHANGED)
    height, width, depth = img.shape
    img_scale = round(cell_height/height, 3)
   
    worksheet1.set_row(i, cell_height)
    worksheet1.insert_image(i, 1, results[i-1], {'x_scale':img_scale, 'y_scale': img_scale})
    i +=1

workbook.close()
