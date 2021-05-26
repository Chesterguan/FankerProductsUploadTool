#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed May 26 15:58:04 2021

@author: ziyuanguan

Usage :  Re-origanize the products folders
         Fetch product properties
         Pack Folders into ZIP waiting for upload.
"""

import pandas as pd 
import xlwt
import os
import subprocess


productColumns = ['商品名称', '商品图片','编号','材质','类型','颜色','价格','详情介绍', '分类名称', '标签' ]

outputBase = '../../Processed/'
outputImageFolder = outputBase + 'pics/'
outputDetailsFolder = outputBase + 'details/'

sizeLimit = 1.2 # GB

failed_proc = []


def getSize(srcFolder):
    size = int(subprocess.check_output(['du','-sh', srcFolder]).split()[0].decode('utf-8').split('K')[0])
    sizeInG = round(size / 1024 / 1024 , 6)
    return sizeInG
    
    
def zipFiles(srcFolder,name):
    ## example zip -r hello.zip products/*
    command = 'zip -r %s.zip %s/*' %(name , srcFolder)
    proc = subprocess.Popen(command , shell = True)
    if(not proc.returncode == 0):
        failed_proc.append(srcFolder)


def copyFolderToDes(srcPath):
    ## example : cp -r ../8301 /Processed/
    command  = 'cp -r %s %s'%(srcPath , outputImageFolder)
    proc = subprocess.Popen(command , shell = True)
    if(not proc.returncode == 0):
        failed_proc.append(srcPath)
        
        

## get product details from '进货单'
## copy image and detail folders
## generate list.xlsx
## split size if too large
## packed as ZIP
def run(srcFolder):
    _temp = {}
    for folder in os.list(srcFolder):
        ## GET multiple folder and each folder refers to one product
        if os.path.isdir(folder) :
            _temp['name'] = folder
            
            
    