# -*- coding:utf-8 -*-
import xml.etree.cElementTree as et
import os
import re
import sys
import string
import xml.dom.minidom as minidom

OutFilePath_Json = r"E:\Unity2017Work\Lemon\Lemon_assetsdata\Templates"

def GetFileList(dirPath):
    fileList=[]
    for s in os.listdir(dirPath):
        if '.xml' in s:
            fileList.append(dirPath+s)

    return fileList
    
def GenJson(filePath):
    if not os.path.isfile(filePath):
        return 
    print("[GenJson] begin process file : " + filePath)
    className = filePath.split('\\')[-1].split('.')[0]

    f_tree = et.parse(filePath)
    f_root = f_tree.getroot()
    f_data = f_root.find('data')
    f_class = f_data.find(className)
    isAddSplit = False
    style={}
    str_json = "{\"entry\":[\n"
    for f_entry in f_class:
        tag = 1
        styleTag = 0
        for entry in f_entry:
            if not entry.text is None and "#类型" in entry.text:
                styleTag = 1

            if not entry.text is None and "#" in entry.text:
                tag = 0
                break

        if tag == 0:
            #忽略带#标签的条目
            if styleTag == 1:
                for entry in f_entry:
                    if "#" in entry.text:
                        continue
                    style[entry.tag] = entry.text
            continue

        if isAddSplit:
            str_json=str_json+",\n"

        isAddSplit = True
        str_json=str_json+"{"
        couter = len(f_entry)
        for fields in f_entry:
            if fields.text is None:
                continue
            if style[fields.tag] == "string":
                str_json = str_json +"\"" + fields.tag + "\":\"" + fields.text+"\""
            else:
                str_json = str_json +"\"" + fields.tag + "\":" + fields.text
			
            if couter > 1:
                couter = couter-1
                str_json = str_json+","
        
        str_json = str_json+"}"

    str_json = str_json + "\n]}"
    out_file = OutFilePath_Json + "\\" + className + ".json"
    f = open(out_file,"w+", encoding = "UTF-8")
    f.writelines(str_json)
    f.close()

if ( __name__ == "__main__"):
    print("Program path ", sys.path[0])
    fileNames = GetFileList(sys.path[0]+"\\")
    print("files count: "+ str(len(fileNames)))
    for fName in fileNames:
        if ".xml" in fName:
            GenJson(fName)

    res = input('press any key to continue.')
    