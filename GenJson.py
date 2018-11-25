# -*- coding:utf-8 -*-
import xml.etree.cElementTree as et
import os
import re
import sys
import string
import xml.dom.minidom as minidom

OutFilePath_Json = r"E:\Unity2017Work\GameFrameWork\AssetsData"

def GetFileList(dir, fileList):
    for s in os.listdir(dir):
        if os.path.isfile(s):
            fileList.append(s)
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

    str_json = "{\"entry\":[\n{"
    for f_entry in f_class:
        tag = 1   
        for entry in f_entry:
            for text in entry.text:
                if "#" in text:
                    tag = 0
        
        if tag == 0:
            #忽略带#标签的条目
            continue

        for fields in f_entry:
            str_json = str_json +"\"" + fields.tag + "\":\"" + fields.text+"\","
        
        str_json = str_json+"},\n"
        
    str_json = str_json + "]}"
    out_file = OutFilePath_Json + "\\" + className + ".json"
    f = open(out_file,"w+", encoding = "UTF-8")
    f.writelines(str_json)
    f.close()

if ( __name__ == "__main__"):
    print("Program name", sys.argv[0])
    fileNames = GetFileList(sys.path[0]+"\\",[])
    for fName in fileNames:
        if ".xml" in fName:
            GenJson(fName)


    res = input('press any key to continue.')
    