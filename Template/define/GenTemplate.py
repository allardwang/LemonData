# -*- coding:utf-8 -*-
import xml.etree.cElementTree as et
import os
import re
import sys
import string
import xml.dom.minidom as minidom

#脚本说明：
#
#该脚本会遍历所在目录及其子目录下所有的xml文件
# 并将其生成.cs和.xml
#
#运行环境：python 3.x  (python 2.x可以使用 GenTemplate_py2x.py)
#
#Author：王飞飞
#Time：2018-11
#

OutFilePath_cs = r"E:\Unity2017Work\Lemon\Lemon_client\Assets\DataTables"
OutFilePath_Xml = r"E:\Unity2017Work\Lemon\Lemon_data\Template"
TablesName_cs = "Tables.cs"
Namespace_cs = "Tables"

def GetFileList(dir, fileList):
    if os.path.isfile(dir):
        fileList.append(dir)
    elif os.path.isdir(dir):
        for s in os.listdir(dir):
            newDir = os.path.join(dir,s)
            GetFileList(newDir, fileList)  
    return fileList

def XMLDefineHandle(filePath):
    if not os.path.isfile(filePath):
        return 
    print("[XMLDefineHandle] begin process file : "+filePath)
    
    tree = et.parse(filePath)
    root = tree.getroot()
    _tableName = root.tag
    _rootAttr = root.attrib
    key = "id"
    if "key" in _rootAttr:
        key = _rootAttr["key"]
    
    filedNameAndTypes = {}

    for child in root:
        _filedName = child.tag
        _filedType  = child.attrib["type"]
        _filedComment = child.attrib["comment"]
        #如果定义文件增添新字段，请在此处处理
        filedNameAndTypes[_filedName] = {"type":_filedType,"comment":_filedComment}
    
    GenerateClass(_tableName, key, filedNameAndTypes)
    GenerateXml(_tableName, filedNameAndTypes)
    
def GenerateClass(className, classkey, classValue):
    #生成CS文件
    fileNmae = OutFilePath_cs + "\\"+TablesName_cs
    fp = open(fileNmae,"a")
    fp.writelines("\nnamespace "+ Namespace_cs+"\n")
    fp.writelines("{"+"\n")
    fp.writelines("\tpublic class "+className+"\n")
    fp.writelines("\t{\n")
    for kv in classValue:
        t = classValue[kv]["type"]
        if t.lower() == "int32":
            t="int"
        fp.writelines("\t\tpublic "+ t +" "+kv+";"+"\n")
    fp.writelines("\t}\n")
    fp.writelines("}\n")
    fp.close()
    print("[XMLDefineHandle] Generate "+TablesName_cs+" success.")

def GenerateXml(className, classValue):
    #生成配置表数据
    root_name = et.Element("root")
    root_name.attrib["xmlns:xsi"]="http://www.w3.org/2001/XMLSchema-instance"
    data_name = et.SubElement(root_name, "data")
    class_name = et.SubElement(data_name, className)
    entry_name_1 = et.SubElement(class_name, "entry")
    field_1_1 = et.SubElement(entry_name_1,"REMARK0")
    field_1_1.text = "#" + className
    for kv in classValue:
        key = et.SubElement(entry_name_1, ""+kv)
        key.text = (classValue[kv]["comment"])
    field_1_2 = et.SubElement(entry_name_1, "REMARK1")
    field_1_2.text="-"
    field_1_3 = et.SubElement(entry_name_1, "REMARK2")
    field_1_3.text="-"

    entry_name_2 = et.SubElement(class_name, "entry")
    field_2_1 = et.SubElement(entry_name_2, "REMARK0")
    field_2_1.text = u"#类型"
    for kv in classValue:
        key = et.SubElement(entry_name_2, ""+kv)
        key.text = classValue[kv]["type"]
    
    #原始表格数据添加（如果原本配置表已经有数据，将会保留配置数据）
    out_file = OutFilePath_Xml+"\\"+className+".xml"
    if os.path.isfile(out_file):
        print(out_file)
        f_tree = et.parse(out_file)
        f_root = f_tree.getroot()
        f_data = f_root.find('data')
        f_class = f_data.find(className)
 
        for f_entry in f_class:
            tag = 1
            for entry in f_entry:
                for text in entry.text:
                    if "#" in text:
                        tag = 0
            
            if tag == 0: #忽略带#标签的条目
                continue

            newData = et.SubElement(class_name, "entry")
            for kv in classValue:
                newData_field = et.SubElement(newData, kv)
                entry_field = f_entry.find(kv)
                if not entry_field is None:
                    newData_field.text = entry_field.text

    res = out_xml(className, root_name)
    if res:
        print("Generate XML Template success.")
    else:
        print("Generate XML Template failed.")

def out_xml(className, root):
    #生成XML文件
    rough_string = et.tostring(root, encoding = "UTF-8")
    out_file = OutFilePath_Xml+"\\"+className+".xml"
    reared_content = minidom.parseString(rough_string)
    with open(out_file, 'w+', encoding = "UTF-8") as fs:
        reared_content.writexml(fs, indent='', addindent = "\t", newl = "\n", encoding = "UTF-8")
    return True


if ( __name__ == "__main__"):
    fileNames = GetFileList(sys.path[0]+"\\",[])
    if os.path.isfile(OutFilePath_cs + "\\"+TablesName_cs):
        os.remove(OutFilePath_cs + "\\"+TablesName_cs)
    
    fp = open(OutFilePath_cs + "\\"+TablesName_cs,"w")
    fp.writelines("using System;\n")
    fp.writelines("using System.Collections;\n")
    fp.close()

    for fName in fileNames:
        if ".xml" in fName:
            XMLDefineHandle(fName)

    res = input("press any key to continue.")