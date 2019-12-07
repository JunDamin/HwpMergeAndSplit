# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

#module 불러오기
import os
import win32com.client as wc
import xml.etree.cElementTree as et
import time

#hwp파일 열어 hml 파일로 저장하기

def hwpSaveAs(HwpAuto, os, file_address, save_ext='.hml'):
    '''설명을 달아야지'''
    format_set = {'.hwp':'HWP', '.hml':'HWPML2X'}
    file_set = os.path.splitext(file_address)
    file_path = os.path.split(file_address)[0]
    #파일 열기
    HwpFileOpen=HwpAuto.CreateAction("FileOpen")
    HwpFileOpenPara = HwpFileOpen.CreateSet()
    HwpFileOpen.GetDefault(HwpFileOpenPara)
    HwpFileOpenPara.SetItem("FileName", file_address)
    HwpFileOpenPara.SetItem("Format", format_set[file_set[1]])
    HwpFileOpen.Execute(HwpFileOpenPara)
    #파일 저장하기
    save_file_address = os.path.join(file_path, file_set[0]+save_ext)
    HwpFileSaveAs=HwpAuto.CreateAction('FileSaveAs')
    HwpFileSaveAsPara=HwpFileSaveAs.CreateSet()
    HwpFileSaveAs.GetDefault(HwpFileSaveAsPara)
    HwpFileSaveAsPara.SetItem("FileName", save_file_address)
    HwpFileSaveAsPara.SetItem("Format", format_set[save_ext])
    HwpFileSaveAs.Execute(HwpFileSaveAsPara)
    #파일 닫기
    HwpAuto.Run("FileClose")
    return(save_file_address)

def hmlParse(hml_file_address, et):
    '''파일을 파싱해서 tree를 반환 '''
    if len(os.path.dirname(hml_file_address))==0:
        hml_file_address = os.path.abspath(hml_file_address)
    hml = open(hml_file_address, encoding='utf-8')
    hml_read = hml.read()
    hml.close()
    tree = et.fromstring(hml_read)
    return(tree)

def folderGenerator(file_address):
    '''폴더 만들기 있면 패스'''
    if len(os.path.dirname(file_address))==0:
        folder_name = os.path.splitext(file_address)[0]
        folder_address = os.getcwd()
    else:
        folder_name = "(#분리본) " + os.path.splitext(os.path.basename(file_address))[0]
        folder_address = os.path.dirname(file_address)
    new_folder_address = os.path.join(folder_address, folder_name)
    print(new_folder_address)
    if os.path.isdir(new_folder_address):
        print("폴더가 존재합니다")
        return(new_folder_address)
    else:
        os.mkdir(new_folder_address)
        return(new_folder_address)

def hmlSeedMaker(hml_root):
    for index in range(len(hml_root[1])-1, -1, -1):
        hml_root[1].remove(hml_root[1][index])
    return(hml_root)
    
def hmlMake(section, seed):
    seed[1].append(section)
    return(seed)

def sectionList(hml_root):
    sectionlist = hml_root[1].getchildren()
    return(sectionlist)

def hmlGenerator(sectionlist, seed, folder_address):
    number = 0
    hml_list = []
    char_list = ["    ", "ㅁ", "ㅇ", "○", "□", "◎", "▣", "◈"]
    replace_char = {"/":"-", "\\":"-", ":":";", "*":"+", "?":"+", "\"":"'", "<":"(", ">":")", "|":"!"}
    for section in sectionlist:
        number += 1
        name_lenth = []
        text = "" #저장하기 위한 텍스트
        for char in section.findall(".//CHAR"): #문자열 뽑아내기
            if char.text != None:
                text += char.text
        for i in char_list:
            index = text.find(i)
            if index >0:
                name_lenth.append(text.find(i))
        name_lenth.append(40)
        #print(name_lenth)
        lenth = max(13, min(name_lenth))
        #print(lenth)
        output_name = "("+format(number, '03') + ") " + text[0:lenth] +".hml"
        for i in replace_char:
            output_name = output_name.replace(i, replace_char[i])        
        element = hmlMake(section, seed)
        output = et.tostring(element, encoding='utf-8')
        #파일 저장하기
        output_address = os.path.join(folder_address, output_name)
        #print(output_address)
        hml_output = open(output_address, 'w', encoding='utf-8')
        hml_decoded = output.decode('utf-8')
        #print(type(hml_decoded))
        hml_decoded = '<?xml version="1.0" encoding="utf-8" standalone="no" ?>' + hml_decoded #정보 추가 입력
        hml_output.write(hml_decoded)
        hml_output.close()
        print(os.path.basename(output_name))
        hml_list.append(output_name)
        seed = hmlSeedMaker(seed)
    return(hml_list)


def hml_spliter(HwpAuto, os, file_address, et):
    hml_file_address = hwpSaveAs(HwpAuto, os, file_address, save_ext='.hml')
    root = hmlParse(hml_file_address, et)
    folder_address = folderGenerator(hml_file_address)
    sectionlist = sectionList(root)
    seed = hmlSeedMaker(root)
    hml_list = hmlGenerator(sectionlist, seed, folder_address)
    os.remove(hml_file_address)
    return(hml_list)


# 작업파일 위치
path = os.getcwd()
os.chdir(path)
file_list = os.listdir(path)
HwpAuto = wc.gencache.EnsureDispatch('HWPFrame.HwpObject.1')

#작업할 파일 정의
hwp_list = []
for file in file_list:
    file_ext = os.path.splitext(file)[1]
    if file_ext == '.hwp':
        hwp_list.append(file)

        
#작업
log_txt = open("log.txt", 'w', encoding = 'utf-8')

for hwp in hwp_list:
    hwp_address = os.path.abspath(hwp)
    log = hml_spliter(HwpAuto, os, hwp_address, et)
    log_txt.write("\n ".join(log) + "\n")
    

number = 0

for item in os.walk(path):
    #print(item)
    folder_address = item[0]
    file_list = item[2]
    save_ext = '.hwp'
    for file in file_list:
        if os.path.splitext(file)[1]==".hml":
            hml_file = os.path.join(folder_address, file)
            save_file_address = os.path.join(folder_address, os.path.splitext(file)[0]+save_ext)
            if os.path.isfile(save_file_address):
                os.remove(save_file_address)
            #print(hml_file)
            hwp_file = hwpSaveAs(HwpAuto, os, hml_file, save_ext)
            number += 1
            log_txt.write(format(number, '03') + " " + os.path.basename(hwp_file) +"\n")
            print(os.path.basename(hwp_file))
            time.sleep(1)
            os.remove( os.path.join(folder_address, file))
        
HwpAuto.Quit()
log_txt.close()

import time
print('\n\n', number, "건으로 분리 작업 완료되었습니다")
time.sleep(3)