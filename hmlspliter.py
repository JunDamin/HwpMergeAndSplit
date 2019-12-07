#module 불러오기
import os
from lxml import etree as et

def parse_hml(hml_addr):
    '''파일을 파싱해서 tree를 반환 '''
    if len(os.path.dirname(hml_addr))==0:
        hml_addr = os.path.abspath(hml_addr)
    tree = et.parse(hml_addr)
    root = tree.getroot()
    return(root)

def make_folder(file_address):
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

def make_hmltemplate(hml_root):
    for index in range(len(hml_root[1])-1, -1, -1):
        hml_root[1].remove(hml_root[1][index])
    return(hml_root)
    
def make_hml(section, hmltemplate):
    hmltemplate[1].append(section)
    return(hmltemplate)

def make_sectionlist(hml_root):
    sectionlist = hml_root[1].getchildren()
    return(sectionlist)

def generate_hml(sectionlist, hmltemplate, folder_address):
    hml_list = []
    number = 0
    for section in sectionlist:
        number += 1
        output_name = make_outputname(section, number)       #이름 추출해서 넣기
        output_address = os.path.join(folder_address, output_name)
        write_hml(hmltemplate, section, output_address)
        print(os.path.basename(output_name))
        hml_list.append(output_address)
        hmltemplate = make_hmltemplate(hmltemplate)
    return(hml_list)

def make_outputname(section, number):
    char_list = ["    ", "ㅁ", "ㅇ", "○", "□", "◎", "▣", "◈"]
    replace_char = {"/":"-", "\\":"-", ":":";", "*":"+", "?":"+", "\"":"'", "<":"(", ">":")", "|":"!"}
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
    return output_name

def write_hml(hmltemplate, section, output_address):
    output = make_hml(section, hmltemplate)
    output = output.getroottree()
    output.write(output_address, xml_declaration=True)

def split_hml(hml_addr):
    hml_addr
    root = parse_hml(hml_addr)
    folder_address = make_folder(hml_addr)
    sectionlist = make_sectionlist(root)
    hmltemplate = make_hmltemplate(root)
    hml_list = generate_hml(sectionlist, hmltemplate, folder_address)
    os.remove(hml_addr)
    return(hml_list)

