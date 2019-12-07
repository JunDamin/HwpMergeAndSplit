import hmlspliter as hs
from hwpapi import HwpApi
import os

def convert_file(file_addr, new_format=".hml"):
    hwpapi = HwpApi()
    hwpapi.hwpOpen(file_addr)
    save_addr = hwpapi.hwpSaveAs(file_addr, new_format)
    hwpapi.hwpFileClose()
    hwpapi.hwpQuit()
    return save_addr

def convert_files(file_addr_list, new_format=".hml"):
    hwpapi = HwpApi()
    save_addr_list = []
    for file_addr in file_addr_list:
        hwpapi.hwpOpen(file_addr)
        save_addr = hwpapi.hwpSaveAs(file_addr, new_format)
        save_addr_list.append(save_addr)
        hwpapi.hwpFileClose()
    hwpapi.hwpQuit()
    return save_addr_list

def make_abspath(file_addr):
    if os.path.dirname(file_addr):
        return file_addr
    else:
        return os.path.abspath(file_addr)

def process(file_addr):
    abs_addr = make_abspath(file_addr)
    hml_addr = convert_file(abs_addr, new_format=".hml")
    hml_list = hs.split_hml(hml_addr)
    convert_files(hml_list, new_format=".hwp")
    for hml in hml_list:
        os.remove(hml)