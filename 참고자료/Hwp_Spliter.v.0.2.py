# -*- coding: utf-8 -*-
"""
HWP Api를 파이썬에서 쓰기 쉽도록 정리한 클래스이다. 

"""
'''잘라 넣기 기능을 구현한다. 기본적 아이디어는 간단하다. 
1. 구획을 이동한다.
2. 거기서 부터 Top까지 이동한다.
3. 복사해서 클립보드에 넣는다.
4. 클립보드에서 Txt를 불러온다.
5. Numbering과 함께 파일명을 만든다. (저장폴더도 생성한다.)
6. 파일을 부분 저장한다.
7. 선택영역을 지운다.
8. 이동할 수 없을 때까지 1~7를 반복한다.
9. 파일을 저장하지 않고 닫는다.
'''
from hwpapi import HwpApi as HwpApi

def hwp_spliter(hwp_path):
    import pyperclip as pc
    
    hwp_folder = os.path.dirname(hwp_path)    
    hwp = HwpApi()
    hwp.hwpOpen(hwp_path)
    
    switch = True
    n = 0
    while switch:
        n += 1
        try:
            hwp.Goto(2)
            hwp.MoveSelTopLevelBegin()
            print("OK")
            hwp.Copy() #복사하기 기능을 구현해야 한다.
            text_lines = pc.paste()
            first_line = text_lines.splitlines()[0]
            print(first_line)
            number = "{:03.0f}".format(n)
            save_file = number + '. ' + first_line + '.hwp'
            print(save_file)
            file_addr = os.path.join(hwp_folder, save_file)
            hwp.hwpSaveAs_S(file_addr)
            hwp.Delete()
        except:
            switch = False
    hwp.FileClose()