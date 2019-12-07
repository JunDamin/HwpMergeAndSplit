# -*- coding: utf-8 -*-
"""
HWP Api를 사용 하위 폴더에 있는 hwp항목들을 각 폴더명 파일로 저장하는 툴

"""
import os
import sys
from hwpapi import HwpApi

#경로별로 작업
def hwpMerger(file_list):
    hwpapi = HwpApi()
    #파일 목록 만들기(역순으로 만들어야 끼어 넣기가 순서대로)
    FileList=file_list
    FileList.reverse() 
    print("\n".join(FileList), '\n')
    HwpFileList=[file for file in FileList if file[-3:]=="hwp"]
    if len(HwpFileList)==0:
        hwpapi.hwpQuit()
    print("\"모두 허용\"을 선택하시면 병합이 시작됩니다.")      
    #병합하기
    for i in HwpFileList:
        ongoing = hwpapi.HwpInsertFileFuction(i)
        print(ongoing)
    hwpapi.HWPAUTO.Run("Delete") #첫페이지(비어 있음) 삭제
    HwpPageNumber=hwpapi.HWPAUTO.CreateAction("PageNumPos") #페이지 번호 넣는 액션 생성
    HwpPageNumberPara=HwpPageNumber.CreateSet() #파라미터 세트 생성
    HwpPageNumberPara.SetItem("DrawPos", "BottomCenter") #파라미터 지정
    HwpPageNumber.GetDefault(HwpPageNumberPara) #파라미터 연결
    HwpPageNumber.Execute(HwpPageNumberPara) #실행
    print("작업이 완료되었습니다.")