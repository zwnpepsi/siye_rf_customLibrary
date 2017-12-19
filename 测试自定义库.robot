*** Settings ***
Library    customLibrary/DoExcel.py    


*** Test Cases ***
测试使用自定义库
    Read Excel
    Write Excel    1    1   '小简老师棒棒哒！'