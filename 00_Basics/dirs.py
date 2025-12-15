#-*-coding:UTF-8

"""
Description : Dir 에 대한 기본적인 정보 가져오기
"""
import os
import sys
import time
import datetime
import pandas as pd
from python_calamine import CalamineWorkbook



print(f"\n")
print(f"현재 호출한 dir 위치   / os.getcwd()             : {os.getcwd()}")
print(f"현재 실행중인 .py 이름 / __file__                : {__file__}")
print(f".py 의 경로만 / os.path.dirname                  : {os.path.dirname(__file__)}")
print(f".py 의 파일명만 / os.path.basename               : {os.path.basename(__file__)}")
print(f".py 의 절대경로+이름 / os.path.abspath(__file__) : {os.path.abspath(__file__)}")
print(f"\n")
print(f"호출위치의 상대 경로+이름 / abspath.replace(cwd) : {os.path.abspath(__file__).replace(os.getcwd(),'')}")
print(f"호출위치의 상대 이름만 / abspath.replace.replace : {os.path.abspath(__file__).replace(os.getcwd(),'').replace(os.path.basename(__file__),'')}")
print(f"\n")


print(f"\n")
# print(f" : {}")
# print(f" : {}")
# print(f" : {}")
# print(f" : {}")
# print(f" : {}")


sys.exit()


### 0. 준비 ################################################################################
# 프로그램 시작 메시지를 출력합니다.
# sys.argv 의 갯수를 확인합니다.
if len(sys.argv) != 2:
    print(f"사용방법 : {sys.argv[0]} [엑셀파일들이 저장된 폴더 이름]")
    print(f"사용예시 : {sys.argv[0]}  xlsx_directory_name ")
    sys.exit()


### 1. 입력 파라미터 확인 ##############################################################
# 작업 시작 메시지를 출력합니다.
print(f"Starting {sys.argv[0]}\n")

# 시작 시점의 시간을 기록합니다.
start_time = time.time()

# 현재 파이썬 파일이 저장된 폴더의 절대 경로를 구합니다.
pwd = os.path.dirname(os.path.abspath(__file__)) 

# 파일들이 저장된 폴더 이름을 입력 받은 파라미터로 받아옵니다.
source_dir = sys.argv[1]

# 엑셀 파일들이 저장된 폴더의 절대 경로를 구합니다.
source_abs_dir = pwd + "\\" + source_dir

# 엑셀 파일들의 dictionary 를 생성합니다.
file_list = os.listdir(source_abs_dir)

### 2. 엑셀 파일 분석 ######################################################################
# 파일의 종류를 기록할 딕셔너리를 만듭니다.
FILE_TYPES = {}

# 시트가 여러 개인 엑셀 파일의 갯수를 기록할 딕셔너리를 만듭니다.
MULTI_SHEET_FILES = {}
multi_sheet_file_count = 0

# 헤더들을 저장할 딕셔너리를 만듭니다.
HEADERS = {}

# file_list에 저장된 파일 이름을 한 번에 하나씩 불러옵니다.
# 양식이 몇 종류인지 분석합니다.
for file_name in file_list:
    # 1. 파일 확장자를 확인합니다. ####################################################################
    file_ext = file_name.split('.')[-1].lower()
    #print(f"File Name : {file_name}, File Ext : {file_ext}")
    
    # 딕셔너리에서 파일의 동일한 확장자가 이미 헤더가 삽입되어 있는지 확인합니다.
    if file_ext in FILE_TYPES:
        # 이미 삽입되어 있다면 값을 1개 증가시킵니다.
        FILE_TYPES[file_ext] += 1
    else:
        # 없다면 새로 삽입하고 값을 1로 설정합니다.
        FILE_TYPES[file_ext] = 1
    
    # 2. 간혹 xlsx 파일명이 아닌 파일이 섞여있을 수 있습니다. 이걸 걸러냅니다. ##########################
    if ".xlsx" not in file_name:
        continue
    
    # 3. 엑셀 파일이 맞다면, 파일을 읽어옵니다. ###########################################
    file_abs_name = source_abs_dir + "\\" + file_name
    print(f"Reading : {file_name}")
    file = pd.ExcelFile(file_abs_name, engine='calamine')
    
    # 엑셀파일의 시트 이름들을 구합니다.
    sheet_names = file.sheet_names
    # print(f"'{file_name}' 파일의 시트 이름 목록:")
    # print(sheet_names)
    
    # 엑셀 파일이지만 시트가 1개 이상인지 확인합니다.
    if len(sheet_names) > 1 :
        # 시트가 2개 이상이라면, 딕셔너리에 시트가 여러 개인 엑셀 파일로 기록합니다.
        MULTI_SHEET_FILES[file_name] = len(sheet_names)
        multi_sheet_file_count += 1
    
    # 시트가 2개 이상이라면, 모든 시트를 불러와 DataFrame 으로 변환합니다.
    #for sheet_name in sheet_names :
    #    df = pd.read_excel(file_name, sheet_name=sheet_name, engine='calamine')
    #    print(df)
    # dataframe_list = pd.read_excel(file_name, sheet_name=None)  # sheet_name=None 으로 하면 모든 시트를 불러옴
    # print(dataframe_list)  # 딕셔너리 형태로 반환됨, key 는 시트 이름, value 는 DataFrame
    # for sheet_name, df in dataframe_list.items():
    #     print(f"Sheet name: {sheet_name}")
    #     print(df) # 각 시트의 DataFrame 출력
    # if len(sheet_names) > 1 :
    #     print(f"'{file_name}' 파일은 시트가 {len(sheet_names)}개 입니다.")  
    # else :
    #     print(f"'{file_name}' 파일은 시트가 1개 입니다.")       
    # #   
    # print(dataframe_list[sheet_names[0]])  # 첫 번째 시트의 DataFrame 출력
    
            
    # 엑셀 파일의 첫번째 시트를 불러와 DataFrame 으로 변환합니다.
    df = pd.read_excel(file_abs_name, sheet_name=file.sheet_names[0], engine='calamine')
    #print(df)
        
    # 엑셀 파일의 첫 번째 열, 그러니까 헤더만 불러와 스트링으로 변환합니다.
    #header = str(df.iloc[0, :].tolist())    # 이건 데이터의 첫 행
    header = str(df.columns.tolist())        # 이건 컬럼명   

    # 딕셔너리에서 엑셀파일의 같은 헤더가 삽입되어 있는지 확인합니다.
    if header in HEADERS:
        # 이미 삽입되어 있다면 값을 1개 증가시킵니다.
        HEADERS[header] += 1
    else:
        # 없다면 새로 삽입하고 값을 1로 설정합니다.
        HEADERS[header] = 1


### 3. 결과 보고서 작성 ######################################################################
# 결과물 레포트를 작성하기 위한 스트링을 생성합니다.
REPORT = ""

# 파일 갯수를 레포트에 작성합니다.
REPORT += "### 폴더 내 총 파일 갯수 : " + str(len(file_list)) + " ###\n\n"

# 레포트에 추출한 파일 종류를 기록합니다.
REPORT += "### 폴더 내 파일 종류 ###\n"
for key in FILE_TYPES:
    REPORT += "File Type : " + key + " / "
    REPORT += "Count : " + str(FILE_TYPES[key]) + "\n"    

# 레포트에 시트가 여러 개인 엑셀 파일들을 기록합니다.
REPORT += "\n### 시트가 여러 개인 엑셀 파일수 : " + str(multi_sheet_file_count) +" ###\n"
for key in MULTI_SHEET_FILES:
    REPORT += "File Name : " + key + " / "
    REPORT += "Sheet Count : " + str(MULTI_SHEET_FILES[key]) + "\n"

# 레포트에 추출한 헤더정보를 기록합니다.
REPORT += "\n### 폴더 내 엑셀 파일들의 헤더 종류 ###\n"
for key in HEADERS:
    REPORT += "Header : " + key + " / "
    REPORT += "Count : " + str(HEADERS[key]) + "\n"

# # 레포트를 화면에 출력합니다.
# print(REPORT)


# 레포트 파일에 레포트를 저장합니다.
#datetime.datetime(2024, 7, 30, 16, 43, 34, 530921)
filename_result_output = "result_REPORT-" + datetime.datetime.now().strftime('%y%m%d_%H%M%S') + ".txt"
report_file = open(pwd + "\\" + filename_result_output, 'w')
report_file.write(REPORT)
report_file.close()
print(f"Wrote Report File : {filename_result_output}")


# 작업 종료 메시지를 출력합니다.
print("\nProcess Done.")

### 4. 작업 종료 ################################################################################
# 작업에 총 몇 초가 걸렸는지 출력합니다.
end_time = time.time()
print("The Job Took " + str(end_time - start_time) + " seconds.")
