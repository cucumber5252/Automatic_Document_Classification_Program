import pandas as pd
import os
import natsort
import re
import shutil
import time

# 엑셀 파일을 읽어오는 함수
def read_excel_files():
   # 각각의 파일을 구분하여 엑셀 파일을 로드
   s = pd.read_excel('Z:/2023 센터/2023 공문/수발신문서대장/2023 수신문서등록대장.xlsx', sheet_name='수신', usecols="L")
   b = pd.read_excel('Z:/2023 센터/2023 공문/수발신문서대장/2023 발신문서등록대장.xlsx', sheet_name='발신', usecols="M")
   n = pd.read_excel('Z:/2023 센터/2023 공문/수발신문서대장/2023 내부문서등록대장(일반문서).xlsx', sheet_name='내부', usecols="I")
   us = pd.read_excel('Z:/2023 센터/2023 공문/수발신문서대장/2023 업무연락수신대장.xlsx', sheet_name='업무수신', usecols="J")
   ub = pd.read_excel('Z:/2023 센터/2023 공문/수발신문서대장/2023 업무연락발신대장.xlsx', sheet_name='업무발신', usecols="K")
   return s, b, n, us, ub



# 데이터를 합치고 정리하기 위한 함수
def create_concat_file_list(s, b, n, us, ub):
   # 필요한 부분만 추출하여 합치기
   ##########concat_list만 맞게 기재하시면 파일 종류 상관 없이 정상 작동합니다##########
   ##########스캔 파일가 수신 문서 201부터 203호, 발신 문서 123호부터 125호, 내부 89호부터 92호, 업무연락전수신 3호부터 5호라면##########
   ##########concat_list = [s.loc[201:203], b.loc[123:125], n.loc[90:93], us.loc[3:5]]로 기재##########
   ##########n은 .loc 입력 시 +1 해줘야 합니다(내부 85-1 문서 때문)##########
   concat_list = [n.loc[109+1:111+1], us.loc[48:60], ub.loc[14:19]]
   print("-----concat_list입니다-----")
   print(concat_list)
   df = pd.concat(concat_list, ignore_index=True).dropna().reset_index(drop=True)

   return df



# 파일 이름 변경하는 함수
def rename_files(df, folder_path, sorted_file_list):
   # 파일 이름을 새로운 이름으로 변경 및 변경된 이름 리스트 반환
   print("-----df입니다-----")
   print(df)
   print("-----sorted_file_list입니다-----")
   print(sorted_file_list)

   new_file_list = []  # 변경된 파일명을 저장할 리스트
   for index, row in df.iterrows():
      old_named_file = os.path.join(folder_path, sorted_file_list[index])
      new_named_file = os.path.join(folder_path, row['문서 파일명'] + '.pdf')
      os.rename(old_named_file, new_named_file)  
      new_file_list.append(row['문서 파일명'] + '.pdf')  # 새 파일명 저장
      print(f"{old_named_file}   ===>   {new_named_file} \n")
   return new_file_list


# 파일 이동하는 함수
def move_files(folder_path, sorted_file_list):
   # 각각의 경로 설정
   s_path = "Z:/2023 센터/2023 공문/전자문서/수신 전자문서"
   b_path = "Z:/2023 센터/2023 공문/전자문서/발신 전자문서"
   n_path = "Z:/2023 센터/2023 공문/전자문서/내부 전자문서"
   us_path = "Z:/2023 센터/2023 공문/전자문서/업무연락전 수신 전자문서"
   ub_path = "Z:/2023 센터/2023 공문/전자문서/업무연락전 발신 전자문서"

   # 정규표현식을 사용하여 패턴 정의
   s_pattern = r'^\d{8}_\d{3}_[\s\S]*\.pdf$'
   b_pattern = r'^경기수통광주2023-\d{3}_[\s\S]*\.pdf$'
   n_pattern = r'내부2023-\d{3}_[\s\S]*\.pdf$'
   us_pattern = r'^\d{3}_\d{8}_[\s\S]*\.pdf$'
   ub_pattern = r'\d{8}_0\d{1,2}_[\s\S]*\.pdf$'


   # 파일 이름에 따라 적절한 폴더로 이동
   for file_name in sorted_file_list:
      if re.match(ub_pattern, file_name):
         source = os.path.join(folder_path, file_name)
         shutil.move(source, ub_path)
         print(f"{file_name}을 업무연락전 발신 공문 폴더로 이동하였습니다.\n")

      else:
         if re.match(s_pattern, file_name):
            source = os.path.join(folder_path, file_name)
            shutil.move(source, s_path)
            print(f"{file_name}을 수신 공문 폴더로 이동하였습니다.\n")

         elif re.match(b_pattern, file_name):
            source = os.path.join(folder_path, file_name)
            shutil.move(source, b_path)
            print(f"{file_name}을 발신 공문 폴더로 이동하였습니다.\n")

         elif re.match(n_pattern, file_name):
            source = os.path.join(folder_path, file_name)
            shutil.move(source, n_path)
            print(f"{file_name}을 내부 문서 폴더로 이동하였습니다.\n")

         elif re.match(us_pattern, file_name):
            source = os.path.join(folder_path, file_name)
            shutil.move(source, us_path)
            print(f"{file_name}을 업무연락전 수신 공문 폴더로 이동하였습니다.\n")

         else:
            print(f"{file_name}은 어떤 조건에도 해당하지 않아 이동되지 않았습니다.\n")

# 메인 함수 실행
if __name__ == '__main__':
   s, b, n, us, ub = read_excel_files()
   df = create_concat_file_list(s, b, n, us, ub)
   
   folder_path = "C:/Users/GUEST-02/Desktop/공문 정리 자동화 프로그램/스캔 파일 제목 변경"
   file_list = os.listdir(folder_path)
   sorted_file_list = natsort.natsorted(file_list)
   
   renamed_file_list = rename_files(df, folder_path, sorted_file_list)
   move_files(folder_path, renamed_file_list)