import os
import time
import pandas as pd
import zipfile
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# 엑셀 파일을 읽고 데이터 프레임으로 반환하는 함수
def read_excels():
   # 엑셀 파일을 읽어 데이터 프레임으로 변환
   s = pd.read_excel('Z:/2023 센터/2023 공문/수발신문서대장/2023 수신문서등록대장.xlsx', sheet_name='수신', usecols="J")
   b = pd.read_excel('Z:/2023 센터/2023 공문/수발신문서대장/2023 발신문서등록대장.xlsx', sheet_name='발신', usecols="K")
   

   # 데이터 프레임을 합쳐서 반환
   ##########수신, 발신 문서는 concat_list만 맞게 기재하시면 정상 작동합니다##########
   ##########스캔 파일가 수신 문서 201부터 203호, 발신 문서 123호부터 125호라면##########
   ##########concat_list = [s.loc[201:203], b.loc[123;125]]로 기재##########
   ##########n은 .loc 입력 시 +1 해줘야 합니다(내부 85-1 문서 때문)##########
   concat_list = [s.loc[202:221], b.loc[95:95], b.loc[106:129]]
   df = pd.concat(concat_list, ignore_index=True).dropna().reset_index(drop=True)
   print("-------df입니다------")
   print(df)
   print('\n')


   ##########오류(중간에 스캔이 덜 된 페이지가 있기 때문)가 발생한다면 page_start 부분만 수정하시면 정상 작동합니다##########
   #오류 발생 시 수동 df 설정
   print("======수동 설정 시작합니다======")

   ##########여기에 각 문서 첫 페이지만 list 형식으로 입력해주세요##########
   page_start = []
   page_start_to_find_error=[]


   for each_page in page_start:
      page_start_to_find_error.append(each_page - page_start[0] +1)

   print("-------page_start_to_find_error입니다------")
   print(page_start_to_find_error)
   print('\n')

   page_count = []

   for i in range(len(page_start) - 1):
      interval = page_start[i+1] - page_start[i]
      page_count.append(interval)

   df = pd.DataFrame(page_count, columns=['페이지 수'])
   
   print("-------df 수정되었습니다------")
   print(df)
   print('\n')


   ##########내부, 업무연락전은 엑셀에 페이지 수를 기재하지 않기에 수기로 넣어주셔야 합니다##########
   ##########add_page_start에 원하는 문서의 첫 페이지를 순서대로 list 형태로 기재해주세요##########
   ##########ex.54페이지까지 수/발신 문서이고 55-57 내부 문서, 58-60 내부 문서, 61-62 업무연락전 수신, 63 업무연락전 발신 의 경우##########
   ##########add_page_start = [55,58,61,63,64]로 기재 (맨 마지막 페이지+1은 항상 list 끝에 넣어주셔야 합니다)##########

   # add_page_start = []

   # add_page_count = []

   # for i in range(len(add_page_start) - 1):
   #    interval = add_page_start[i+1] - add_page_start[i]
   #    add_page_count.append(interval)


   # for idx, value in enumerate(add_page_count, start=len(df)):
   #    df.loc[idx] = [value]
   # print(df)
   # print('\n')


   return df



# 데이터 프레임을처리하여 새로운 수정된 데이터 프레임을 반환하는 함수
def prepare_data(df):
   new_df = []
   m = 1
   new_df.append(m)

   # 페이지 수를 기반으로 파일 범위를 계산
   for a, b in df.iterrows(): 
      m += b['페이지 수']
      new_df.append(int(m-1))
      new_df.append(int(m))

   new_df.pop()
   print(new_df)
   return new_df



# 파일 이름 변경하는 함수
def rename_file(folder_path):
   to_split_change = os.listdir(folder_path)
   old_named_file = os.path.join(folder_path, to_split_change[0])
   new_named_file = os.path.join(folder_path, 'to_split.pdf')
   os.rename(old_named_file, new_named_file)
   print(f"{old_named_file}   ===>   {new_named_file} \n")



# Chrome 드라이버 설정 및 웹 사이트를 여는 함수
def setup_driver():
   chrome_options = Options()
   chrome_options.add_experimental_option("detach", True)
   driver = webdriver.Chrome(options=chrome_options)
   driver.get("https://www.ilovepdf.com/ko/split_pdf#split,range")
   return driver



# PDF 파일을 업로드하는 함수
def upload_pdf(driver, file_path):
   time.sleep(1)
   file_input = driver.find_element(By.XPATH, "//input[@type='file']")
   file_input.send_keys(os.path.expanduser(file_path))



# 범위를 적용하는 함수
def apply_ranges(driver, n_ranges, new_df):
   btn_add_range = driver.find_element(By.XPATH, '//*[@id="tab-content-range"]/div[2]/div[2]/div/button')
   # 범위를 추가하고 입력하도록 수행
   for a in range(n_ranges-1):
      time.sleep(0.5)
      driver.execute_script("arguments[0].scrollIntoView();", btn_add_range)
      btn_add_range.click()

   # 각 범위에 대하여 시작 및 끝 페이지를 입력
   for i in range(n_ranges):
      btn_range_start = driver.find_element(By.XPATH, f'//*[@id="range-option-{i+1}"]/div[2]/div[2]/input')
      btn_range_end = driver.find_element(By.XPATH, f'//*[@id="range-option-{i+1}"]/div[2]/div[3]/input')
      driver.execute_script("arguments[0].value = ''", btn_range_start)
      driver.execute_script("arguments[0].value = ''", btn_range_end)
      btn_range_start.send_keys(new_df[i*2])
      btn_range_end.send_keys(new_df[i*2 + 1])

   time.sleep(1)

   btn_split = driver.find_element(By.XPATH, '//*[@id="processTask"]')
   # 스플릿을 시작하고 결과를 다운로드 받음
   btn_split.click()


def wait_for_download(folder_path):
    time.sleep(3)  # 다운로드를 시작하기 위해 기다립니다.
    past_files = os.listdir(folder_path)
    downloading_files = [
        f for f in past_files if f.endswith(".crdownload") or f.endswith(".part")
    ]

    while downloading_files:
        time.sleep(1)
        current_files = os.listdir(folder_path)
        downloading_files = [
            f for f in current_files if f.endswith(".crdownload") or f.endswith(".part")
        ]


def download_and_extract_pdf(driver):
   folder_path = "C:/Users/GUEST-02/Downloads"
   before_files = set(os.listdir(folder_path))

   WebDriverWait(driver, 50).until(
      EC.element_to_be_clickable((By.XPATH, '//*[@id="pickfiles"]'))
   )
   btn_download = driver.find_element(By.XPATH, '//*[@id="pickfiles"]')
   btn_download.click()

   wait_for_download(folder_path)  # 다운로드가 완료될 때까지 충분히 기다려야 합니다.

   # 다운로드 받은 압축 파일을 압축 해제하여 지정한 폴더에 저장
   after_files = set(os.listdir(folder_path))
   downloaded_file = list(after_files - before_files)[0]
   print(downloaded_file)
   destination_folder = "C:/Users/GUEST-02/Desktop/공문 정리 자동화 프로그램/스캔 파일 제목 변경"

   with zipfile.ZipFile(os.path.join(folder_path, downloaded_file), 'r') as zip_ref:
      zip_ref.extractall(destination_folder)





df = read_excels()  # 엑셀 파일을 읽고 데이터 프레임 생성
new_df = prepare_data(df)  # 데이터 프레임 전처리
n_ranges = len(new_df) // 2  # 범위의 개수 계산
folder_path = "C:/Users/GUEST-02/Desktop/공문 정리 자동화 프로그램/파일 분할"
rename_file(folder_path)  # 파일 이름 변경
driver = setup_driver()  # 드라이버 설정 및 웹 사이트 접속
upload_pdf(driver, "C:/Users/GUEST-02/Desktop/공문 정리 자동화 프로그램/파일 분할/to_split.pdf")  # PDF 파일 업로드
apply_ranges(driver, n_ranges, new_df)  # 페이지 범위 적용
download_and_extract_pdf(driver)  # PDF 파일 다운로드 및 압축 해제