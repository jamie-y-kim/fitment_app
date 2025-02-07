import requests
from bs4 import BeautifulSoup
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
import time

# eBay 상품 URL
# url = "https://www.ebay.com/itm/133055316399?_trkparms=amclksrc%3DITM%26aid%3D777008%26algo%3DPERSONAL.TOPIC%26ao%3D1%26asc%3D20230823115209%26meid%3Dddff90da76d44343bb9bd2445639ea20%26pid%3D101800%26rk%3D1%26rkt%3D1%26itm%3D133055316399%26pmt%3D1%26noa%3D1%26pg%3D4375194%26algv%3DRecentlyViewedItemsV2SignedOut&_trksid=p4375194.c101800.m5481&_trkparms=parentrq%3Adea6a25f1940a67a10b37a2fffffabec%7Cpageci%3A60b106d4-e50b-11ef-b6d4-ba2a3d3bf849%7Ciid%3A1%7Cvlpname%3Avlp_homepage"
url = input("eBay URL: ")
file_name = input("Output file name: ")

# Selenium WebDriver 설정 (헤드리스 모드로 실행 가능)
driver = webdriver.Chrome()  # 크롬 브라우저 사용
driver.get(url)  # 시작 페이지

# 페이지 로딩 대기 시간
time.sleep(3)

# 데이터를 저장할 리스트
all_data = []

# 계속해서 'Next' 버튼을 눌러가며 데이터 크롤링
while True:
  # 현재 페이지에서 데이터를 추출
  soup = BeautifulSoup(driver.page_source, 'html.parser')
  rows = soup.find_all('tr', {'data-testid': 'ux-table-section-body-row'})
  
  for row in rows:
    all_data.append([cell.text.strip() for cell in row.find_all('td')])

  # 페이지네이션 버튼 클릭 (예시로 "Next" 버튼 클릭)
  try:
    next_button = driver.find_element(By.XPATH, '//button[contains(@aria-label, "Next compatibility table page")]')

    # aria-disabled 속성 확인 (비활성화된 버튼인지 체크)
    if next_button.get_attribute("aria-disabled") == "true":
      print("No more pages to load.")
      break

    # 버튼 클릭
    next_button.click()
    time.sleep(3)  # 페이지가 로드될 때까지 기다림
      
  except Exception as e:
    print("Error or no more pages:", e)
    break

driver.quit()

# 추출된 데이터를 DataFrame으로 변환
df = pd.DataFrame(all_data, columns=["Year", "Make", "Model", "Trim", "Engine", "Notes"])

# 엑셀로 저장
df.to_excel(f"{file_name}.xlsx", index=False)

print(f"{file_name}.xlsx 파일이 저장 완료:D")