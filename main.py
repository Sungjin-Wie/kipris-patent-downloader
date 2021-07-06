from selenium import webdriver
import time
import os, shutil
import openpyxl
import pandas as pd
import datetime
import chromedriver_autoinstaller

chrome_ver = chromedriver_autoinstaller.get_chrome_version().split('.')[0]

profile = {'savefile.default_directory': os.getcwd(), 'download.default_directory': os.getcwd()}

chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])
chrome_options.add_experimental_option('prefs', profile)
try:
    driver = webdriver.Chrome(f'./{chrome_ver}/chromedriver.exe', options=chrome_options)   
except:
    chromedriver_autoinstaller.install(True)
    driver = webdriver.Chrome(f'./{chrome_ver}/chromedriver.exe', options=chrome_options)

driver.implicitly_wait(10)


URL = "http://www.kipris.or.kr/khome/main.jsp"
interval = 1

keyWord = input("키워드를 입력해주세요: ")
splitKeyWord = keyWord.split("*")[0]
now = datetime.datetime.now()
nowDatetime = now.strftime('%Y-%m-%d-%H-%M-%S')

resultFile = f"result_{splitKeyWord}_{nowDatetime}.xlsx"
print(f"{keyWord}에 관한 내용을 {resultFile}에 저장합니다...")

if not os.path.isfile(resultFile):
    wb = openpyxl.Workbook()
    wb.save(resultFile)
else:
    wb = openpyxl.load_workbook(resultFile)


driver.get(url=URL)
time.sleep(interval)

driver.find_element_by_name('inputQueryText').send_keys(keyWord)
time.sleep(interval)

searchSelector = "#SearchPara > fieldset > div.float_left > span > button"
driver.find_element_by_css_selector(searchSelector).click()

isSearchingSelector = "#patentResultCountBoard"
searchingText = driver.find_element_by_css_selector(isSearchingSelector).text
while searchingText == "검색 중입니다.":
    time.sleep(interval)
    searchingText = driver.find_element_by_css_selector(isSearchingSelector).text


time.sleep(interval)
patentSelector = "#resultCountPatent > span"
totalResults = driver.find_element_by_css_selector(patentSelector).text.replace(",", "")
if totalResults == "0":
    print("검색 결과가 없습니다.")
else:
    driver.find_element_by_css_selector(patentSelector).click()
    time.sleep(interval)
    
    while searchingText == "검색 중입니다.":
        time.sleep(interval)
        searchingText = driver.find_element_by_css_selector(isSearchingSelector).text

    totalResultsSelector = "#patentResultCountBoard > em.txt_bold"
    totalResults = int(driver.find_element_by_css_selector(totalResultsSelector).text.replace(",", ""))
    print(f"총 검색결과는 {totalResults}개입니다.")
    pages = totalResults // 90 + 1
    time.sleep(interval)
    
    newFileName = "temp.xls"
    
    for page in range(1, pages + 1):
        

        if page == 1:
            optionSelector = "#opt28 > option:nth-child(3)"
            driver.find_element_by_css_selector(optionSelector).click()
            applySelector = "#pageListSetBtn"
            driver.find_element_by_css_selector(applySelector).click()
            time.sleep(interval)
        
        searchingText = driver.find_element_by_css_selector(isSearchingSelector).text
        while searchingText == "검색 중입니다.":
            time.sleep(interval)
            searchingText = driver.find_element_by_css_selector(isSearchingSelector).text

        excelDownloadSelector = "#btnDownloadExcel"
        driver.find_element_by_css_selector(excelDownloadSelector).click()
        time.sleep(interval)
    
        filepath = os.getcwd()
        filename = max([filepath + '\\' + f for f in os.listdir(filepath)], key=os.path.getctime)
        if os.path.isfile(newFileName):
            os.remove(newFileName)       
        shutil.move(os.path.join(filepath, filename), newFileName)

        sheet = wb.active

        tempData = pd.read_excel(newFileName, sheet_name="검색결과")
        dataColumn = tempData.columns.tolist()
        dataList = tempData.values.tolist()

        if page == 1:
            sheet.append(dataColumn)
        for data in dataList:
            sheet.append(data)
        
        if page < pages:
            driver.execute_script(f"getSearchResultPage({page + 1})")
        
        print(f"진행도: {page} / {pages}")
        if page % 5 == 0:
            wb.save(resultFile)
            print("임시 저장을 진행합니다...")
    
    if os.path.isfile(newFileName):
        os.remove(newFileName)        
    wb.save(resultFile)
    print("저장이 완료되었습니다.")

os.system("pause")