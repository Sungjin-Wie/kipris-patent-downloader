from re import split
from selenium import webdriver
import time
import os, shutil
import openpyxl
import pandas as pd
import datetime
import chromedriver_autoinstaller
from PyQt5 import sip
from pathlib import Path
from tkinter import *
import tkinter.messagebox as msgbox
import tkinter.ttk as ttk 

class KIPRISDownloader():
    def __init__(self):

        self.URL = "http://www.kipris.or.kr/khome/main.jsp"
        self.interval = 1
        self.chrome_ver = chromedriver_autoinstaller.get_chrome_version().split('.')[0]
        self.chrome_options = webdriver.ChromeOptions()
        self.chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])

        try:
            self.driver = webdriver.Chrome(f'./{self.chrome_ver}/chromedriver.exe', options=self.chrome_options)   
        except:
            chromedriver_autoinstaller.install(True)
            self.driver = webdriver.Chrome(f'./{self.chrome_ver}/chromedriver.exe', options=self.chrome_options)

        self.driver.implicitly_wait(10)

        self.driver.get(url=self.URL)
        time.sleep(self.interval)

        self.isFirstTime = True
        self.searchResultText = ""
        self.keyWord = ""
        self.exceptionKeyWord = ""
        self.splitKeyWord = ""
        self.searchKeyWord = ""
        self.newFileName = "temp.xls"
        self.nowDateTime = ""
        self.wb = ""
        self.totalResults = 0
        self.page = 0
        self.pages = 1
        self.resultFile = ""
        self.progressLabelText = ""
        self.isDetailSearching = False

        self.root = Tk()
        self.root.title("KIPRIS Patent Downloader")
        self.root.geometry("270x270")
        self.root.resizable(height=False, width=False)

        self.searchInputLabel = Label(self.root, text="검색 키워드를 입력하세요.")
        self.searchInputLabel.grid(row=0, column=0, padx=30, pady=5)
        # self.searchInputLabel.pack()

        self.searchInputEntry = Entry(self.root, width=30)
        self.searchInputEntry.grid(row=1, column=0, columnspan=2, padx=30, pady=5)
        # self.searchInputEntry.pack()

        self.isDetailSearchingCheckButton = Checkbutton(self.root, text="결과 내 재검색", command=self.ToggleIsDetailSearching)

        self.searchExceptLabel = Label(self.root, text="제외할 키워드를 입력하세요.")
        self.searchExceptLabel.grid(row=2, column=0, padx=30, pady=5)
        # self.searchExceptLabel.pack()

        self.searchExceptEntry = Entry(self.root, width=30)
        self.searchExceptEntry.grid(row=3, column=0, columnspan=2, padx=30, pady=5)
        # self.searchExceptEntry.pack()

        self.searchButton = Button(self.root, command=self.OnSearchButtonClick, text="검색")
        self.searchButton.grid(row=4, column=0, padx=30, pady=5)

        self.searchResultLabel = Label(self.root, text=self.searchResultText)
        self.searchResultLabel.grid(row=5, column=0, padx=30, pady=5)

        self.downloadButton = Button(self.root, command=self.OnDownloadClick, text="다운로드")

        self.progressLabel = Label(self.root, text=self.progressLabelText)
        self.progressLabel.grid(row=7, column=0, padx=30, pady=5)

        self.progressBar = ttk.Progressbar(self.root, maximum=100, value=(self.page / self.pages) * 100)

        self.root.mainloop()

    def ToggleIsDetailSearching(self):
        self.isDetailSearching = not self.isDetailSearching


    def OnSearchButtonClick(self):
        self.keyWord = self.searchInputEntry.get()
        self.exceptionKeyWord = self.searchExceptEntry.get()
        if self.keyWord == "":
            msgbox.showinfo("알림", f"검색어를 입력해주세요.")
            return
            

        self.downloadButton.grid_forget()
        if not self.isFirstTime:
            self.searchResultText = ""
            self.searchResultLabel.config(text=self.searchResultText)
            self.driver.back()
            time.sleep(self.interval)
            self.driver.back()
            time.sleep(self.interval)
        
        self.isFirstTime = False

        self.driver.find_element_by_name('inputQueryText').clear()
    
        if self.exceptionKeyWord != "":
            self.searchKeyWord = self.keyWord + f"*!{self.exceptionKeyWord}"
        else:
            self.searchKeyWord = self.keyWord
    
        self.driver.find_element_by_name('inputQueryText').send_keys(self.searchKeyWord)
        time.sleep(self.interval)
        
        searchSelector = "#SearchPara > fieldset > div.float_left > span > button"
        self.driver.find_element_by_css_selector(searchSelector).click()

        isSearchingSelector = "#patentResultCountBoard"
        searchingText = self.driver.find_element_by_css_selector(isSearchingSelector).text
        while searchingText == "검색 중입니다.":
            time.sleep(self.interval)
            searchingText = self.driver.find_element_by_css_selector(isSearchingSelector).text


        time.sleep(self.interval)
        patentSelector = "#resultCountPatent > span"
        self.totalResults = self.driver.find_element_by_css_selector(patentSelector).text.replace(",", "")
        if self.totalResults == "0":
            msgbox.showinfo("결과", "검색 결과가 없습니다.")
        else:
            self.driver.find_element_by_css_selector(patentSelector).click()
            time.sleep(self.interval)
    
        while searchingText == "검색 중입니다.":
            time.sleep(self.interval)
            searchingText = self.driver.find_element_by_css_selector(isSearchingSelector).text

        totalResultsSelector = "#patentResultCountBoard > em.txt_bold"
        self.totalResults = int(self.driver.find_element_by_css_selector(totalResultsSelector).text.replace(",", ""))
        searchResultText = f"총 검색결과는 {self.totalResults}개입니다."
        self.searchResultLabel.config(text=searchResultText)
        msgbox.showinfo("결과", "검색이 완료되었습니다.")

        self.pages = self.totalResults // 90 + 1
        time.sleep(self.interval)

        self.downloadButton.grid(row=6, column=0, padx=30, pady=5)

        
    def OnDownloadClick(self):
        now = datetime.datetime.now()
        self.nowDatetime = now.strftime('%Y-%m-%d-%H-%M-%S')
        self.splitKeyWord = self.keyWord.split("*")[0]
        self.resultFile = f"result_{self.splitKeyWord}_{self.nowDatetime}.xlsx"
        
        msgbox.showinfo("알림", f"{self.splitKeyWord}에 관한 결과를 {self.resultFile}에 저장합니다.")
        
        # self.progressBar.pack()

        if not os.path.isfile(self.resultFile):
            self.wb = openpyxl.Workbook()
            self.wb.save(self.resultFile)
        else:
            self.wb = openpyxl.load_workbook(self.resultFile)
        for page in range(1, self.pages + 1):
            self.page = page
            if page == 1:
                optionSelector = "#opt28 > option:nth-child(3)"
                self.driver.find_element_by_css_selector(optionSelector).click()
                applySelector = "#pageListSetBtn"
                self.driver.find_element_by_css_selector(applySelector).click()
                time.sleep(self.interval)
                
            isSearchingSelector = "#patentResultCountBoard"
            searchingText = self.driver.find_element_by_css_selector(isSearchingSelector).text
            while searchingText == "검색 중입니다.":
                time.sleep(self.interval)
                searchingText = self.driver.find_element_by_css_selector(isSearchingSelector).text

            excelDownloadSelector = "#btnDownloadExcel"
            self.driver.find_element_by_css_selector(excelDownloadSelector).click()
            time.sleep(self.interval)
    
            filepath = str(os.path.join(Path.home(), "Downloads"))
            filename = max([filepath + '\\' + f for f in os.listdir(filepath)], key=os.path.getctime)
            if os.path.isfile(self.newFileName):
                os.remove(self.newFileName)       
            shutil.move(os.path.join(filepath, filename), self.newFileName)

            sheet = self.wb.active

            tempData = pd.read_excel(self.newFileName, sheet_name="검색결과")
            dataColumn = tempData.columns.tolist()
            dataList = tempData.values.tolist()

            if page == 1:
                sheet.append(dataColumn)
            for data in dataList:
                sheet.append(data)
        
            if page < self.pages:
                self.driver.execute_script(f"getSearchResultPage({page + 1})")

            self.progressLabel.config(text=f"진행도: {page} / {self.pages}")
            if page % 5 == 0:
                self.wb.save(self.resultFile)
                print("임시 저장을 진행합니다...")
    
        if os.path.isfile(self.newFileName):
            os.remove(self.newFileName)        
        self.wb.save(self.resultFile)
        # self.progressBar.pack_forget()
        msgbox.showinfo("알림", "저장이 완료되었습니다.")

app = KIPRISDownloader()