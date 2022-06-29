from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
import os
import pandas as pd
import xlrd
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException

options = Options()
options.headless = True

#options.add_argument("--headless")
options.add_argument('--disable-gpu')
driver = webdriver.Chrome(options=options, executable_path="E:\Textbooks and notes\ME\TA\Data\chromedriver.exe")
folder = pd.read_excel("E:\Textbooks and notes\ME\TA\Data\secondary structure for MBNL1.xlsx") # can also index sheet by name or fetch all sheets
mylist = folder['Gene Name'].tolist()
for x in mylist:
    newfolder=x
    directory = newfolder
    parent_dir = "E:\Textbooks and notes\ME\TA\Data\Output\Energy Dot Plots\Dot Plots"
    path = os.path.join(parent_dir, directory)
    os.mkdir(path)

    params = {'behavior': 'allow', 'downloadPath': path}
    driver.execute_cdp_cmd('Page.setDownloadBehavior', params)

    df = pd.read_excel("E:\Textbooks and notes\ME\TA\Data\secondary structure for MBNL1.xlsx")
    mylist = df['Sequence'].tolist()
    for x in mylist:
        query=x
        driver.get("http://www.unafold.org/mfold/applications/rna-folding-form.php")   #input url

        df = pd.read_excel("E:\Textbooks and notes\ME\TA\Data\secondary structure for MBNL1.xlsx") # can also index sheet by name or fetch all sheets

    #print(mylist)


        driver.find_element_by_xpath("/html/body/main/div[2]/form/div[2]/textarea").send_keys(query)
        driver.find_element_by_xpath("/html/body/main/div[2]/form/div[3]/input[1]").click()
        driver.find_element_by_xpath("/html/body/main/div[2]/form/input[5]").click()
        delay = 5

        try:
            myElem = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.ID, 'IdOfMyElement')))
        except TimeoutException:
            print("hey")

        query=driver.current_url

        driver.get(query)

        #for downloading the pdf
        folder = driver.find_element_by_xpath("/html/body/main/div[2]/a[4]/em")  #input xpath
        folder.click()
        pdf= driver.current_url
        #print(pdf)
        import wget
        url= pdf
        wget.download(url, path)

        # for downloading Text
        driver.get(query)  # input url
        folder = driver.find_element_by_xpath("/html/body/main/div[2]/a[2]")
        folder.click()
        pdf = driver.current_url
        # print(pdf)
        url = pdf
        wget.download(url, path)

        # For downloading PostScript
        driver.get(query)  # input url
        folder = driver.find_element_by_xpath("/html/body/main/div[2]/a[3]/em")
        folder.click()

        # for downloading the png
        folder = driver.find_element_by_xpath("/html/body/main/div[2]/a[5]/em")  # input xpath
        folder.click()
        pdf = driver.current_url
        # print(pdf)
        src = driver.find_element_by_xpath("/html/body/main/div[2]/form/div/input").get_attribute("src")
        import wget

        url = src
        wget.download(url, path)

        # for downloading Sorted
        driver.get(query)  # input url
        sortedlink = driver.find_element_by_xpath("/html/body/main/div[2]/p[4]/a[1]")
        sortedlink.click()
        sl = driver.current_url
        # print(pdf)
        sortedfile = sl
        wget.download(sortedfile, path)

        # for downloading hnum
        driver.get(query)  # input url
        hnumlink = driver.find_element_by_xpath("/html/body/main/div[2]/p[4]/a[2]/i")
        hnumlink.click()
        hl = driver.current_url
        # print(pdf)
        hnum = hl
        wget.download(hnum, path)

        # for downloading pnum
        driver.get(query)  # input url
        pnumlink = driver.find_element_by_xpath("/html/body/main/div[2]/p[4]/a[3]/i")
        pnumlink.click()
        pl = driver.current_url
        # print(pdf)
        pnum = pl
        wget.download(pnum, path)

        # for downloading logfile
        driver.get(query)  # input url
        logfilelink = driver.find_element_by_xpath("/html/body/main/div[2]/p[4]/a[4]/i")
        logfilelink.click()
        lfl = driver.current_url
        # print(pdf)
        logfile = lfl
        wget.download(logfile, path)
        break