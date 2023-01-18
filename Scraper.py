from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas
import xlsxwriter
import time
import os
from random import randint
import warnings 
import chromedriver_autoinstaller
from selenium.webdriver.chrome.service import Service

warnings.filterwarnings("ignore", category=UserWarning)

start_time = time.time()
print("Started...")
df = pandas.read_excel ('Input.xlsx')

Iterator = df['Iterator'].tolist()
MainURL = df["URL"].tolist()
Settings = df["Settings"].tolist()
ElementNameList = df["ElementName"].tolist()
ElementPathList = df["Xpath"].tolist()
ElementTypeList = df["ElementType"].tolist()

Iterator = list(filter(lambda x: str(x) != 'nan', Iterator))          #to remove Excel null cells
ElementNameList = list(filter(lambda x: str(x) != 'nan', ElementNameList))
ElementPathList = list(filter(lambda x: str(x) != 'nan', ElementPathList))
ElementTypeList = list(filter(lambda x: str(x) != 'nan', ElementTypeList))


FullScreenshot_Option = Settings[0]
LoadImages_Option = Settings[1]
MaxRandomPause_Option = Settings[2]
UsingIterator_Option = Settings[3]


CreationTime = time.strftime("%m-%d-%Y_%H-%M-%S", time.localtime())
FileName = "Scraper " + CreationTime 
Main_dir = "./" + FileName
FullScreenshots_dir = "./" + FileName + "/FullScreenshots"
os.makedirs(Main_dir)
workbook = xlsxwriter.Workbook(Main_dir + "/" + FileName + '.xlsx')
worksheet = workbook.add_worksheet("Results")

#Excel Headers
Headers = []
for elementname in ElementNameList:
    Headers.append(elementname)
if FullScreenshot_Option == "On":
    os.makedirs(FullScreenshots_dir)

HeaderCount = 0
for X in Headers:
    worksheet.write(0 , HeaderCount, X)
    HeaderCount = HeaderCount + 1

HeadersList = [
    "Mozilla/5.0 (Linux; U; Android 4.4.2; en-us; SCH-I535 Build/KOT49H) AppleWebKit/534.30 (KHTML, like Gecko) Version/4.0 Mobile Safari/534.30",
    "Mozilla/5.0 (iPhone; CPU iPhone OS 10_3_1 like Mac OS X) AppleWebKit/603.1.30 (KHTML, like Gecko) Version/10.0 Mobile/14E304 Safari/602.1",
    "Mozilla/5.0 (Linux; Android 7.0; SM-A310F Build/NRD90M) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.91 Mobile Safari/537.36 OPR/42.7.2246.114996",
    "Mozilla/5.0 (Android 7.0; Mobile; rv:54.0) Gecko/54.0 Firefox/54.0"
]

mobileEmulation = {"deviceMetrics": {"width": 2000, "height": 5000, "pixelRatio": 3.0}, "userAgent": HeadersList[randint(0,3)]}
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('ignore-certificate-errors')
chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
chrome_options.add_experimental_option('mobileEmulation', mobileEmulation)
chrome_options.headless = True

if LoadImages_Option == "Off":
    chrome_options.add_argument('--blink-settings=imagesEnabled=false') 

chromedriver_autoinstaller.install(True)
chromedriver_majorversion = chromedriver_autoinstaller.get_chrome_version().split(".")[0]
service = Service(executable_path=chromedriver_majorversion+"/chromedriver")
driver = webdriver.Chrome(service=service, options=chrome_options)

results = {}

def Scrape(URL):
    try:
        driver.get(URL)
        time.sleep(randint(1,MaxRandomPause_Option))
        driver.save_screenshot(FullScreenshots_dir + "/" + time.strftime("%m-%d-%Y_%H-%M-%S", time.localtime()) +'.png')
        for elementname in ElementNameList:
            path = ElementPathList[ElementNameList.index(elementname)]

            if elementname not in results:
                results[elementname] = []
            content = driver.find_elements(By.XPATH, path)

            if ElementTypeList[ElementNameList.index(elementname)] == "text":
                for element in content:
                    results[elementname].append(element.text)
            elif ElementTypeList[ElementNameList.index(elementname)] == "href":
                for element in content:
                    results[elementname].append(element.get_attribute("href"))
    except:
        print("error!")

def Writer(input_dic):
    for key in input_dic:
        row = 1
        for line in input_dic[key]:
            worksheet.write(row , Headers.index(key), line)
            row += 1

#Reserved characters in the parameter
Reserved_dict = {"%":"%25","␣":"%20","!":"%21","#":"%23","$":"%24",
                "&":"%26","'":"%27","(":"%28",")":"%29","*":"%2A","+":"%2B",",":"%2C","/":"%2F",
                ":":"%3A",";":"%3B","=":"%3D","?":"%3F","@":"%40","[":"%5B","]":"%5D"}

if UsingIterator_Option == "On":
    for i in Iterator:
        try:
            Mod_i = i
            for t,v in Reserved_dict.items():
                Mod_i = Mod_i.replace(t,v)
            Url = MainURL[0].replace("[iterator]", Mod_i)
        except:
            print("Error, Check the Input Excel file!")
        Scrape(Url)
elif UsingIterator_Option == "Off":
    Url = MainURL[0]
    Scrape(Url)

Writer(results)
workbook.close()
driver.quit()
end_time = time.time()
print(round((end_time - start_time),2), "seconds")