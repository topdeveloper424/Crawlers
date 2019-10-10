import pyodbc
import pymssql
import decimal
import time
import sys
import xml.etree.ElementTree as ET
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options 
from selenium.webdriver.support.ui import Select
from fake_useragent import UserAgent
from bs4 import BeautifulSoup as bfs
from selenium.webdriver.common.action_chains import ActionChains
from random import randrange

def getDriver():
    #chrome_options = Options()
    chrome_options = webdriver.ChromeOptions()
    ua = UserAgent()
    userAgent = ua.random
    chrome_options.add_experimental_option("excludeSwitches",["ignore-certificate-errors"])
    #chrome_options.add_argument('headless')

    chrome_options.add_argument(f'user-agent={userAgent}')
    prefs = {'profile.managed_default_content_settings.images':2}
    chrome_options.add_experimental_option("prefs", prefs)

    prefs = {"profile.default_content_setting_values.notifications" : 2}
    chrome_options.add_experimental_option("prefs",prefs)
        
    chrome_path = r"chromedriver.exe"
    driver = webdriver.Chrome(chrome_path,options=chrome_options)
    #driver = webdriver.Chrome(chrome_options=chrome_options)
    #driver.set_window_position(-10000,0)
    driver.maximize_window()
    return driver


def loadContentPlaceholder(soup):
    a = []
    try:
        table = soup.find('table', {'id':'ContentPlaceHolder1_rptBLNo_gvLatestEvent_0'})
        tbody = table.find('tbody')
        trs = tbody.find_all('tr')
        return trs
    except:
        return a


def findLastInsertedRow(website, sqlcursor, mscursor,conn):
    searchkey = ''
    count = 0
    try:
        sqlcursor.execute("select top 1 *from dbo.BL where ID_FONTE ='"+website+"' order by ID_BL desc")
        rows= sqlcursor.fetchall()

        mscursor.execute("select distinct *from Data_Crawler where WEBSITE = '"+website+"'" )
        msrows = mscursor.fetchall()    
    except:
        print('failed to fetch from db due to db lock operation')
    for row in rows:
        searchkey = row[1]  

    if len(rows)>0:
        i = 0
        for trow in msrows:
            i = i+1
            searchKey = searchkey.replace(website,'').strip()
            rowtext = trow[1].replace(website, '').strip()
            if(searchKey == rowtext ):
                count = i
                
                sqlcursor.execute("select *from dbo.BL_MOVIMENT where BL='"+searchKey+"'") 
                blrows =sqlcursor.fetchall()
                if(len(blrows)<1):
                    sqlcursor.execute("Delete from dbo.BL where BL ='"+searchKey+"'")
                    conn.commit()
                    print('deleted from BL')
                    count = count-1
        
    return count



tree = ET.parse('server.xml')
root = tree.getroot()
mdb_tag = root.find('MDB')
dsn = mdb_tag.get('DSN')
uid = mdb_tag.get('UID')
pwd = mdb_tag.get('PWD')
server_tag = root.find('SERVER')
server = server_tag.get('SERVER')
username = server_tag.get('USERNAME')
password = server_tag.get('PASSWORD')
database = server_tag.get('DATABASE')

'''
conn1 = pyodbc.connect('DSN='+dsn+';UID='+uid+';PWD='+pwd)
cursor = conn1.cursor()
rows= cursor.execute("select * from Data_Crawler_YMLU where WEBSITE = 'YMLU'").fetchall()'''


conn = pymssql.connect(server=server, user=username, password=password, database=database) 
excursor = conn.cursor()


conn1 = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\Uttam\Documents\Database2.accdb')
cursor = conn1.cursor()
cursor.execute("select distinct *from Data_Crawler where WEBSITE = 'YMLU'")
rows = cursor.fetchall()
conn1.commit()



driver = getDriver()
path = "https://www.yangming.com/e-service/track_trace/track_trace_cargo_tracking.aspx"
driver.get(path)
time.sleep(1)
timeout = 10


def searchTrackingNumber(searchItem):
    element_present = EC.presence_of_element_located((By.ID,"ContentPlaceHolder1_rdolType_0"))
    WebDriverWait(driver, timeout).until(element_present)
    search_key = searchItem.replace('YMLU','')
    search_key = searchItem.replace('N','') #row[1]
    
    driver.find_element_by_id("ContentPlaceHolder1_rdolType_1").click()
    search_input = driver.find_element_by_css_selector('#ContentPlaceHolder1_num1')
    search_input.clear()
    search_input.send_keys(search_key)
    search_input.send_keys(Keys.ENTER)

        
    

lastRow = findLastInsertedRow('YMLU',excursor,cursor,conn)
captchaPage = "https://www.yangming.com/VerifyYourID.aspx"

#for row in rows:
for i in range(lastRow,len(rows)):
    try:
        print('Starting Step: '+str(i+1))
        search_key = rows[i][1]
        print(search_key)
        searchTrackingNumber(search_key)
        #searchTrackingNumber("YMLUB951052782")

        
        page_content = driver.page_source
        soup = bfs(page_content, "html.parser")
        if soup is None:
            print('Search result of '+ search_key + ' is empty. Trying next item')
        else:            
            trs = loadContentPlaceholder(soup)
    
            tr_len = len(trs)
            if(tr_len ==0):
                print('Search result of '+ search_key + ' is empty. Trying next item')
                
                if(driver.current_url == captchaPage ):
                    print("Captcha page found. exit program. Need to rerun to continue")
                    sys.exit()
            try:
                excursor.execute("INSERT INTO dbo.BL(BL,ID_FONTE) VALUES ('"+search_key+"','YMLU')")
                excursor.execute("SELECT SCOPE_IDENTITY()")
                bl_id = excursor.fetchone()[0]
                conn.commit()
            except:
                print('BL insert error')
            count = 0
            while count < tr_len:
                soup = bfs(driver.page_source,'html.parser')
                trs = loadContentPlaceholder(soup)
                tds = trs[count].find_all("td")
                con_num = tds[0].text
                temp = tds[2].text
                temp = temp.strip()
                pos = temp.find(" ")
                temp = temp[0:pos]
                con_cap = tds[1].text + " "+temp
                con_cap = con_cap.replace("'","`")
                print(con_num)
                print(con_cap)
                try:
                    excursor.execute("INSERT INTO dbo.EQP_MOVIMENT(ID_BL, ContainerNumber,Container_Capacity) VALUES ("+str(bl_id)+",'"+str(con_num)+"','"+str(con_cap)+"')")
                except:
                    print('EQP_MOVIMENT insert error')
                a_tag = tds[0].find("a")
                #a_tag.click()

                tab = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.LINK_TEXT,a_tag.text)))
                ActionChains(driver).move_to_element(tab).click().perform()

                soup=bfs(driver.page_source, 'html.parser')
                #containerInformation(soup)
                try:
                    detail_table = soup.find('table',{'id':'ContentPlaceHolder1_gvContainerNo'})
                    detail_tbody = detail_table.find('tbody')
                    de_trs = detail_tbody.find_all('tr')
                    for de_tr in de_trs:
                        de_tds = de_tr.find_all('td')
                        date = de_tds[0].text
                        date = date.replace('\n','')
                        datetime_object = datetime.strptime(date,'%Y/%m/%d %H:%M')
                        date = datetime_object.strftime("%Y-%m-%d %H:%M:%S")
                        status = de_tds[1].text
                        sp = status.find("(")
                        if sp != -1:
                            status = status[0:sp]
                        place = de_tds[2].text
                        pos = place.find("(")
                        if pos != -1:
                            place = place[0:pos]
                        temp = de_tds[4].text
                        pp = temp.find("(")
                        if pp != -1:
                            transport = temp[0:pp]
                            voy = temp[pp+1:len(temp)-1]
                        else:
                            transport = temp
                            voy = ""
                        transport= transport.replace("\n","")
                        voy = voy.replace("\n","")
                        print(date)
                        print(status)
                        print(place)
                        print(transport)
                        print(voy)
                        try:
                            excursor.execute("INSERT INTO dbo.BL_MOVIMENT(ID_BL,BL, Status,Place_Active,Time,Transport,Voy) VALUES ("+str(bl_id)+",'"+search_key+"','"+str(status)+"','"+str(place)+"','"+str(date)+"','"+str(transport)+"','"+str(voy)+"')")
                        except:
                            print('BL_MOVIMENT insert error')
                except TimeoutException:
                        print("Timed out waiting for page to load")            
                driver.back()
                count = count + 1
                #time.sleep(2)
            conn.commit()
            driver.back()
            time.sleep(randrange(3))
            print('Completed step:'+str(i+1))
    except TimeoutException:
        print("Timed out waiting for page to load")
        

cursor.close()
excursor.close()
conn1.close()
conn.close()
driver.close()
    
