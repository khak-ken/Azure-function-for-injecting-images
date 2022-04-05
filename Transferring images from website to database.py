from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.expected_conditions import presence_of_element_located
import time
import os
import pandas as pd
import requests

driver = None
#login
cpath = "C:/Users/test/chromedriver"
driver = webdriver.Chrome(cpath)

master_file_raw=pd.read_excel(r'C:\Users\test\test.xlsx',sheet_name='test')
master_file=master_file_raw.dropna(subset=['Link'])
master_file=master_file.drop_duplicates(subset ='Link',keep='first') 
master_file=master_file[master_file['Condition1']!='--']
#master_file.to_csv(r'C:\Users\test\test2.csv')

master_file_part1=master_file.iloc[20:30]

driver.get("https://test.test.com/login")
e = driver.find_element(By.ID, "userNameField")
e.send_keys("TestUser")
e = driver.find_element(By.ID, "userPassField")
e.send_keys("TestPassword")
e = driver.find_element(By.ID, "UserLoginForm")
e.submit()

 
def extract_every_report(): 
    #get report:
    driver.get(report_url)
    wait = WebDriverWait(driver, 10)
    wait.until(presence_of_element_located((By.CLASS_NAME, "pdf-summary")))
    time.sleep(3)  
    output= driver.find_elements_by_class_name("pdf-summary")
    if len(output)==0:
        driver.get("https://test.test.com/login")
        e = driver.find_element(By.ID, "userNameField")
        e.send_keys("TestUser")
        e = driver.find_element(By.ID, "userPassField")
        e.send_keys("TestPassword")
        e = driver.find_element(By.ID, "UserLoginForm")
        e.submit()
        #get report:
        driver.get(report_url)
        wait = WebDriverWait(driver, 10)
        wait.until(presence_of_element_located((By.CLASS_NAME, "pdf-summary")))
        time.sleep(3)
        output= driver.find_elements_by_class_name("pdf-summary")
        
    # # Connection to DB:
    import pyodbc
    cnxn = pyodbc.connect(os.environ.get('SQL_ODBC_CONNECTION'))
    cursor = cnxn.cursor()   
    crt= ''' IF NOT EXISTS (SELECT name FROM sys.tables WHERE name='myimages') 
    CREATE TABLE myimages(Id int IDENTITY(1,1) NOT NULL,report_id varchar(20) NOT NULL,test_info_2 varchar(max),test_info_3 varchar(max),
    test_info_4 varchar(max),test_info_5 varchar(max),test_info_6 varchar(max),test_info_7 varchar(max),
    test_info_8 varchar(max),test_info_9 varchar(max),test_info_10 varchar(max),class_from_report int,
    class_id int, img varbinary(max),
    PRIMARY KEY CLUSTERED  (Id ASC),FOREIGN KEY (class_id) REFERENCES class(class_id),
    FOREIGN KEY (class_from_report) REFERENCES class(class_id))''' 
    cursor.execute(crt)
    
    def extract_middle_str(line,start,end):
        return line[line.find(start)+len(start):line.rfind(end)].strip()
    
    for quote in output:
          text = quote.text.split('\n')
          reference=extract_middle_str(text[0],'Test:','test:')
          test_a=extract_middle_str(text[0],'test:','')
          
          
    #component
    class_from_intellispec=driver.find_elements_by_xpath('/html/body/div[1]/div[6]/div[4]/div[2]/div/table[4]/tbody/tr[3]')
    class_from_report=extract_middle_str(class_from_intellispec[0].text,'2.  of Rust','')
    class_from_report=int(class_from_report)
    if(class_from_report in [1,2]):
        class_adj=1
    elif(class_from_report in [3,4]):
        class_adj=2
    elif(class_from_report in [5,6]):
        class_adj=3
    elif(class_from_report in [7,8]):
        class_adj=4 
                  
    images=driver.find_elements_by_xpath("//img[starts-with(@alt,'Report Image')]") 
    cnt=1
    for img in images:
        print(cnt)
        src=img.get_attribute('src')
        #print(src)
        file=requests.get(src).content
        print("class before inserting:",class_adj)
        qry = '''INSERT INTO myimages values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) '''
        class_from_report="SELECT class from class where class.class_id=myimages.class_from_report"
        param_values = [url_id+'_'+str(cnt),test etc]
        cursor.execute(qry,param_values)    
        print('{0} row inserted successfully.'.format(cursor.rowcount))
        cursor.commit()
        cnt=cnt+1
        
    cursor.close()
     
for ind in master_file_part2.index: 
    report_url=master_file_part2['Link'][ind]
    url_id=report_url.rsplit('/', 1)[-1] 
    extract_every_report()


from azure.storage.blob import BlobServiceClient, BlobClient, ContainerClient
import requests

secret=os.environ.get('AzureWebJobsStorage')
serviceclient = BlobServiceClient.from_connection_string(conn_str=secret)

def write_img_to_blob(img):
    img='Report.jpg'
    blob_client = serviceclient.get_blob_client(container='scapedimages', blob='./'+img)
    blob_client.upload_blob(file, overwrite=True)  
    
write_img_to_blob('Imahe.png')
