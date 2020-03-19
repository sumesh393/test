from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
import time
import openpyxl as op
import pyautogui as pa
from pynput.keyboard import Key, Controller
from keyboard import press
import os
#import os
import io
import os, shutil
import datetime
from datetime import date
from pdfminer.converter import TextConverter
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfpage import PDFPage
import re
import openpyxl
from openpyxl import load_workbook,Workbook




##
##
try:
   
   os.remove(r'D:\Users\Kaushal\AppData\Local\Programs\Python\Python38-32\offc work\indiamart\pdfindi\e1.pdf')
   print("file removed")
except:
   a="no file found"
   print(a)


dt = []
url = "https://mahatenders.gov.in/nicgep/app"
kurl="N.A"


driver = webdriver.Chrome() 
driver.maximize_window()
driver.get(url)
opening_date = driver.find_elements_by_xpath('//*[@id="PageLink_17"]')
for i in opening_date:
   i.click()
time.sleep(1)

         
td = driver.find_elements_by_xpath('//*[@id="DirectLink"]')
for k in td:
   kurl=k.get_attribute('href')
print("tender Url :-",kurl)


for j in td:
   j.click()
time.sleep(1)


pa.hotkey('ctrl', 'p')
time.sleep(5)
pa.click(1401,194)
time.sleep(2)
pa.click(1414,239)
time.sleep(2)
pa.click(1405,871)
time.sleep(4)
pddf=pa.write(r"C:\Users\sumesh\Desktop\India mart\e1.pdf")
#for i in r

time.sleep(400)
pa.click(758,554)
time.sleep(2)
pa.click(758,554)
print("saved")


def extract(pdf_path):
  
  resource_manager = PDFResourceManager()
  fake_file_handle = io.StringIO()
  converter = TextConverter(resource_manager, fake_file_handle)
  page_interpreter = PDFPageInterpreter(resource_manager, converter)
  with open(pdf_path, 'rb') as fh:
  
      for pg in PDFPage.get_pages(fh,
                                    caching=True,
                                    check_extractable=True):
          page_interpreter.process_page(pg)
      text = fake_file_handle.getvalue()
      converter.close()
      fake_file_handle.close()
      print(text)

      #tender notice
      tn=re.search(r'Reference.*?Tender',text).group()
      tn1=re.search(r'Number.*?Tender',tn).group()
      tn1=tn1.replace('Number','')
      tn1=tn1.replace('Tender','')
      
      #tnotice=re.search(r'\d{2}(.)\d{2}(-)\d{2}',tn).group()
      print('Tender notice no :-',tn1)

      #tender type
      ty=re.search(r"Reference.{200}",text).group()
      tenderty=re.search(r"Type.*?Form",ty).group().replace("Type",'').replace("Form",'').replace(" ",'')
      print("Tender Type :-",tenderty)

      #tender catagory
      tenderc=re.search(r"Category.*?No",text).group().replace("Category",'').replace("No",'').replace(" ",'')
      print("Tender Category :-",tenderc)

      #ernest money
      emd=re.search(r"Amount.*?EMD",text).group()
      ernest=re.search(r"₹.*?E",emd).group().replace('E','').replace(' ','')
      print("Earnest Money :- ",ernest)

      #Published date

      pd=re.search(r'Published.*?Bid',text).group()
      publishdate=re.search(r'Date.*?Bid',pd).group().replace("Date",'').replace("Bid",'')#.replace(" ","")
      print("Published date :-",publishdate)

      #Tender value

      tv=re.search(r'Value.*?Product',text).group()
      tenderval=re.search(r'₹.*?P',tv).group().replace('P','')
      print("Tender value :-",tenderval)

      #Tenderfee
      tf=re.search(r'Total.*?Payable',text).group()
      tf1=re.search(r"Tender.*?P",tf).group()
      tenderfee=re.search(r"₹.*?P",tf1).group().replace('P','').replace("Fee",'')
      print('Tender fee :-',tenderfee)

      #Authority name
      an=re.search(r'Authority.*?Address',text).group()
      authname=re.search(r'Name.*?Address',an).group().replace("Name",'').replace("Address",'')
      print("Authority name :-",authname)

      #Authority add
      ad=re.search(r'Authority.*?Back',text).group()
      authadd=re.search(r'Address.*?Back',ad).group().replace('Address','').replace('Back','')
      print("Authority Address :-",authadd)

      #location
      loc=re.search(r'Period.*?Pincode',text).group()
      location=re.search(r'Location.*?Pincode',loc).group().replace('Location','').replace('Pincode','')
      print("Tender Location :-",location)

      #Bid submission start date
      bsd=re.search(r'Submission.*?Bid',text).group()
      bidstart=re.search(r'Date.*?Bid',bsd).group().replace("Date",'').replace("Bid",'')
      print("Bis submission start date :-",bidstart)
      
      #Bid submission end date
      bed=re.search(r'Submission.*?Tender',text).group()
      bidend=re.search(r'End.*?Tender',bed).group().replace('End','').replace('Tender','').replace('Date','')
      print("Bid submission end date :-",bidend)

      #Bid opening date
      bod=re.search(r'Published.{130}',text).group()
      bod1=re.search(r'Bid.*?Document',bod).group()
      bidopen=re.search(r'Date.*?Document',bod1).group().replace('Date','').replace('Document','')
      print('Bid opening date :-',bidopen)

      #Tender Title
      tt=re.search(r'EMD.*?NDA',text).group()
      tt1=re.search(r'Title.*Work',tt).group()
      tt1=tt1.replace('Title','')
      tt1=tt1.replace('Work','')
      print("Tender Title :-",tt1)

      #Work Description
      #wdd=re.search(r'Fee.*?Qu',text).group()
      wd=re.search(r'EMD.*?NDA',text).group()
      workdes=re.search(r'Work.*?NDA',wd).group()
      workdes=workdes.replace('Work','')
      workdes=workdes.replace('NDA','')
      print("Work description :-",workdes)

      #Project State:
      prostate="Maharastra"
      print("Project state :-",prostate)

      #Project country:
      country="India"
      print("Tender Country",country)

      #product Catagory:
      pcat=re.search(r'Product.*?Sub',text).group()
      pcat=pcat.replace('Product','')
      pcat=pcat.replace('Category','')
      pcat=pcat.replace('Sub','')
      print("Product Catagory :-",pcat)

      #Document sale start date
      dsd=re.search(r'Sale.{100}',text).group()
      dsd1=re.search(r'Date.*?Document',dsd).group()
      dsd1=dsd1.replace('Date','')
      dsd1=dsd1.replace('Document','')
      print("Document sale start date :-",dsd1)

      #Document sale end date
      ded=re.search(r'Sale.{100}',text).group()
      ded1=re.search(r'End.*?Clarification',ded).group()
      ded2=re.search(r'Date.*?Clarification',ded1).group()
      ded2=ded2.replace('Date','')
      ded2=ded2.replace('Clarification','')
      print('Document sale end date  :-',ded2)


      #Product name:
      #pn=re.search()
      product="N.A"
      condition="False"
      list=["Flowers","High Security Registration Plates","R.O.Plant","SOLAR STREET LIGHT AND SOLAR PUMP","White LED"]
      for i1 in list:
         if i1 in tt1:
            product=i1
            condition="True"
     


      wb=op.load_workbook('tenderauto.xlsx')
      ws=wb.active

      ws.append(['Tender Notice NO','Tender type','Product category','Authority Name','Project State','EMD','Tender Value','Tender Country',
                 'Contact Email ID','Authority website','Tender Title','Tender Description','Bid Open Date','Phone no',
                 'Fax no','Document sale end date','Product name','Tender publish date','Tender document url'])
      wb.save(filename='tenderauto.xlsx')

      data=[tn1,tenderty,pcat,authname,prostate,ernest,tenderval,country,"",url,tt1,workdes,bidopen,"NA","NA",ded2,product,publishdate,kurl]
      #print(data)
      ws.append(data)
      wb.save(filename='tenderauto.xlsx')

      


for data in os.listdir(r'C:\Users\sumesh\Desktop\India mart'):
    data2 = (r'C:\Users\sumesh\Desktop\India mart\%s')%data
    if ".pdf" in data2:
        #print(data2)
        extract(data2)
  

driver.close()
      






