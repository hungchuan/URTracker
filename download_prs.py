#!/usr/bin/env python

import os
import sys
import re
import xlrd
import splinter
import urtracker as URT
import time
import pandas as pd
import pygsheets
import numpy as np
import shutil 
import pdb
from selenium import webdriver
from selenium.webdriver.support.select import Select

#from datetime import datetime, date, timedelta

from print_log import log_print,Emptyprintf

#PROJECT_IDS = [1429]
# PROJECT_IDS = [712, 746, 750]

ALL_PROBLEMS = "/media/sf_linux-share/all.txt"
#PROBLEM = "~/Downloads/ProblemList.xls"
PROBLEM = "ProblemList.xls"

def DF2List(df_in):
    train_data = np.array(df_in)#np.ndarray()
    df_2_list=train_data.tolist()#list
    return df_2_list

def waiting_for_update(br,xpath):
    button = False
    for i in range(0,10):
        time.sleep(1) 
        print ('waiting for update Count: %d' % i)
        try:
            button = br.find_element_by_xpath(xpath)
        except:
            continue
        break
    print(xpath)
    print("waiting_for_update button=",button)
    return button    

def download_from_google(file, sheet):

    JSON_DIRECTORY = os.path.abspath('.')    

    JSON_name = os.path.join(JSON_DIRECTORY,'PythonUpload-cfde37284cdc.json')
    print ('JSON_name= %s' % JSON_name)

    #gc = pygsheets.authorize(service_file=JSON_name)
    #gc = pygsheets.authorize()
    
    try:
        gc = pygsheets.authorize(service_file=JSON_name)
        #gc = pygsheets.authorize(service_file='PythonUpload-cfde37284cdc.json')
    except:
        print("can not find json file")
    
    sh = gc.open(file)

    try:
        wks = sh.worksheet_by_title(sheet)
    except:
        wks = sh.add_worksheet(sheet,rows=1,cols=30,index=0)    
     
    df = wks.get_as_df()   
    return df    

def upload_to_google(file, sheet, TCA_df):
    try:
        gc = pygsheets.authorize(service_file='PythonUpload-cfde37284cdc.json')
    except:
        return filename
    
    sh = gc.open(file)

    try:
        wks = sh.worksheet_by_title(sheet)
    except:
        wks = sh.add_worksheet(sheet,rows=1,cols=30,index=0)    
     
    wks.set_dataframe(TCA_df, (1, 1))

    return wks        
    
def download_issues (br, proj, filename,directory,config):

    print ('= %s' % directory)
    print ('filename= %s' % filename)
    ProblemList_name = os.path.join(directory,filename)
    print ('ProblemList_name= %s' % ProblemList_name)
    
    print('proj=',proj)
    #proj = int(proj)
    print('Downloading issues for project: %d' % (proj))
    print('urt=', URT.URTRACKER_URL + "/Pts/issuelist.aspx?project=%d" % proj)
    
    #sel = input("pause 58")    
    
    #br.visit (URT.URTRACKER_URL + "/Pts/issuelist.aspx?project=%d" % proj)
    #br.get(URT.URTRACKER_URL + "/Pts/issuelist.aspx?project=%d" % proj)
    #sel = input("pause 62")    
    
    result_info = waiting_for_update(br,config ['xpath_mylist'])    
    print (result_info.text)
    while (result_info.text != "我的事務"):
        print (result_info.text)
        #count += 1
        time.sleep (1)
        result_info = waiting_for_update(br,config ['xpath_mylist'])    
    
    br.get (URT.URTRACKER_URL + "/Pts/issuelist.aspx?project=%d" % proj)
    
    time.sleep(5)  
         
    project = waiting_for_update(br,config ['xpath_project'])
    project_name = project.text
    print('project_name = ',project_name)
    name_find= project_name.find('項目列表 »')
    len_title = len('項目列表 » ')
    print('len_title =',len_title)
    gh_name=project_name[name_find+len_title:]
    print('gh_name =',gh_name)
    newfilename = gh_name+'.xls'
    print('newfilename =',newfilename)
   
    button = br.find_element_by_partial_link_text("所有")
    button.click() #All     

    result_info = waiting_for_update(br,config ['xpath_all_str'])
    while (result_info.text != "所有事務"):
        print (result_info.text)
        time.sleep (1)
        result_info = waiting_for_update(br,config ['xpath_all_str'])    
        
    button = waiting_for_update(br,config ['xpath_export'])
    button.click() #export     
    
    result_info = waiting_for_update(br,config ['xpath_export_list'])
    while (result_info.text != "導出事務列表"):
        print (result_info.text)
        time.sleep (1)
        result_info = waiting_for_update(br,config ['xpath_export_list'])
        
    button = waiting_for_update(br,config ['xpath_download'])    
    button.click() #download
    time.sleep(1)  

    count = 0
    while (os.path.exists (ProblemList_name) == False):
        print ('waiting for download Count: %d' % count)
        count += 1
        time.sleep (1)
    
    URT_df = pd.DataFrame(columns=['#','事務編碼','待辦人','狀態','Subject','Symptom category','Priority','Designer priority','Module','Severity','CS Priority','SWQE suggest priority fix','Actual Domain Owner'])
    
    df = pd.read_excel(ProblemList_name)
    #new_df=df.fillna(0)

    df_column=URT_df.columns.values.tolist()
    
    for i in df_column:
        try:
            URT_df[i]=df[i] 
            print('Try OK',i)
        except:
            print('Try NOK',i)
    
    new_URT_df=URT_df.fillna("")
    
    print('newfilename = ',newfilename)
    PACKAGE_DIRECTORY = os.path.abspath('.')
    newfilename = os.path.join(PACKAGE_DIRECTORY,newfilename)
    print('newfilename222 = ',newfilename)
    
    if (os.path.isfile (newfilename)):
        try:
            os.remove(newfilename)
        except OSError as e:
            print(e)
        else:
            print("File is deleted successfully")
    
    shutil.move(ProblemList_name,newfilename) 
    #new_df.to_excel(newfilename)
    
    current_df=download_from_google(gh_name,gh_name) 
    df_zero=current_df.copy()
    for col in df_zero.columns: 
        df_zero[col]="" 
    upload_to_google(gh_name,gh_name,df_zero) # upload to google sheet    
    time.sleep(3)  
    upload_to_google(gh_name,gh_name,new_URT_df) # upload to google sheet       
    
    return newfilename
   
'''    
    elem = None
    for e in br.find_by_xpath ("//*[@class='ctl00_CP1_tvNav_0']"): 
        if re.search ("Open", e.value):
            elem = e

    if elem != None:
        elem.click ()

        URT.wait_for_update_progress (br, "//*[@id='ctl00_UpdateProgress1']")

        br.find_by_xpath ("//*[@id='ctl00_CP1_lnkExport']").first.click ()

        URT.wait_for_xpath (br, "//*[@id='ctl00_CP1_btnExport']")

        filename = os.path.expanduser (filename)
        if os.path.exists (filename):
            os.unlink (filename)

        br.find_by_xpath ("//*[@id='ctl00_CP1_btnExport']").first.click ()

        filename = URT.complete_download (filename, proj)
        return filename
    else:
        raise Exception ("Couldn't download file for proj: %d" % (proj))

'''
def combine_problems (prdataset, target):
    SEP = "$;"
    columns = ["#", "Issue Code", "Assignee", "Last Process User", "Pillar", "Subsystem", "APK", "Subject", "State", "PR Due Date Initial", "PR Due Date Revised"]

    all_prs = []
    
    for prdata in prdataset:
        book = xlrd.open_workbook (prdata)
        sheet = book.sheets() [0]

        colmap = {}

        for colid in range (0, sheet.ncols):
            header = sheet.cell_value (0, colid).strip ()
            if header in columns:
                colmap [header] = colid

        for rowid in range (1, sheet.nrows):
            prval = {}
            for colkey in columns:
                if sheet.cell_type (rowid, colmap [colkey]) == xlrd.XL_CELL_NUMBER:
                    prval [colkey] = int (sheet.cell_value (rowid, colmap [colkey]))
                else:
                    prval [colkey] = unicode (sheet.cell_value (rowid, colmap [colkey]))
            all_prs.append (prval)

    tf = open (target, "wb")

    tf.write (SEP.join (columns) + "\n")

    for r in all_prs:
        for c in columns:
            if c == "#":
                tf.write ("%d" % (r [c]))
            else:
                if len (r [c]) != 0:
                    tf.write (r [c].encode ("utf-8"))
                else:
                    tf.write (" ")
            tf.write (SEP)
        tf.write ("\n");

def file_download (directory,config):
    projdata = []
    
    if os.path.exists(directory):
        shutil.rmtree(directory)
    else:       
        # 使用 try 建立目錄
        try:
            os.makedirs(directory)
        # 檔案已存在的例外處理
        except FileExistsError:
            print("目錄已存在。")    
    

    single_pr= config
    config = 'config'
    
    config = URT.read_config (config)   
    
    if single_pr!='config':
        config ['project_ids'] = single_pr
        
	
    options = webdriver.ChromeOptions()
    prefs = {'profile.default_content_settings.popups': 0, 'download.default_directory': directory}
    options.add_experimental_option('prefs', prefs)
    #br = webdriver.Chrome(executable_path='D:\Python\Python37-32\chromedriver.exe', chrome_options=options)
    
    if (os.path.isfile ('c:\chromedriver.exe')):
        print('c:\chromedriver.exe')
        br = webdriver.Chrome(executable_path='c:\chromedriver.exe', options=options)
    else:
        print('False')    
        br = webdriver.Chrome(options=options)
    
    URT.login (br, config)
    
    PROJECT_IDS_str = config ['project_ids']
    print('PROJECT_IDS_str =',PROJECT_IDS_str)
    PROJECT_IDS = [int(s) for s in PROJECT_IDS_str.split() if s.isdigit()]
    print('PROJECT_IDS =',PROJECT_IDS)
    #time.sleep(10) 
    all_problems = os.path.expanduser (ALL_PROBLEMS)
    print('all_problems = ',all_problems)
    
    for prjid in PROJECT_IDS:
        projdata.append (download_issues (br, prjid, PROBLEM,directory,config))

    br.quit ()
    return True  
      
        
def main (args):
    global log
    config = 'config'
    print('args = ',args)
    args_len = len(args)
    
    log = Emptyprintf
    
    if (args_len==2):     
        if (args[1]=="debug"):
            log = log_print
        else:
            log = Emptyprintf
            config = args[1]

        
    CURRENT_PACKAGE_DIRECTORY = os.path.abspath('.')    
    PACKAGE_DIRECTORY = CURRENT_PACKAGE_DIRECTORY + '\download' 
    Backup_DIRECTORY = CURRENT_PACKAGE_DIRECTORY + '\\backup' 
    
    download_file = file_download(PACKAGE_DIRECTORY,config) #下載URT檔案
    
'''    
    PROJECT_IDS_str = config ['project_ids']
    print('PROJECT_IDS_str =',PROJECT_IDS_str)
    PROJECT_IDS = [int(s) for s in PROJECT_IDS_str.split() if s.isdigit()]
    print('PROJECT_IDS =',PROJECT_IDS)
    #time.sleep(10) 
    all_problems = os.path.expanduser (ALL_PROBLEMS)
    print('all_problems = ',all_problems)
    
    for prjid in PROJECT_IDS:
        projdata.append (download_issues (br, prjid, PROBLEM))

    #combine_problems (projdata, all_problems)
    #print ('all projects combined into file: %s' % (all_problems))

    br.quit ()
'''
if __name__ == '__main__':
    main (sys.argv)
