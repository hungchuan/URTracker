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

#PROJECT_IDS = [1429]
# PROJECT_IDS = [712, 746, 750]

ALL_PROBLEMS = "/media/sf_linux-share/all.txt"
PROBLEM = "~/Downloads/ProblemList.xls"

def DF2List(df_in):
    train_data = np.array(df_in)#np.ndarray()
    df_2_list=train_data.tolist()#list
    return df_2_list
    
def download_issues (br, proj, filename):
    print('proj=',proj)
    #proj = int(proj)
    print('Downloading issues for project: %d' % (proj))
    print('urt=', URT.URTRACKER_URL + "/Pts/issuelist.aspx?project=%d" % proj)
    
    br.visit (URT.URTRACKER_URL + "/Pts/issuelist.aspx?project=%d" % proj)
    
    project = br.find_by_xpath('//*[@id="ctl00_pnlHeader"]/table/tbody/tr/td')
    project_name = project.text
    print('project_name = ',project_name)
    name_find= project_name.find('項目列表 »')
    len_title = len('項目列表 » ')
    print('len_title =',len_title)
    gh_name=project_name[name_find+len_title:]
    print('gh_name =',gh_name)
    newfilename = gh_name+'.xls'
    print('newfilename =',newfilename)
    
    result_info = br.find_by_xpath('//*[@id="ctl00_CP1_lblResultInfo"]')
    print('result_info = ',result_info.text)
    
    button= br.find_by_text('所有').click() #click 所有    
    #button = br.find_by_xpath('//*[@id="ctl00_CP1_tvNavt11"]').click() #click 所有
    #URT.wait_for_xpath (br, '//*[@id="ctl00_CP1_tvNavt11"]')
    
    count = 0
    result_info = br.find_by_xpath('//*[@id="ctl00_CP1_lblResultInfo"]')

    while (result_info.text != '所有事務'):
        print ('waiting for update Count: %d' % count)
        count += 1
        time.sleep (1)
        result_info = br.find_by_xpath('//*[@id="ctl00_CP1_lblResultInfo"]')

        
    button= br.find_by_text('導出').click() #click 導出    
    #button = br.find_by_xpath('//*[@id="ctl00_CP1_lnkExport"]').click()#click 導出
    #URT.wait_for_xpath (br, '//*[@id="ctl00_CP1_lnkExport"]')

    #URT.wait_for_xpath (br, "//*[@id='ctl00_CP1_lnkExport']")
    
    filename = os.path.expanduser (filename)
    print('filename_aaa = ',filename)
    
    if (os.path.isfile(filename)==True):
        print('remove file')
        os.remove(filename)
    
    time.sleep(1) 
    #URT.wait_for_xpath (br, "//*[@class='ctl00_CP1_tvNav_0']")
    # //*[@id="ctl00_CP1_btnExport"]
    
   
    #br.find_by_xpath ('//*[@id="ctl00_CP1_btnExport"]').first.click ()

    #URT.wait_for_xpath (br, '//*[@id="ctl00_CP1_btnExport"]')

    #filename = os.path.expanduser (filename)
    #print('filename = ',filename)
    #if os.path.exists (filename):
    #    os.unlink (filename)


    
    count = 0
    result_info = br.find_by_xpath('//*[@id="ctl00_CP1_Label1"]')

    while (result_info.text != '導出事務列表'):
        print ('waiting for update Count: %d' % count)
        count += 1
        time.sleep (1)
        result_info = br.find_by_xpath('//*[@id="ctl00_CP1_Label1"]')

        
    #br.find_by_xpath ('//*[@id="ctl00_CP1_btnExport"]').first.click ()
    br.find_by_xpath ('//*[@id="ctl00_CP1_btnExport"]').click () #click 導出
    #URT.wait_for_xpath (br, '//*[@id="ctl00_CP1_btnExport"]')
    #print('filename111 = ',filename)
    #time.sleep(2) 
    #filename = URT.complete_download (filename, proj)
    
    count = 0
    while (os.path.exists (filename) == False):
        print ('waiting for download Count: %d' % count)
        count += 1
        time.sleep (1)
    
    
    df = pd.read_excel(filename)
    new_df=df.fillna(0)
    
    print('newfilename = ',newfilename)
    PACKAGE_DIRECTORY = os.path.abspath('.')
    newfilename = os.path.join(PACKAGE_DIRECTORY,newfilename)
    print('newfilename222 = ',newfilename)
    shutil.move(filename,newfilename) 
    #new_df.to_excel(newfilename)
    
    try:
        gc = pygsheets.authorize(service_file='PythonUpload-cfde37284cdc.json')
    except:
        return filename
    
    sh = gc.open('URTracker')

    try:
        wks = sh.worksheet_by_title(gh_name)
    except:
        wks = sh.add_worksheet(gh_name,rows=9000,cols=100,index=0)    
     
    wks.set_dataframe(new_df, (1, 1))
    
    return filename
   
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


def main (args):
    urtracker_config = 'urtracker_config'
    print('args = ',args)
    projdata = []
    #print('args [1] = ',args [1])
    config = URT.read_config (urtracker_config)
    #br = splinter.Browser ('chrome', profile=config ['chrome_profile_path'])
    br = splinter.Browser(driver_name='chrome')

    URT.login (br, config ['username'], config ['password'])
    
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

if __name__ == '__main__':
    main (sys.argv)
