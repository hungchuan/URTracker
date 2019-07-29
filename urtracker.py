#!/usr/bin/env python

# encoding: utf-8
#
# Copyright 2011, Apurva Sharan
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#       http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

import os
import re
import time
import xlrd

URTRACKER_URL="https://fwtrack.tpv-tech.com"

def login (br, username, password):
    print ('Entering Login')
    br.visit (URTRACKER_URL + "/Accounts/Login.aspx")
    br.fill ("txtEmail", username)
    br.fill ("txtPassword", password)
    br.find_by_name ("btnLogin").first.click ()
    br.is_element_present_by_xpath ('//*[@id="ctl00_Siteheader1_lblGreeting"]', wait_time = 10)
    print ('Leaving Login')


def read_config (config):
    print('read_config')
    retmap = {}
    configs = open (config).readlines ()
    print('configs = ',configs)
    for config in configs:
        (k,v) = config.split ("=")
        if "USERNAME" == k.strip ():
            retmap ['username'] = v.strip ()
        if "PASSWORD" == k.strip ():
            retmap ['password'] = v.strip ()
        if "MOZILLA_PROFILE_PATH" == k.strip ():
            retmap ['mozilla_profile_path'] = v.strip ()
        if "PROJECT_IDS" == k.strip ():
            retmap ['project_ids'] = v.strip ()
    return retmap
    

def xpath_select_dropdown (br, xpath, text):
    for e in br.find_by_xpath (xpath):
        if re.search (text, e.text) != None:
            e.click ()
            return
    raise Exception ("Couldn't find required element: %s" % (text))

def xpath_click_first_checkbox (br, xpath):
    br.find_by_xpath (xpath).first.click ()

def wait_for_xpath (br, xpath):
    done = False
    count = 0
    while (done == False):
        try:
            if (len (br.find_by_xpath (xpath)) != 0):
                done = True
        except e:
            print ('Retrying for xpath %s .. count : %d' % (xpath, count))
            count += 1
            time.sleep (1)


def wait_for_update_progress (br, xpath):
    # Ensure that update spinner shows up and goes off
    datafill = br.find_by_xpath (xpath)
    while (datafill.visible == False): time.sleep (1)
    while (datafill.visible == True): time.sleep (1)


def ensure_file_doesnt_exist (filename):
    if (os.path.exists (filename) == True):
        os.unlink (filename)


def complete_download (filename, proj):
    #PACKAGE_DIRECTORY = os.path.abspath('.')
    
    #filename= os.path.join(PACKAGE_DIRECTORY,filename)   
    #print('filename =',filename)    
    
    filename = os.path.expanduser (filename)
    print('filename = ',filename)
    
    count = 0
       
    while (os.path.exists (filename) == False):
        print ('Count 1: %d' % count)
        count += 1
        time.sleep (1)

    while (os.path.exists (filename + ".part") == True):
        print ('Count 2: %d' % count)
        count += 1
        time.sleep (1)

    while (os.path.getsize (filename) == 0):
        print ('Count 3: %d' % count)
        count += 1
        time.sleep (1)

    # download complete- rename the file and return the new name
    target = "%s.%d" % (filename, proj)
    os.rename (filename, target)
    return target

    
