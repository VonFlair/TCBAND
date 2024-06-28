import pandas as pd
import os
import time
import selenium.webdriver as webdriver
from selenium.webdriver.common.by import By
from threading import Timer
#from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
import openpyxl

import json


DUT_list = ("MTK", "Hisi", "Huawei", "QCSD", "QIPL", "Samsung", "Unisoc", "A2650_IMS-Off", "VN009_Light", "Google", "Oppo", "OnePlus")

def modifydut(dut,*args):
    for x in args:
        if x.lower() in dut.lower():
    #        x = x.replace("QIPL","QC")
            return x
    return dut
    
    
def bandmatch(df_C_BAND_passed:pd.Series):
    df_C_BAND_passed = df_C_BAND_passed.str.replace("DC_","EN-DC_")
    df_C_BAND_passed = df_C_BAND_passed.str.replace("redcap","RedCap")
    df_C_BAND_passed = df_C_BAND_passed.str.replace('n48_2A', 'n48(2A)')
    df_C_BAND_passed = df_C_BAND_passed.str.replace('n66_2A','n66(2A)')
    df_C_BAND_passed = df_C_BAND_passed.str.replace('n71_2A','n71(2A)')
    df_C_BAND_passed = df_C_BAND_passed.str.replace('CA_', 'CA_DL_')
    df_C_BAND_passed = df_C_BAND_passed.str.replace('+','_')
    df_C_BAND_passed = df_C_BAND_passed.str.replace('LTE','')
    return df_C_BAND_passed

def response_url_time(driver:webdriver.Chrome,t):
    driver.execute_script("window.stop();")
    driver.refresh()
    t.stop()    

def response_element(driver:webdriver.Chrome):
    res = "False"
    looptimes = 0 
    while res == "False" or looptimes > 30:
        try:
            driver.find_element(By.TAG_NAME, 'pre')
            res = "True"
        except:
            #driver.execute_script("window.stop();")
            driver.refresh()
            time.sleep(1)
            looptimes += 1
    if looptimes > 30:
        element = webdriver.Remote._web_element_cls.value_of_css_property
        element.text = '{"abc":"abc"}' #workout
    else:
        element = driver.find_element(By.TAG_NAME, 'pre')
    return element

def element_check(driver,element,url):
    looptime = 0
    while element.text[-1] != "}":
    #driver.get(responseURL)
        driver.refresh()
        element = response_element(driver)
        looptime += 1
        if looptime > 5:
            looptime = 0
            driver.close()
            driver = webdriver.Chrome()
            driver.get(url)
    return element


def main(version_filter, verdict_filter, target_path):

    #filter_arg1 = input('Please provide Filter Argument for TC Type (NR5GC or ENDC')

    # URL Definitions
    STORAGE_URL = f"https://mrt-result-storage.cloud.rsext.net/Results/blobs/"


    base_url = 'http://mrt-reports.cloud.rsext.net/odata/ApplicationRuns'

    query_filter = []

    if version_filter:
        query_filter.append(f"contains(Application/Version, '{version_filter}')")

    if verdict_filter:
        query_filter.append(f"Verdict eq '{verdict_filter}'")

    filter_expression = " and ".join(query_filter)
    filter_string_input = f"{base_url}?$filter={filter_expression}"

    print('queries set: ', query_filter)
    print('Filter expression: ', filter_expression)
    print('Final Filter String: ', filter_string_input)

    # f"http://mrt-reports.cloud.rsext.net/odata/ApplicationRuns?$filter=contains(Application/Category,'PCT') and contains(Application/Version,'23.12.2') and Verdict eq 'Passed'"
    # f"http://mrt-reports.cloud.rsext.net/odata/ApplicationRuns?$filter=contains(Application/Category,'PCT') and contains(Application/Version,'23.12.2')"

    #url = 'https://mrt-reports.cloud.rsext.net/odata/ApplicationRuns?$filter=contains%28Application%2FVersion%2C%20%2723.37.1%27%29%20and%20Verdict%20eq%20%27Passed%27'


    # Session Setup
    responseURL = filter_string_input

    #responseURL = 'https://mrt-reports.cloud.rsext.net/odata/ApplicationRuns?$filter=contains%28Application%2FVersion%2C%20%2723.37.2%27%29%20and%20Verdict%20eq%20%27Passed%27'

    #new decode for the script
    driver = webdriver.Chrome()#Chrome driver need download driver from https://googlechromelabs.github.io/chrome-for-testing/#stable and put it in python.exe folder
    driver.maximize_window() #max Chrome window
    odlink = '@odata.nextLink' #check for the next page
    text = {odlink:responseURL} #web link
    driver.implicitly_wait(10)
    

    log_data = pd.DataFrame()
    start_loop = time.time()
    work_time = 0
    print(text[odlink])

    #loop for the page one by one
    while odlink in text :
        start_time = time.time()
        responseURL = text[odlink]
        try:
            driver.get(responseURL)
        except:
            print("response timeout: ",responseURL)
        element = response_element(driver)
        element = element_check(driver,element,responseURL)  
        text = json.loads(element.text)
        #driver.close()
        end_time =time.time()
        time.sleep(1)
        #print(end_time-start_time)
        work_time = end_time - start_loop
        onepage_data = pd.DataFrame(text['value'])
        log_data = pd.concat([log_data,onepage_data])
    print(responseURL)
    driver.quit()
    end_loop = time.time()
    print('Loop done')
    print('Loop time: ', end_loop - start_loop) # out the time of getting all data      


    #some info in Application and spread the column 'Application' move to row end
    log_data.reset_index(inplace=True, drop=True)
    tc_data = pd.DataFrame(log_data['Application'].to_list())
    df_all = pd.concat([log_data,tc_data],axis=1)
    df_all['CORC'] = STORAGE_URL + df_all['RunId']
    df_all['Dut_update'] = df_all['Dut'].apply(lambda x: x['Name'])
    df_all['DUT_modify'] =  df_all['Dut_update'].apply(modifydut, args=tuple(DUT_list))
    """
    df_all_new = df_all.copy()
    df_all_new['CORC'] = STORAGE_URL + df_all['RunId']
    df_all_new['DUT'] = df_all['Dut'].apply(lambda x:x['Name'])
    df_all_new['DUT_modify'] = df_all['DUT'].apply(modifydut,args=tuple(DUT_list))
    """
    workbook_name = 'Report_' + version_filter + verdict_filter + '.xlsx'
    target_path_filename = os.path.join(target_path,workbook_name)
    #workbook_name_new = str(target_path_filename).replace('.xlsx','_link.xlsx')
    workbook = pd.ExcelWriter(target_path_filename)
    #workbook1 = pd.ExcelWriter(workbook_name_new)
    print('Writing ' + target_path_filename)
    pd.DataFrame(df_all).to_excel(workbook, sheet_name= 'Report')
    #pd.DataFrame(df_all_new).to_excel(workbook1, sheet_name='Report')
    workbook.close()
    #workbook1.close()
    print('download finished')

"""
    df = df_all.copy()
    TC_BAND_passed = df['Title'] + df['RunConfiguration'] + df['Verdict']
    df.insert(loc = 1,column = 'TC_BAND_passed',value= TC_BAND_passed,allow_duplicates= True)
    #TC_BAND_passed.name = 'TC_BAND_passed'
    #Get the all DUT values
    log_group_count = pd.DataFrame(columns=['TC_BAND_passed','DUT','Number'])
    log_group_count['TC_BAND_passed'] = df['TC_BAND_passed'].drop_duplicates().sort_values().reset_index().drop(columns = 'index')
    #print(log_group_count)

    #group log and change the dut name
    for i in log_group_count.index:
        temp = log_group_count['TC_BAND_passed'][i]
        dftemp = df.query('TC_BAND_passed == @temp')
        tempdut = dftemp['DUT_modify'].drop_duplicates().to_list()
        #print(tempdut)
        log_group_count['DUT'][i] =  tempdut
        log_group_count['Number'][i] = len(tempdut)
	
    log_group_count['TC_BAND_passed'] = bandmatch(log_group_count['TC_BAND_passed'])#change the band to match GCF 
    workbook_name_new = str(target_path_filename).replace('Report_','Loggroup_report_')
    workbook1 = pd.ExcelWriter(workbook_name_new)
    print('Writing ' + workbook_name_new)
    log_group_count.to_excel(workbook_name_new, sheet_name= 'Reports')
"""

    

#url = 'https://mrt-reports.cloud.rsext.net/odata/ApplicationRuns?$filter=contains%28Application%2FVersion%2C%20%2723.37.1%27%29%20and%20Verdict%20eq%20%27Passed%27'
















