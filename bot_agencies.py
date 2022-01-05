# Libraries -----------------------------------------------------------------------------------------------------------------------
from RPA.Browser.Selenium import Selenium
from RPA.Browser.Selenium import webdriver
from RPA.Excel.Files import Files
import time
import xlrd
from xlwt import Workbook
from xlutils.copy import copy as xl_copy
import os
script_path = os.path.abspath(__file__) # i.e. /path/to/dir/foobar.py
script_dir = os.path.split(script_path)[0] #i.e. /path/to/dir/
rel_path = 'agency.conf'
abs_file_path = os.path.join(script_dir, rel_path)
exc_rel_path = './output/file_agencies.xls'
exc_file = os.path.join(script_dir, exc_rel_path)

download_path = script_path.replace("bot_agencies.py", "output/")

browser_lib = Selenium()

# Functions -----------------------------------------------------------------------------------------------------------------------
def get_agencies_spendings():    
    # Open website
    browser_lib.open_available_browser("https://itdashboard.gov/")


    print("Click on button DIVE IN - Start")
    # Click by Xpath
    dive_in_xpath = '//*[@id="node-23"]/div/div/div/div/div/div/div/a'
    try:
        browser_lib.driver.find_element_by_xpath(dive_in_xpath).click()
    except:
        browser_lib.close_browser()
        browser_lib.open_available_browser("https://itdashboard.gov/")
        browser_lib.driver.find_element_by_xpath(dive_in_xpath).click()
    print("Click on button DIVE IN - Stop")
    time.sleep(2)

    agencies_data_with_spendings = browser_lib.driver.find_elements_by_xpath('//div/div/div/div/a/span')
    # print(F"# Agencies = {str(int(len(agencies_data_with_spendings) / 2))}")
    if int(len(agencies_data_with_spendings) / 2) > 26:
        agencies_data_with_spendings = agencies_data_with_spendings[0:52]
    names = []
    spendings = []
    for x in range (0, len(agencies_data_with_spendings), 2):
        names.append(agencies_data_with_spendings[x].text)
    for x in range (1, len(agencies_data_with_spendings), 2):
        spendings.append(agencies_data_with_spendings[x].text)
    if len(names) > 26:
        names = list(set(names))
        spendings= spendings[0:len(names)]
    write_agencies_spendings_to_excel(names, spendings)
    return names, spendings

def get_individual_investments(names):
    with open(abs_file_path) as f:
        agency_conf = f.read()
    print(F"Agency in conf file: {agency_conf}")

    path_agency = F'//*[@id="agency-tiles-widget"]/div/div/div/div/div/div/div[1]/a/span[contains(text(),"{agency_conf}")]'
    agency_from_conf = browser_lib.driver.find_element_by_xpath(path_agency).click()
    time.sleep(5)
    
    path_table = '//*[@id="investments-table-widget"]'
    print("Wait until table appears - Start")
    browser_lib.wait_until_element_is_visible(path_table, timeout=15)
    print("Wait until table appears - Stop")

    path_filter_all = '//*[@id="investments-table-object_length"]/label/select/option[4]'
    

    select_all = browser_lib.driver.find_element_by_xpath(path_filter_all).click()
    # browser_lib.wait_until_element_contains('//*[@id="investments-table-object_next"]', 'disabled')
    # select_all.select_by_visible_text("All")
    time.sleep(15)


    table_text = []

    path_td = '//*[@id="investments-table-object"]/tbody/tr/td'
    table_elements_td = browser_lib.driver.find_elements_by_xpath(path_td)
    for x in table_elements_td:
        table_text.append(x.text)
        
    # Aggrupate by 7
    table_text = [table_text[n:n+7] for n in range(0, len(table_text), 7)]
    write_agency_inv_to_excel(table_text, agency_conf)

    path_td_a = '//*[@id="investments-table-object"]/tbody/tr/td/a'
    table_elements_td_a = browser_lib.driver.find_elements_by_xpath(path_td_a)

    links_pdfs = []
    for element in table_elements_td_a:
        links_pdfs.append(element.get_attribute("href"))
    
    print(F"# of UII PDFs = {str(len(links_pdfs))}")
    download_pdfs(links_pdfs)

def write_agencies_spendings_to_excel(names, spendings):
    print("Writing Excel - Start")
    # app = Files()
    # time.sleep(1)
    # try:
    #     app.open_workbook('file_agencies.xlsx')
    # except:
    #     app.create_workbook('file_agencies.xlsx')
    #     app.open_workbook('file_agencies.xlsx')
    # app.set_active_worksheet(sheetname='Agencies')
    # app.write_to_cells(row=0, column=0, value='Agency')
    # app.write_to_cells(row=0, column=1, value='Spendings')
    # for index, (agency, spending) in enumerate(zip(names, spendings)):
    #     app.write_to_cells(row=index+1, column=0, value=agency)
    #     app.write_to_cells(row=index+1, column=1, value=spending)
    # app.save_excel()
    # app.quit_application()

    # Ubuntu
    wb = Workbook()
    sheet1 = wb.add_sheet('Agencies')
    sheet1.write(0, 0, 'Agency')
    sheet1.write(0, 1, 'Spendings')
    for index, (agency, spending) in enumerate(zip(names, spendings)):
        sheet1.write(index+1, 0, agency)
        sheet1.write(index+1, 1, spending) 
    wb.save(exc_rel_path)
    print("Writing Excel - Stop")

def write_agency_inv_to_excel(table_text, agency):
    print("Writing Excel - Start")
    # Ubuntu
    rb = xlrd.open_workbook(exc_rel_path, formatting_info=True)
    wb = xl_copy(rb)
    sheet1 = wb.add_sheet(agency)
    sheet1.write(0, 0, 'UII')
    sheet1.write(0, 1, 'Bureau')
    sheet1.write(0, 2, 'Investment Title')
    sheet1.write(0, 3, 'Total FY2021 Spending ($M)')
    sheet1.write(0, 4, 'Type')
    sheet1.write(0, 5, 'CIO Rating')
    sheet1.write(0, 6, '# of Projects')
    for index, content in enumerate(table_text):
        sheet1.write(index+1, 0, content[0])
        sheet1.write(index+1, 1, content[1])
        sheet1.write(index+1, 2, content[2])
        sheet1.write(index+1, 3, content[3])
        sheet1.write(index+1, 4, content[4])
        sheet1.write(index+1, 5, content[5])
        sheet1.write(index+1, 6, content[6])
    wb.save(exc_rel_path)
    print("Writing Excel - Stop")

def download_pdfs(links):
    for index, link in enumerate(links):
        # Open website
        prefs = {
            'download.default_directory' : download_path, 
            'directory_upgrade': True, 
            'profile.default_content_settings.popups': 0,
            'profile.default_content_setting_values.automatic_downloads': 1
        }
        browser_lib.open_available_browser(link, preferences=prefs)
        time.sleep(2)
        print(F"Download PDF #{str(index + 1)} - Start")
        path_download_pdf = '//*[@id="business-case-pdf"]/a'
        try:
            browser_lib.driver.find_element_by_xpath(path_download_pdf).click()
            browser_lib.wait_until_element_is_not_visible('//*[@id="business-case-pdf"]/span', timeout=10)
        except:
            browser_lib.close_browser()
            browser_lib.open_available_browser(link, preferences=prefs)
            browser_lib.driver.find_element_by_xpath(path_download_pdf).click()
            browser_lib.wait_until_element_is_not_visible('//*[@id="business-case-pdf"]/span', timeout=10)
        download_wait(download_path, 10)
        # time.sleep(1)
        print(F"Download PDF #{str(index + 1)} - Stop")
        browser_lib.close_browser()

def download_wait(directory, timeout, nfiles=None):
    seconds = 0
    dl_wait = True
    while dl_wait and seconds < timeout:
        time.sleep(1)
        dl_wait = False
        files = os.listdir(directory)
        if nfiles and len(files) != nfiles:
            dl_wait = True

        for fname in files:
            if fname.endswith('.crdownload'):
                dl_wait = True
        seconds += 1
    return seconds


def run():
    try:
        names, spendings = get_agencies_spendings()
        time.sleep(1)
        get_individual_investments(names)
    finally:
        browser_lib.close_all_browsers()

# Main ---------------------------------------------------------------------------------------------------------------------------
if __name__ == "__main__":
    run()
    time.sleep(5)
    