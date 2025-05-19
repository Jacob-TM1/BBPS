"""
This process generates copies of the 'BBPS-Outlet_YYYYMM.xlsm'  which is found in the ...Dashboards\Live folder of the current months publication. Make sure that the 'YYYYMM' is for the current year and month. 

"""
from calendar import month
import win32com.client as win32
import shutil
from TM1py.Services import TM1Service
from TM1py.Objects import Subset
import configparser
import time
import concurrent.futures 
from tqdm import tqdm
import os 

start = time.time()
# == VARIABLES 
config = configparser.ConfigParser()
config.read(".\config.ini")
publication_path = config.get('Publication', 'publication_folder')
publication_year = config.get('Publication', 'publication_year')
month_number =  config.get('Publication', 'month_number')
current_month_folder = publication_year + month_number 

def produce_dashboards(element, source, destination):
    element2 = element.replace('!', '/')
    shutil.copy(source, destination)

    # -- Change Cell Value in Excel Dashboard
    workbook = excel.Workbooks.Open(destination)
   
    sheet = workbook.Worksheets("Master")
    cells = sheet.Cells
    cells(21,'R').Value = element2

    #excel.CalculateUntilAsyncQueriesDone()
    workbook.Close(True)
    
def convert_to_preferred_format(sec):
   sec = sec % (24 * 3600)
   hour = sec // 3600
   sec %= 3600
   min = sec // 60
   sec %= 60
   return "%02d:%02d:%02d" % (hour, min, sec) 

    # --Excel
excel = win32.Dispatch("Excel.Application")
#excel.Visible = False
excel.ScreenUpdating = False
excel.DisplayAlerts = False
excel.EnableEvents = False

    # --TM1 Login
server_name = 'ABSA Bank'
dimension_name = 'BBPS-Site'

dashboards_and_subsets = {
    "BBPS-Super Region" : ("MD.Super Regions", "2.Super Regions"),
    # "BBPS-Regional Executive" : ("MD.Regional Executives","3.Regional Executives"),
    # "BBPS-Regional Managers" : ("MD.Regional Managers","4.Regional Managers"),
    # "BBPS-Outlet" : ("ActiveSites","6.BBPS Outlets"), 
    #  "BBPS-Outlet" : ("Temp","Temp"), 
}

with TM1Service(**config['BBPS_Development']) as tm1_dev:
    print(f'{server_name}: Development server logged in successfully')

    
    # -- Get all elements in Subset and then create a list of elements in that subset
    for dashboard, subset in dashboards_and_subsets.items():
        all_elements_in_subset = tm1_dev.dimensions.subsets.get_element_names(dimension_name=dimension_name, hierarchy_name=dimension_name, subset_name=subset[0])
        all_elements_branch_name_attribute = tm1_dev.elements.get_attribute_of_elements(dimension_name, dimension_name, 'BranchName', all_elements_in_subset, False)
        print(dashboard, ': ',len(all_elements_branch_name_attribute))
        
        for element, attribute in tqdm(all_elements_branch_name_attribute.items()):
            attribute = attribute.replace('/','!')
            source = f'{publication_path}\{publication_year}\{current_month_folder}\Dashboards\Live\BBPS-Outlet_{publication_year}{month_number}.xlsm'
            destination = f'{publication_path}\{publication_year}\{current_month_folder}\Dashboards\Live\{subset[1]}\{dashboard}_{publication_year}{month_number}_{attribute}.xlsm'
            #if(attribute == True): # Replace "True" with name of site to produce only one site
            if os.path.exists(destination):
                print(f"File already exists, skipping: {destination}")
                continue
            try:
                produce_dashboards(attribute, source, destination)
            except AttributeError:
                continue
 
end = time.time()
print("Completed in: ", convert_to_preferred_format(end-start))
#==========================
