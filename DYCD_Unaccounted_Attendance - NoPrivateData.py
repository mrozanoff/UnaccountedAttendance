import time
from datetime import datetime, timedelta, date
import pandas as pd
import traceback
import chromedriver_autoinstaller
from pathlib import Path
import os

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

#-----------------------------------FUNCTIONS

def enter_DYCD(): #login and get to reports page
    driver.get("https://www.dycdconnect.nyc/Home/Login")
    assert "DYCD" in driver.title

    #input user and pass
    username = driver.find_element(By.ID, "UserName");
    password = driver.find_element(By.ID, "Password");

    username.send_keys('user')
    password.send_keys('pass')

    #login
    buttons = driver.find_elements(By.TAG_NAME, 'button')
    buttons[1].click()

    time.sleep(t*0.5)

    #click on PTS/EMS
    pts = WebDriverWait(driver, t).until(EC.visibility_of_all_elements_located((By.CLASS_NAME, 'btn-group')))
    pts[1].click()

    #switch to the new tab
    driver.switch_to.window(driver.window_handles[1])

    #open menu and click reports
    menu = driver.find_element(By.CLASS_NAME, 'navBarTopLevelItem')
    button = menu.find_element(By.ID, 'TabMainMenu')
    button.click()

    time.sleep(t*0.1)

    nav_groups = driver.find_elements(By.CLASS_NAME, 'nav-subgroup')
    nav_groups[3].click()

    #Now at reports page! --------

    #switch to iframe within reports page
    iframe = driver.find_element(By.XPATH, '//*[@id="contentIFrame0"]')
    driver.switch_to.frame(iframe)

    time.sleep(1)

def next_page(): #click next page
    element = driver.find_element(By.XPATH, '//*[@id="_nextPageImg"]')
    driver.execute_script("arguments[0].click();", element)

def find_report(report): #while on main reports page, find report, and click on it
    #click tar
    tar = driver.find_element(By.XPATH, report) 
    a = tar.find_element(By.TAG_NAME, 'a')

    actions = ActionChains(driver)
    actions.move_to_element(a).perform()

    time.sleep(1)

    #
    driver.execute_script("arguments[0].click();", a)
    time.sleep(1)

    #Go to report builder pag
    driver.switch_to.window(driver.window_handles[2])
    driver.maximize_window()

    #switch to iframe within report window
    driver.switch_to.frame(0)

def select_element(xpath, n):
    select = Select(driver.find_element(By.XPATH, xpath))
    select.select_by_value(str(n))

def select_workscope_element(xpath, n):
    select = Select(driver.find_element(By.XPATH, xpath))
    select.select_by_value(str(n))

    return select.first_selected_option.text

def download_report(): #click view report

    driver.find_element(By.XPATH, '//*[@id="reportViewer_ctl08_ctl00"]').click()

    # element = WebDriverWait(driver, 4).until(
    #     EC.visibility_of_element_located((By.XPATH, '//*[@id="VisibleReportContentreportViewer_ctl13"]'))
    # )
    time.sleep(5)
    # element.click()
    driver.find_element(By.XPATH, '//*[@id="reportViewer_ctl09"]/div/div[5]').click()
    time.sleep(2)

    #click excel
    driver.find_element(By.XPATH, '//*[@id="reportViewer_ctl09_ctl04_ctl00_Menu"]/div[2]/a').click()  
 
def global_prev(program_area): #speed things up in TAR by keeping certain elements plugged in
    global prev 
    prev = program_area

def fill_ua(program_area, workscope): #Run thru the Unaccounted attendance report

    if prev != program_area: #do this IF not the same as previous program.
        #program area
        select_element('//*[@id="reportViewer_ctl08_ctl06_ddValue"]', program_area)

        time.sleep(1)

        #provider
        select_element('//*[@id="reportViewer_ctl08_ctl10_ddValue"]', 1)

        time.sleep(2)

    #workscope
    name = select_workscope_element('//*[@id="reportViewer_ctl08_ctl12_ddValue"]', workscope)

    time.sleep(1)

    #start date
    start = driver.find_element(By.XPATH, '//*[@id="reportViewer_ctl08_ctl20_txtValue"]')
    start.clear()
    time.sleep(1)

    start.send_keys(str(start_date))
    start.send_keys(Keys.RETURN)
    time.sleep(2)

    #end date
    select_element('//*[@id="reportViewer_ctl08_ctl22_ddValue"]', 7)

    time.sleep(3)

    #group by activity or group
    select_element('//*[@id="reportViewer_ctl08_ctl24_ddValue"]', 1)

    time.sleep(3)
    download_report()

    global_prev(program_area)
    time.sleep(5)

    return name

bronx = [' Fairmont Neighborhood School', ' P.S. 211', ' I.S. X318 Math, Science & Technology Through Arts', ' P.S. 061 Francisco Oller', ' IS 219 New Venture School -FY23']
manhattan = [' M.S. 319 - Maria Teresa', ' M.S. 324 - Patria Mirabal', ' P.S. 005 Ellen Lurie', ' P.S. 152 Dyckman Valley', ' Central Park East II', ' City College Academy of the Arts', ' Middle School 322', ' P.S. 008 Luis Belliard']
centers = [' Frederick Douglass Center', ' Dunlevy Milbank Center', ' The Lexington Academy',  ' The Dunlevy Milbank Center', ' I.S. 061 William A Morris', ' The Childrens Aid Society']
hs = [' Curtis High School']

def sort_cohort(name):
    match name:
        case name if name in bronx:
            return 'Bronx'
        case name if name in manhattan:
            return 'Manhattan'
        case name if name in centers:
            return 'Centers'
        case name if name in hs:
            return 'High Schools'
        case _:
            return 'Literacy Services'

def create_summary(workscope):
    #calculate needed numbers for ua
    cancelled = len(df[df['Appointment Status'] == 'Canceled'])

    temp_df = df.loc[df['Appointment Status'] == 'Scheduled']
    count_1 = temp_df['Unaccounted Attendance'].sum()
    count_2 = temp_df['Unaccounted Attendance'][temp_df['Unaccounted Attendance'] > 0].count()

    if section[0] == 1:
        new_name = str(name.split('***')[-1])
    elif section[0] == 3:
        new_name = ' ' + ' '.join(name.split('-')[0:2]) + str(workscope)
    else:
        new_name = str(name.split('***')[-1]) + ' - ' + str(name.split()[1].split('-')[0])

    total_workscopes.append([sort_cohort(str(name.split('***')[-1])), new_name, cancelled, count_1, count_2])

    #summary df
    summary = [['Start Date:', str(start_date).split(' ')[0]],
                  ['End Date:', str(end_date).split(' ')[0]],
                  ['Report Run On:', str(date.today())],
                  ['',''],
                  ['Cancelled Activities / Groups:', cancelled],
                  ['Unaccounted Attendance Sum*:', count_1],
                  ['Unaccounted Attendance Day & Activity Count**:', count_2]
    ]

    summary_df = pd.DataFrame(data=summary)

    #start reading into excel
    writer = pd.ExcelWriter(folder_name + new_name.replace(' ', '_')[1:] + '_' + str(start_date).split(' ')[0] + '.xlsx') 

    notes_df = pd.DataFrame(data=['* This is the total number of missing attendance entries for this time period and represents the number of data entries needed to be caught up. This does not include canceled days',
    '** This is the number of instances of missing attendance across days and activities/groups. This does not include canceled days.'])


    #enter the dfs
    summary_df.to_excel(writer, index=False, header=False)
    df.to_excel(writer, startrow=len(summary_df)+1, index=False)
    notes_df.to_excel(writer, startrow=len(summary_df)+len(df)+3, index=False, header=False)

    # Auto-adjust columns' width
    for column in df:
        column_width = max(df[column].astype(str).map(len).max(), len(column))
        col_idx = df.columns.get_loc(column)
        writer.sheets['Sheet1'].set_column(col_idx, col_idx, column_width)

    writer.save()

    return cancelled, count_1, count_2

#---------------Post Process Function

def get_UA_tab():
    downloads_path = str(Path.home() / "Downloads")

    file = str(downloads_path) + r'\Unaccounted Attendance.xlsx'

    file = r'C:\Users\mrozanoff\Downloads\Unaccounted Attendance.xlsx'

    df = pd.read_excel(file, skiprows=21)

    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]

    os.remove(file)

    return df

#-----------------------VARIABLES

t = 1

chromedriver_autoinstaller.install()  # Check if the current version of chromedriver exists
                                      # and if it doesn't exist, download it automatically,
                                      # then add chromedriver to path

options = webdriver.ChromeOptions()

start_date = input('DID YOU CHANGE DATE?')
start_date = pd.to_datetime('3/6/23')
end_date = start_date + timedelta(days=6)

compass_workscopes = [2, [2, 4, 7, 9, 11, 13, 15, 17, 19, 21, 23, 25, 27, 29, 31, 33, 34, 35, 37, 38, 40, 41, 42]]
beacon_workscopes = [1, [1]]
alit_workscopes = [3, [1, 2, 3, 4]]
# 28 Total!

workscopes = [beacon_workscopes, alit_workscopes, compass_workscopes] #28 Total

#create variables to store site information
total_cancelled = 0
total_UA_sum = 0
total_count = 0

# to store list of sites information before making it into a df
total_workscopes = []

name = 'NA' #Placeholder

#-----------------MAIN

#create date folder
folder_name = "A:/Office of Performance Management/Service Data/Python/Unaccounted Attendance Reports/" + str(start_date.date()) + ' to ' + str(end_date.date()) + '/'

Path(folder_name).mkdir(parents=True, exist_ok=True)

#Open DYCD and begin scrape
driver = webdriver.Chrome(options=options)

enter_DYCD()

next_page() #all reports are on page 2

find_report('//*[@filename="UnaccountedForAttendance.rdl"]')

global_prev(0)

for section in workscopes:
    for workscope in section[1]:
        for i in range(3):
            try:
                time.sleep(2)

                name = fill_ua(section[0], workscope)

                #get the tab
                df = get_UA_tab()

                df = df.ffill()
                
                cancelled, count_1, count_2 = create_summary(workscope)

                #add calulcations to yd totals
                total_cancelled += cancelled
                total_UA_sum += count_1
                total_count += count_2

            except Exception as e:
                print(e)
                print(name, 'Not Working')
                time.sleep(3)
            else:
                break


# creating yd totals summary
YD_summary = [['Start Date:', str(start_date).split(' ')[0]],
              ['End Date:', str(end_date).split(' ')[0]],
              ['Report Run On:', str(date.today())],
              ['',''],
              ['YD Summary:', ''],
              ['Total Cancelled:', total_cancelled],
              ['Total Unaccounted Attendance Sum*:', total_UA_sum],
              ['Total Unaccounted Attendance Day & Activity Count**:', total_count],
]

YD_summary_df = pd.DataFrame(data=YD_summary, columns=['col1', 'col2'])

#build dataframe with all school details
total_df = pd.DataFrame(data=total_workscopes, columns=['Cohort', 'Site', 'Cancelled', 'Unaccounted Attendance Sum', 'Unnacounted Attendance Day & Activity Count'])
total_df['# of Workscopes with Missing Attendance Data'] = total_df.apply(lambda row: 1 if row['Unaccounted Attendance Sum'] > 0 else 0 , axis=1)
total_df = total_df.sort_values(['Cohort', 'Site'])

# create cohort summary df
cohorts_df = total_df['# of Workscopes with Missing Attendance Data'].gt(0).astype(int).groupby(total_df['Cohort']).sum()

# no longer need this col, used above
total_df = total_df.drop(['# of Workscopes with Missing Attendance Data'], axis=1)

# Create YD summary excel
writer = pd.ExcelWriter(folder_name + 'UA_YDsummary_' + str(start_date).split(' ')[0] + '.xlsx')

notes_df = pd.DataFrame(data=['* This is the total number of missing attendance entries for this time period and represents the number of data entries needed to be caught up. This does not include canceled days',
'** This is the number of instances of missing attendance across days and activities/groups. This does not include canceled days.'])

#write in each df to excel writer
YD_summary_df.to_excel(writer, index=False, header=False)
cohorts_df.to_excel(writer, startrow=len(YD_summary_df)+1)
total_df.to_excel(writer, startrow=len(YD_summary_df)+len(cohorts_df)+3, index=False)
notes_df.to_excel(writer, startrow=len(YD_summary_df)+len(cohorts_df)+len(total_df)+5, index=False, header=False)

# Auto-adjust columns' width for yd sumary df
for column in YD_summary_df:
    column_width = max(YD_summary_df[column].astype(str).map(len).max(), len(column))
    col_idx = YD_summary_df.columns.get_loc(column)
    writer.sheets['Sheet1'].set_column(col_idx, col_idx, column_width)

# Auto-adjust columns' width for total df
for column in total_df:
    column_width = max(total_df[column].astype(str).map(len).max(), len(column))
    col_idx = total_df.columns.get_loc(column)
    writer.sheets['Sheet1'].set_column(col_idx, col_idx, column_width)

writer.save()

driver.quit()