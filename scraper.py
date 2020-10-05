from selenium import webdriver 
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import pandas as pd 
import time 

# Specify webpage browser to use 
PATH = "PATH TO CHROME WEBDRIVER"
driver = webdriver.Chrome(executable_path=PATH) 
# Specify page url 
driver.get("https://enroll.dlsu.edu.ph/dlsu/view_actual_count")
# Specify course codes 
course_codes = ["COURSE1", "COURSE2"]
try:
    #Save the course tables into master table 
    table_df = []
    for course in course_codes: 
        # Wait until the search bar is available 
        search_field = WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.NAME, "p_course_code"))
        )

        # Send a course code query to the search bar 
        search_field.send_keys(course)
        search_field.send_keys(Keys.RETURN)

        # Wait until the table has loaded 
        table =  WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.TAG_NAME, "center"))
        )
        #Parse the table html into df 
        table_html = table.get_attribute('innerHTML')
        course_df = pd.read_html(table_html)[0]
        #Set the headers and remove whitepace
        course_df.columns = course_df.iloc[0]
        course_df.rename(columns=lambda x: x.strip())
        course_df = course_df.drop(course_df.index[0])
        #Preprocess na before appending 
        course_df.fillna(" ", inplace=True)
        if len(course_df) > 0 : 
            table_df.append(course_df)
    
    #Create a sheet for each subject 
    for index, df in enumerate(table_df):  
        name = df.iloc[0]["Course"]
        if index == 0: 
            df.to_excel("Subjects.xlsx",sheet_name=name)
        else: 
            with pd.ExcelWriter('Subjects.xlsx', mode='a') as writer:  
                df.to_excel(writer, sheet_name=name)
finally:
    driver.quit()