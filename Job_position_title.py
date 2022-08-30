import pandas as pd
from datetime import date
import os
import openpyxl
from openpyxl.styles import PatternFill

""" The File needs to have duplicates removed and prepped for reading"""

# A file path on YOUR computer works best. Onedrive will block it.
date = date.today()

# Creating the upload file path. New date for each Penalty upload
if not os.path.exists(fr"C:\Users\kyle.anderson\Documents\Dan\Dan_pacs5"):
    os.makedirs(fr"C:\Users\kyle.anderson\Documents\Dan\Dan_pacs5")

# Hardcoded excel file needs to be saved here
df = pd.read_excel(fr"C:\Users\kyle.anderson\Documents\Dan\PACS5 All Employees Report 220725 final.xlsm")


df = df.drop(columns='Unique EE')
df = df.drop(columns='Last Name')
df = df.drop(columns='First Name')
df = df.drop(columns='Worker')
df = df.drop(columns='Gender')
df = df.drop(columns='Employee ID')
df = df.drop(columns='Social Security Number')
df = df.drop(columns='Employment Status')
df = df.drop(columns='Primary Position Company')
df = df.drop(columns='Alternative Worker ID')
df = df.drop(columns='Seniority Date')
df = df.drop(columns='Position Company')
df = df.drop(columns='Automation_Facility')
df = df.drop(columns='Job Family Group')
df = df.drop(columns='Position')
df = df.drop(columns='Time Type')
df = df.drop(columns='Employee Type')
df = df.drop(columns='Hire Date')
df = df.drop(columns='Job Classifications - Job Profile')
df = df.drop(columns='Time in Position')
df = df.drop(columns='Compensation Plans')
df = df.drop(columns='Pay Rate Type')
df = df.drop(columns='Pay Group')
df = df.drop(columns='Compensation Plan')
df = df.drop(columns='Amount')


building_range = df['Matched_Facility']


list_of_unique_buildings = set(building_range)

for building in list_of_unique_buildings:
    result_df = df[df['Matched_Facility'] == building]
    # df3 = df['Position here'].map(df_p['position_title'].value_counts()).fillna(0).astype(int)
    # print(result_df, " ", df3)
    result_df.to_excel(fr"C:\Users\kyle.anderson\Documents\Dan\Dan_pacs5\{building}.xlsx")


