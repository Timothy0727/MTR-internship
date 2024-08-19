"""
File: trend.py
Author: Lam Wai Taing, Timothy
Date: 2024/08/16
Description: A Python script to generate the trend report of Previous 6 months.
"""

import os
import re
import numpy as np
import pandas as pd
import openpyxl
import time
import tkinter as tk
from datetime import datetime
from tkinter import filedialog

print('Before you start, please go through the following instructions.')
print('1. Please prepare the Find Repeated Report of 3 Exception Reports.')
print('2. Please prepare the 5 previous Exception Reports.')
print('3. Please prepare the Catenary Report of previous 3, 4, and 5 months.\n')
input('Press Enter to continue...')
print()

# inputs and reads find repeated report
try:
    print('Please select the Find Repeated Report after 1 second...')
    time.sleep(1)
    root = tk.Tk()
    root.withdraw()
    path = filedialog.askopenfilename()
    print('Selected: ' + os.path.basename(path) + '\n')

    wire_L2 = pd.read_excel(path, sheet_name='Summary')
    wire_L2.columns = wire_L2.iloc[0]
    wire_L2 = wire_L2[1:]
    wire_L2 = wire_L2[(wire_L2['Exception Type'] == 'Wire Wear') &
                    (wire_L2['Level'] == 'L2') &
                    ((wire_L2['ACTION'] == '') | (pd.isna(wire_L2['ACTION'])) | (wire_L2['ACTION'] == 'NaN'))]
    wire_L2 = wire_L2[['ID', 'StartM', 'EndM', 'MaxValue', 'MaxLocation', 'Level']]

    # reads previous 2 IDs
    previous_2 = pd.read_excel(path, sheet_name='Previous')
    previous_2 = previous_2[previous_2['ID'].astype(str).isin(wire_L2['ID'].astype(str))]
    previous_2 = previous_2[['ID', 'Previous 1', 'Previous 2']]

    wire_L2 = wire_L2.merge(right=previous_2, on='ID').reset_index(drop=True)
except KeyError as err:
    print('ERROR: Invalid or missing column names in ' + path)
    input('Press Enter to exit')
    exit()
except Exception as err:
    print('ERROR: ' + str(err) + ' in ' + path)
    input('Press Enter to exit')
    exit()

# copies and pastes the maxValue of current date
current_date = re.search(r'(\d{8})', path)
if current_date:
    current_date = current_date.group(1)
    current_date = datetime.strptime(current_date, '%Y%m%d').strftime('%d-%m-%Y')
else:
    print('ERROR: No date or invalid date found in the file name ' + path)
    print('Please make sure file name has valid date format of year-month-date, for example, 20240813')
    input('Press Enter to exit')
    exit()
wire_L2[current_date] = wire_L2['MaxValue']

dates = pd.DataFrame()
dates['date'] = ''
dates['days'] = ''
dates.loc[0, 'date'] = current_date
dates.loc[0, 'days'] = 0

database_df = pd.DataFrame()
catenary_df = pd.DataFrame()

try:
    # inputs and reads Previous 2 wire wear Exception Rerpots
    print('Please select the Previous 2 Exception Reports')
    print('Selecting first Previous Exception (' + re.search(r'(.*?_W)', wire_L2['Previous 2'].iloc[0]).group(1)[:-2] + '): ')
    time.sleep(1)
    previous_1_path = filedialog.askopenfilename()
    print('Selected: ' + os.path.basename(previous_1_path) + '\n')
    previous_1st_df = pd.read_excel(previous_1_path, sheet_name='wear exception')
    date_1st_str = wire_L2['Previous 2'].iloc[0][:8]

    # checks if date in file name matches with date in Previous 2 ID
    if (os.path.basename(previous_1_path)[:8] != date_1st_str):
        print("ERROR: date in file name " + os.path.basename(previous_1_path) + f" doesn't match with date in Previous 2 " + date_1st_str)
        input('Press Enter to exit')
        exit()

    date_1st = datetime.strptime(date_1st_str, '%Y%m%d')
    date_1st = date_1st.strftime('%d-%m-%Y')
    dates.loc[1, 'date'] = date_1st

    database_df = previous_1st_df
except Exception as err:
    print('ERROR: ' + str(err) + ' in ' + previous_1_path)
    input('Press Enter to exit')
    exit()

try:
    print('Selecting second Previous Exception (' + re.search(r'(.*?_W)', wire_L2['Previous 1'].iloc[0]).group(1)[:-2] + '): ')
    time.sleep(1)
    previous_2_path = filedialog.askopenfilename()
    print('Selected: ' + os.path.basename(previous_2_path) + '\n')
    previous_2nd_df = pd.read_excel(previous_2_path, sheet_name='wear exception')
    date_2nd_str = wire_L2['Previous 1'].iloc[0][:8]

    # checks if date in file name matches with date in Previous 1 ID
    if (os.path.basename(previous_2_path)[:8] != date_2nd_str):
        print("ERROR: date in file name " + os.path.basename(previous_2_path) + f" doesn't match with date in Previous 1 " + date_2nd_str)
        input('Press Enter to exit')
        exit()

    date_2nd = datetime.strptime(date_2nd_str, '%Y%m%d')
    date_2nd = date_2nd.strftime('%d-%m-%Y')
    dates.loc[2, 'date'] = date_2nd

    database_df = pd.concat([database_df, previous_2nd_df])
except Exception as err:
    print('ERROR: ' + str(err) + ' in ' + previous_2_path)
    input('Press Enter to exit')
    exit()

wire_L2[date_1st] = ''
wire_L2[date_2nd] = ''

# looks up 2 previous wear values from previous 2 Exception Reports
for index, row in wire_L2.iterrows():
    if previous_1st_df[previous_1st_df['id'] == row['Previous 2']]['maxValue'].empty:
        print('ERROR: ID is not found in first previous Exception Report ' + previous_1_path)
        input('Press Enter to exit')
        exit()
    elif previous_2nd_df[previous_2nd_df['id'] == row['Previous 1']]['maxValue'].empty:
        print('ERROR: ID is not found in second previous Exception Report ' + previous_2_path)
        input('Press Enter to exit')
        exit()
    previous_1st = previous_1st_df[previous_1st_df['id'] == row['Previous 2']]['maxValue'].values[0]
    previous_2nd = previous_2nd_df[previous_2nd_df['id'] == row['Previous 1']]['maxValue'].values[0]
    wire_L2.loc[index, date_1st] = previous_1st
    wire_L2.loc[index, date_2nd] = previous_2nd

    # inputs previous 3, 4, 5 Exception Reports
    # if not exception values not found, looks up in Catenary Reports
try:
    for i in range(3, 6):
        print('Selecting Previous ' + str(i) + ' Exception Report')
        time.sleep(1)
        temp_previous_exception_path = filedialog.askopenfilename()
        current_path = temp_previous_exception_path     # for debugging
        temp_previous_exception_df = pd.read_excel(temp_previous_exception_path, sheet_name='wear exception')
        print('Selected: ' + os.path.basename(temp_previous_exception_path) + '\n')
        temp_date_str = temp_previous_exception_df['id'].iloc[0][:8]
        temp_date = datetime.strptime(temp_date_str, '%Y%m%d')
        temp_date = temp_date.strftime('%d-%m-%Y')
        dates.loc[i, 'date'] = temp_date

        database_df = pd.concat([database_df, temp_previous_exception_df])

        # inputs Catenary Report
        print('Selecting Catenary Report of ' + os.path.basename(temp_previous_exception_path).split('_Exception Report')[0])
        time.sleep(1)
        temp_previous_catenary_path = filedialog.askopenfilename()
        current_path = temp_previous_catenary_path      # for debugging
        temp_previous_catenary_df = pd.read_excel(temp_previous_catenary_path, sheet_name='Sheet1')
        print('Selected: ' + os.path.basename(temp_previous_catenary_path) + '\n')
        temp_previous_catenary_df.columns = temp_previous_catenary_df.iloc[1]
        temp_previous_catenary_df = temp_previous_catenary_df[2:].dropna(how='all')
        temp_previous_catenary_df = temp_previous_catenary_df.drop(temp_previous_catenary_df[temp_previous_catenary_df['LINE'].isin(['Date', 'LINE'])].index).reset_index(drop=True)
        
        temp_previous_catenary_df = temp_previous_catenary_df.rename({'RWH1mm': 'WireWear1', 'RWH2mm': 'WireWear2', 'RWH3mm': 'WireWear3', 'RWH4mm': 'WireWear4'}, axis=1)
        catenary_df = pd.concat([catenary_df, temp_previous_catenary_df[['LINE', 'TRACK', 'CHAINAGE', 'WireWear1', 'WireWear2', 'WireWear3', 'WireWear4']]], axis=1)
        catenary_df.insert(catenary_df.shape[1], ' ', '', allow_duplicates=True)
        wire_L2[temp_date] = ''
        
        for index, row in wire_L2.iterrows():
            current_path = temp_previous_exception_path     # for degbugging
            temp_maxLocation = row['MaxLocation']
            
            temp_found_maxLocation = temp_previous_exception_df[(temp_previous_exception_df['startM'].astype(int) <= row['EndM']) &
                                                                (temp_previous_exception_df['endM'].astype(int) >= row['StartM'])]
            
            # temp_found_maxLocation = temp_previous_exception_df[(((temp_previous_exception_df['startM'].astype(int) - 0.5) <= temp_maxLocation) &
            #                                                     ((temp_previous_exception_df['endM'].astype(int) + 0.49) >= temp_maxLocation))]

            # if not found in Exception Report, looks up in Catenary Report
            if (not temp_found_maxLocation.empty):
                wire_L2.loc[index, temp_date] = np.max(temp_found_maxLocation['maxValue'].values)
            else:
                current_path = temp_previous_catenary_path  # for debugging
                temp_found_maxLocation = temp_previous_catenary_df[temp_previous_catenary_df['CHAINAGE'].astype(float) == temp_maxLocation]
                temp_maxValue = np.max(temp_found_maxLocation[['WireWear1', 'WireWear2', 'WireWear3', 'WireWear4']])
                wire_L2.loc[index, temp_date] = temp_maxValue

    database_df.insert(0, 'month', '')
except Exception as err:
    print('ERROR: ' + str(err) + ' in ' + current_path)
    input('Press Enter to exit')
    exit()

try:
    # Calculates days
    for index, row in dates.iterrows():
        if index == 0:
            continue
        else:
            dates.loc[index, 'days'] = (datetime.strptime(current_date, '%d-%m-%Y') - datetime.strptime(row['date'], '%d-%m-%Y')).days

    dates = dates.transpose().reset_index(drop=True)

    wire_L2[['trd pt 1', 'trd pt 2', 'trd pt 3', 'trd pt 4', 'trd pt 5', 'trd pt 6', 'logic 1', 'logic 2', 'Result']] = ''

    for index, row in wire_L2.iterrows():
        row = row.to_frame().transpose()

        trend_df = dates.reset_index(drop=True)
        trend_df.columns = trend_df.iloc[0].values
        trend_df = trend_df.reset_index(drop=True)
        trend_df = trend_df.iloc[1:].reset_index(drop=True)

        date_columns = [col for col in row.columns if col in trend_df.columns]

        trend_df = pd.concat([trend_df, row[date_columns]]).reset_index(drop=True)

        # calculates coefficient
        z = np.polyfit(trend_df.iloc[0].astype(int) ,trend_df.iloc[1].astype(float), 1)

        # calculates trend points
        for i in range(0, 6):
            column_name = 'trd pt ' + str(i + 1)
            wire_L2.loc[index, column_name] = z[0] * trend_df.iloc[0, i] + z[1]

        # assigns logics
        wire_L2.loc[index, 'logic 1'] = (wire_L2.loc[index, 'trd pt 1'] <= 10.2)
        wire_L2.loc[index, 'logic 2'] = (abs(wire_L2.loc[index, 'MaxValue'] - wire_L2.loc[index, 'trd pt 1']) > 0.2)
        if (~wire_L2.loc[index, 'logic 1']):
            wire_L2.loc[index, 'Result'] = 'no action required'
        elif (wire_L2.loc[index, 'logic 1']) & (wire_L2.loc[index, 'logic 2']):
            wire_L2.loc[index, 'Result'] = 'verify on site'
        elif ((wire_L2.loc[index, 'logic 1']) & (~wire_L2.loc[index, 'logic 2'])):
            wire_L2.loc[index, 'Result'] = 'confirmed valid L2'

    wire_L2['ID_num'] = wire_L2['ID'].str.extract(r'_W(\d+)').astype(int)
    wire_L2 = wire_L2.sort_values(by='ID_num').drop(columns='ID_num')
except Exception as err:
    print('ERROR: ' + str(err))
    input('Press Enter to exit')
    exit()

# ---------- Saves report at chosen directory ----------
try:    
    print('Saving as Excel at', datetime.now())
    print('Done!')
    input('Press Enter to select save location')
    directory = filedialog.askdirectory()
    final_path = directory + '/' + current_date + ' Wire Wear L2 Trend Report.xlsx'
    print('Saving at ' + directory)

    # copies and writes the template before writing the report
    # writes the months, date and days
    wb = openpyxl.load_workbook('./Wire Wear L2 Trend Report Template.xlsx')
    ws = wb['template']
    ws.title = 'Summary'

    for row_index, row in dates.iterrows():
        for col_index, value in enumerate(row):
            # print(value)
            ws.cell(row=row_index + 2, column=9 + col_index, value=value)

    wb.save(final_path)

    with pd.ExcelWriter(final_path, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
        wire_L2.to_excel(writer, sheet_name='Summary', startrow=3, header=None, index=False)
        database_df.to_excel(writer, sheet_name='Database', index=False)
        catenary_df.to_excel(writer, sheet_name='Catenary', startrow=1, index=False)
except FileNotFoundError:
    print(f"ERROR: file 'Wire Wear L2 Trend Report Template.xlsx' is not found in the current directory")
    input('Press Enter to exit')
    exit()
except KeyError:
    print(f"ERROR: worksheet 'template' is not found in 'Wire Wear L2 Trend Report Template.xlsx'")
    input('Press Enter to exit')
    exit()
except Exception as err:
    print(err)
    input('Press Enter to exit')
    exit()
# ---------- Saves report at chosen directory ----------
