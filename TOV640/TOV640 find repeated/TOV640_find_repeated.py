"""
File: TOV640_find_repeated.py
Author: Lam Wai Taing, Timothy
Date: 2024/07/15
Description: A Python script to generate a TOV640 find repeated exception report of EAL or TML
from input consecutive exception reports.
"""

import pandas as pd
import time
from datetime import datetime
import tkinter as tk
from tkinter import filedialog
import os
import openpyxl
import numpy as np

def clean_case(case):
    """
    A helper function to clean up the dataframe. Renames the columns and assigns the Previous columns to the dataframe.

    Args:
        case (DataFrame): The input DataFrame to be cleaned up.

    Retruns:
        (DataFrame): The cleaned DataFrame.
    """
    case_fix = case.drop(columns=['exception type_x', 'level_x', 'startM_x', 'endM_x',
                                    'length_x', 'maxValue_x', 'maxLocation_x', 'track type_x',
                                    'key', 'startM_shift_x', 'endM_shift_x',
                                  'maxLocation_shift_x', 'startM_shift_y', 'endM_shift_y', 'maxLocation_shift_y', 'Overlap_x',
                                  'Tension Length_x', 'Landmark_x'])\
        .rename({'id_y': 'id', 'exception type_y': 'exception type', 'level_y': 'level', 'startM_y': 'startM',
                 'endM_y': 'endM', 'length_y': 'length', 'maxValue_y': 'maxValue', 'maxLocation_y': 'maxLocation',
                 'track type_y': 'track type', 'Overlap_y': 'Overlap', 'Tension Length_y': 'Tension Length',
                 'Landmark_y': 'Landmark'}, axis=1)
    if 'previous' in case_fix.columns:
        case_fix['previous'] = case_fix['previous'].astype(str) + ',' + case_fix['id_x'].astype(str)
        # case_fix['Previous 2'] = case_fix['id_x'].astype(str)
    else:
        case_fix['previous'] = case_fix['id_x'].astype(str)
        # case_fix['Previous 1'] = case_fix['id_x'].astype(str)
        # case_fix['Previous 2'] = ''

    case_fix = case_fix.reset_index().drop(columns=['index', 'id_x'])
    return case_fix[['id', 'exception type', 'level', 'startM', 'endM', 'length', 'maxValue',
                     'maxLocation', 'track type', 'Overlap', 'Tension Length', 'Landmark', 'previous']]


def find_repeated(df1, df2, df1_shift, df2_shift):
    """
    Finds any repeated exception in the same category between two DataFrames from the exception reports.

    Args:
        df1 (DataFrame): The first DataFrame.
        df2 (DataFrame): The second DataFrame.
        df1_shift ():
        df2_shift (): 

    Returns:
        (DataFrame): The DataFrame which contains the repeated exception.
    """
    if df1.empty or df2.empty:
        return pd.DataFrame(columns=['id', 'exception type', 'level', 'startM', 'endM', 'length',
                                     'maxValue', 'maxLocation', 'track type', 'Overlap', 'Tension Length', 'Landmark', 'Previous 1', 'Previous 2'])
    else:
        for all in ['startM', 'endM', 'maxLocation']:
            df1[all + '_shift'] = df1[all] + df1_shift
            df2[all + '_shift'] = df2[all] + df2_shift
        merged = df1.assign(key=1).merge(df2.assign(key=1), on='key')

        # ---------- case1 = 1st exception behind, 2nd at front ----------
        case1 = merged.query('(`startM_x`.between(`startM_y`, `endM_y`)) & '
                           '(`endM_x` > `endM_y`) &'
                           '(`maxLocation_x`.between(`startM_x`, `endM_y`)) &'
                           '(`maxLocation_y`.between(`startM_x`, `endM_y`))', engine='python')

        # ---------- case2 = 1st exception at front, 2nd behind ----------
        case2 = merged.query('(`startM_y`.between(`startM_x`, `endM_x`)) & '
                           '(`endM_x` < `endM_y`) &'
                           '(`maxLocation_x`.between(`startM_y`, `endM_x`)) &'
                           '(`maxLocation_y`.between(`startM_y`, `endM_x`))', engine='python')

        # ---------- case3 = 2nd exception covering whole 1st  ----------
        case3 = merged.query('(`startM_x` >= `startM_y`) & '
                           '(`endM_x` <= `endM_y`) & '
                           '(`maxLocation_x`.between(`startM_x`, `endM_x`)) &'
                           '(`maxLocation_y`.between(`startM_x`, `endM_x`))', engine='python')

        # ---------- case4 = 1st exception covering whole 2nd  ----------
        case4 = merged.query('(`startM_x` <= `startM_y`) & '
                           '(`endM_x` >= `endM_y`) & '
                           '(`maxLocation_x`.between(`startM_y`, `endM_y`)) &'
                           '(`maxLocation_y`.between(`startM_y`, `endM_y`))', engine='python')

        return pd.concat([clean_case(case1), clean_case(case2), clean_case(case3), clean_case(case4)])\
            .drop_duplicates(keep='first')\
            .reset_index().drop('index', axis=1)

def main():
    print('================ Finding Repeated Exception Tool ================')
    print('This is a tool to find repeated exceptions of at most 12 consecutive TOV Exception Reports')
    check = input('Enter the number of TOV exception reports to be compared? (input 2-12): ')
    loop = True
    while loop:
        if check.isdigit():
            if (int(check) >= 2) & (int(check) <= 12):
                loop = False
                break
        print('Error! Make sure you only input an integer between 2 and 12')
        check = input('Enter the number of TOV exception reports to be compared? (input 2-12): ')
    
    print('===========================================================')
    print('You are going to compare ' + check + ' TOV Exception Reports')

    # ---------- allow user to select Exception Report files -------
    root = tk.Tk()
    root.withdraw()

    df_path = []
    df_W = []
    df_LH = []
    df_HH = []
    df_SL = []
    df_SR = []

    for i in range(0, int(check)):
        if i == 0:
            print('Select 1st Exception Report after 3 seconds...')
        elif i == 1:
            print('Select 2nd Exception Report after 3 seconds...')
        elif i == 2:
            print('Select 3rd Exception Report after 3 seconds...')
        else:
            print('Select ' + str(i+1) + 'th Exception Report after 3 seconds...')        
        for j in range(3, 0, -1):
            print(f"{j}", end="\r", flush=True)
            time.sleep(1)

        path = filedialog.askopenfilename()
        df_path.append(path)
        print('Selected: ' + os.path.basename(path) + '\n')
        
        wire_wear = pd.read_excel(path, sheet_name='wear exception')
        df_W.append(wire_wear)

        low_height = pd.read_excel(path, sheet_name='low height exception')
        df_LH.append(low_height)

        high_height = pd.read_excel(path, sheet_name='high height exception')
        df_HH.append(high_height)

        stagger_left = pd.read_excel(path, sheet_name='stagger left exception')
        stagger_left = stagger_left.loc[stagger_left['Overlap'] != 'Y']
        df_SL.append(stagger_left)

        stagger_right = pd.read_excel(path, sheet_name='stagger right exception')
        stagger_right = stagger_right.loc[stagger_right['Overlap'] != 'Y']
        df_SR.append(stagger_right)

    repeated_W = df_W[0]
    repeated_LH = df_LH[0]
    repeated_HH = df_HH[0]
    repeated_SL = df_SL[0]
    repeated_SR = df_SR[0]

    for i in range(1, len(df_path)):
        repeated_W = find_repeated(repeated_W, df_W[i], 0, 0)
        repeated_LH = find_repeated(repeated_LH, df_LH[i], 0, 0)
        repeated_HH = find_repeated(repeated_HH, df_HH[i], 0, 0)
        repeated_SL = find_repeated(repeated_SL, df_SL[i], 0, 0)
        repeated_SR = find_repeated(repeated_SR, df_SR[i], 0, 0)

     # ---------- output results ----------
    print('Saving as Excel at', datetime.now())
    print('Done!')
    input('Press Enter to select save location')
    directory = filedialog.askdirectory()
    final_path = directory + '/' + os.path.basename(os.path.splitext(df_path[-1])[0]) + '_' + check + '_repeated.xlsx'

    # Get report templates from metadata
    temp = openpyxl.load_workbook('./metadata/EAL metadata.xlsx')
    sheets = temp.sheetnames
    for s in sheets:
        if s not in ['template', 'Previous template']:
            del temp[s]
    summary = temp['template']
    summary.title = 'Summary'
    previous_id = temp['Previous template']
    previous_id.title = 'Previous'
    temp.save(final_path)
    update = pd.concat([repeated_SL, repeated_SR, repeated_W, repeated_HH, repeated_LH], ignore_index=True, sort=False)

    update['support'] = np.nan
    update['empty1'] = np.nan
    update['empty2'] = np.nan
    update['empty3'] = np.nan
    update['empty4'] = np.nan
    # Stores previous ID in a separate sheet
    previous = pd.DataFrame()
    previous['ID'] = update['id']
    for i in range(int(check) - 1):
        temp = 'Previous ' + str(i + 1)
        previous[temp] = ''
    split_data = update['previous'].str.split(',', expand=True)
    if (not split_data.empty):
        split_data.columns = [f'Previous {i+1}' for i in range(int(check) - 1)]
        for col in split_data.columns:
            previous[col] = split_data[col]

    update = update[['id', 'startM', 'endM', 'length', 'exception type', 'maxValue', 'maxLocation', 'support',
                    'Tension Length', 'track type', 'level', 'empty1', 'empty2', 'empty3', 'empty4']]
    
    # Adds latest LV3 Stagger Exception to the report
    lv3_SL = df_SL[-1][df_SL[-1]['level'] == 'L3'][['id', 'startM', 'endM', 'length', 'exception type', 'maxValue', 'maxLocation', 'Tension Length', 'track type', 'level']]
    lv3_SR = df_SR[-1][df_SR[-1]['level'] == 'L3'][['id', 'startM', 'endM', 'length', 'exception type', 'maxValue', 'maxLocation', 'Tension Length', 'track type', 'level']]
    update = pd.concat([update, lv3_SL, lv3_SR], ignore_index=True, sort=False)
    update = update.drop_duplicates(subset='id', keep='first')

    with pd.ExcelWriter(final_path, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
        update.to_excel(writer, sheet_name="Summary", startrow=2, header=None, index=False)
        previous.to_excel(writer, sheet_name="Previous", startrow=1, header=None, index=False)
    # ----- save as -----

if __name__ == "__main__":
    main()
