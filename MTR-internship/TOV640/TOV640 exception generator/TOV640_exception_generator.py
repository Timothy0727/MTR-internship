"""
File: TOV640_exception_generator.py
Author: Lam Wai Taing, Timothy
Date: 2024/07/11
Description: A Python script to generate a TOV640 exception report of EAL or TML from an input .datac file.
"""

import traceback
import numpy as np
import pandas as pd
import os
import re
from datetime import datetime
import time
import tkinter as tk
from tkinter import filedialog
pd.options.mode.chained_assignment = None

# value of chain length to close consecutive rows
chain_length = 0.002


def file_input():
    """
    Prints out the instructions of using this program and prompts to 
    ask the user to input the details and the data of the line to be analyzed.

    Args:
        None

    Returns:
        line (str): The name of the line.
        section (str): The section of the line.
        track (str): The track of the line.
        raw (DataFrame): The data from input .datac file.
    """
    print('================ TOV640 Exception Report Generator ================')
    print('Before you start, please go through the following instructions' + 
          'or otherwise you might experience errors and the results could be inaccurate!')
    input("Press Enter to continue...")
    print('0. Make sure all metadata.xlsx of each line are in the metadata folder.')
    print('If not, please move them together and restart the program.')
    input("Press Enter to continue...")
    print('When the program is running, MAKE SURE YOU:')
    print('1. Input the name of the Line, e.g. EAL/TML...')
    print('2. Input the section of the Line, e.g. LMC or TUM-HUH')
    print('3. Input the direction of track, i.e. UP/DN')
    print('4. Select the .DATAC file of the TOV data to generate exception report')
    input("Press Enter to continue...")
    print('=========================== IMPORTANT!!! ===========================')
    input('Make sure you read through the above instructions carefully and press enter to start... ')

    line = input('Line: ')
    while line not in ['EAL', 'TML']:
        print('Please make sure you input correct line:')
        line = input('Line: ')

    if line == 'TML':
        section = input('Please input the section range, e.g. TUM-HUH: ')
        while not re.match('^[A-Z]{3}-[A-Z]{3}', section):
            print('Please make sure you input correct section range, e.g. TUM-HUH')
            section = input('Please input the section range: ')

    if line == 'EAL':
        section = input('LMC? [y/n]: ')
        while section not in ['y', 'n']:
            section = input('LMC? [y/n]: ')
        if section == 'y':
            section = 'LMC'
        else:
            section = input('RAC? [y/n]: ')
            while section not in ['y', 'n']:
                section = input('RAC? [y/n]: ')
            if section == 'y':
                section = 'RAC'
            else:
                section = input('LOW S1? [y/n]: ')
                while section not in ['y', 'n']:
                    section = input('LOW S1? [y/n]: ')
                if section == 'y':
                    section = 'LOW'
                    track = 'S1'
                else:
                    section = input('Please input the section range, e.g. UNI-TAP: ')
                    while not re.match('^[A-Z]{3}-[A-Z]{3}', section):
                        print('Please make sure you input correct section range, e.g. UNI-TAP')
                        section = input('Please input the section range: ')

    if section != 'LOW':
        track = input('UP/DN: ')
        while track not in ['DN', 'UP']:
            print('Please make sure you input either UP / DN')
            track = input('UP/DN: ')
    else:
        pass

    print('select the data report in .DATAC format after 1 second...')
    for i in range(1, 0, -1):
        print(f"{i}", end="\r", flush=True)
        time.sleep(1)
    # ---------- allow user to select csv files -------
    root = tk.Tk()
    root.withdraw()
    raw_data_path = filedialog.askopenfilename()
    print("raw:", raw_data_path)
    raw = pd.read_csv(raw_data_path, sep=';', engine='python').apply(lambda x: x.str.strip() if x.dtype == "object" else x)
    print('Selected: ' + os.path.basename(raw_data_path))
    # ---------- allow user to select csv files -------

    return line, section, track, raw


def output_date(raw):
    """
    Reads the date from the .datac file and returns it.

    Args:
        raw (DataFrame): The data from input .datac file.

    Returns:
        date (str): The date of the data was collected.
    """
    for i in range(1, raw.shape[0]):
        d = raw['Date'][i]
        if re.match(r'^\d{2}\.\d{2}\.\d{4}$', d):
            break
    day, month, year = map(int, d.split('.'))
    date = f'{year:04d}{month:02d}{day:02d}'
    return date


def m_to_km(in_table, column1, column2):
    """
    Converts the unit of location from meter to km.

    Args:
        in_table (DataFrame): The table to be converted.
        column1 (str): The name of the column to be converted.
        column2 (str): The name of the column to be converted.

    Returns:
        in_table (DataFrame): The converted table.
    """
    in_table[column1] = in_table[column1] / 1000
    in_table[column2] = in_table[column2] / 1000
    return in_table


def load_metadata(line, section, track):
    """
    Loads metadata from the corresponding .xlsx file and returns four lookup tables of the line/ section.

    If the line is TML, reads the lookup table of the corresponding track in "./metadata/TML metadata.xlsx".
    Converts the units of the values from M to KM since the data in the .datac is in KM.

    If the line is EAL, identifies if the section is either LMC, RAC, or LOW.
    Then, reads the corresponding lookup table in "./metadata/EAL metadata.xlsx".

    Args:
        line (str): The name of the line.
        section (str): The section of the line.
        track (str): The track of the line.

    Returns:
        track_type (DataFrame): The track type at different locations of the line.
        overlap (DataFrame): The overlap and tension length of the line.
        landmark (DataFrame): The landmark of the line.
        threshold (DataFrame): The threshold of the line.
    """
    # Reads the corresponding metadata according to the line and the section
    # Assumes the metadata files are stored in the metadata folder
    # ./metadata/TML metadata.xlsx or ./metadata/EAL metadata.xlsx
    try:
        if line == 'TML':
            sheetName = 'TML ' + track
        elif section in ['LMC', 'RAC', 'LOW']:
            if section == 'LOW':
                sheetName = 'LOW S1'
            else:
                sheetName = section + ' ' + track
        else:
            sheetName = 'EAL ' + track

        metadata = pd.read_excel('./metadata/' + line + ' metadata.xlsx', sheet_name=sheetName)
         # splits the combined lookup table into three tables
        track_type = metadata[['track type', 'Track Type startM', 'Track Type endM']]\
            .dropna(how='all').rename({'Track Type startM': 'startKM', 'Track Type endM': 'endKM'}, axis=1)
        overlap = metadata[['Overlap FromM', 'Overlap ToM', 'Overlap', 'Tension Length']]\
            .dropna(how='all').rename({'Overlap FromM': 'FromKM', 'Overlap ToM': 'ToKM'}, axis=1)
        landmark = metadata[['Landmark FromM', 'Landmark ToM', 'Landmark']]\
            .dropna(how='all').rename({'Landmark FromM': 'FromKM', 'Landmark ToM': 'ToKM'}, axis=1)

        # Converts from m to km to match the units with .datac
        track_type = m_to_km(track_type, 'startKM', 'endKM')
        overlap = m_to_km(overlap, 'FromKM', 'ToKM')
        landmark = m_to_km(landmark, 'FromKM', 'ToKM')
        threshold = pd.read_excel('./metadata/' + line + ' metadata.xlsx', sheet_name='threshold')
        overlap['Overlap'] = overlap['Overlap'].fillna('N')
        print('Loading...')
        return track_type, overlap, landmark, threshold
    except FileNotFoundError as err:
        print(err)
        print('Please make sure the metadata.xlsx file is stored in the metadata folder, and the metadata folder is located at the same directory as this .exe program.')
        input('Press Enter to exit')
        exit()
    except Exception as err:
        print(traceback.format_exc())
        input('Press Enter to exit')
        exit()


def data_process(raw, track_type, line, track, section):
    """
    Processes the data in the .datac file.

    For each location of each requirement, groups and takes the max/ min value of the data out of the 4 channels.

    Args:
        raw (DataFrame): The data from input .datac file.
        track_type (Dataframe): The track type at different locations of the line.

    Returns:
        raw (DataFrame): Cleaned and processed data from input .datac file.
        WH_cleaned_max (DataFrame): The high height values.
        WH_cleaned_min (DataFrame): The low height exception values.
        wear_min (DataFrame): The wire wear exception values.
        stagger_left (DataFrame): The left stagger values.
        stagger_right (DataFrame): The right stagger values.
    """
    try:
        raw.columns = raw.columns.str.replace(' ', '')
        raw.drop(raw[raw.KM == 'KM'].index, inplace=True)
        raw = raw.rename({'STG1c': 'stagger1',
                        'STG2c': 'stagger2',
                        'STG3c': 'stagger3',
                        'STG4c': 'stagger4',
                        'RWH1mm': 'wear1',
                        'RWH2mm': 'wear2',
                        'RWH3mm': 'wear3',
                        'RWH4mm': 'wear4',
                        'WHGT1c': 'height1',
                        'WHGT2c': 'height2',
                        'WHGT3c': 'height3',
                        'WHGT4c': 'height4',
                        'LINE': 'Line',
                        'TRACK': 'Track'}, axis=1)

        raw[['KM', 'LOCATION', 'height1', 'height2', 'height3', 'height4', 'wear1',
            'wear2', 'wear3', 'wear4', 'stagger1', 'stagger2', 'stagger3', 'stagger4']] \
                = raw[['KM', 'LOCATION', 'height1', 'height2', 'height3', 'height4', 'wear1', 'wear2',
                    'wear3', 'wear4', 'stagger1', 'stagger2', 'stagger3', 'stagger4']].apply(pd.to_numeric)

        raw['Km'] = raw['KM'] + raw['LOCATION']*0.001
        raw['height1'] = raw['height1'] + 5300
        raw['height2'] = raw['height2'] + 5300
        raw['height3'] = raw['height3'] + 5300
        raw['height4'] = raw['height4'] + 5300
        raw = raw[['Line', 'Track', 'Km', 'height1', 'height2', 'height3', 'height4', 'wear1',
                'wear2', 'wear3', 'wear4', 'stagger1', 'stagger2', 'stagger3', 'stagger4']]
        raw['Km'] = raw['Km'].round(decimals=5)
        raw['Section'] = ''

        # Finds and assigns the corresponding section to all data points
        exception_boundary = pd.read_excel('./metadata/' + line + ' metadata.xlsx', sheet_name='Exception Boundarys')
        if track == 'UP':
            exception_boundary = m_to_km(exception_boundary, 'Up Track From', 'Up Track To')
            exception_boundary = exception_boundary[['Class', 'Up Track From', 'Up Track To']]\
                .rename({'Up Track From': 'FromKM', 'Up Track To': 'ToKM'}, axis=1)
        else:
            exception_boundary = m_to_km(exception_boundary, 'Down Track From', 'Down Track To')
            exception_boundary = exception_boundary[['Class', 'Down Track From', 'Down Track To']]\
                .rename({'Down Track From': 'FromKM', 'Down Track To': 'ToKM'}, axis=1)
        
        if line == 'TML':
            MOL_FromKM = exception_boundary.loc[exception_boundary['Class'] == 'MOL', 'FromKM'].values[0]
            MOL_ToKM = exception_boundary.loc[exception_boundary['Class'] == 'MOL', 'ToKM'].values[0]
            SCL_FromKM = exception_boundary.loc[exception_boundary['Class'] == 'SCL', 'FromKM'].values[0]
            SCL_ToKM = exception_boundary.loc[exception_boundary['Class'] == 'SCL', 'ToKM'].values[0]
            ETSE_FromKM = exception_boundary.loc[exception_boundary['Class'] == 'ETSE', 'FromKM'].values[0]
            ETSE_ToKM = exception_boundary.loc[exception_boundary['Class'] == 'ETSE', 'ToKM'].values[0]
            KSL_FromKM = exception_boundary.loc[exception_boundary['Class'] == 'KSL', 'FromKM'].values[0]
            KSL_ToKM = exception_boundary.loc[exception_boundary['Class'] == 'KSL', 'ToKM'].values[0]
            WRL_FromKM = exception_boundary.loc[exception_boundary['Class'] == 'WRL', 'FromKM'].values[0]
            WRL_ToKM = exception_boundary.loc[exception_boundary['Class'] == 'WRL', 'ToKM'].values[0]

            if track == 'UP':
                MOL_row_location = raw[(raw['Km'] >= MOL_FromKM) & (raw['Km'] <= MOL_ToKM)].index
                SCL_row_location = raw[(raw['Km'] >= SCL_FromKM) & (raw['Km'] <= SCL_ToKM)].index
                ETSE_row_location = raw[(raw['Km'] >= ETSE_FromKM) & (raw['Km'] <= ETSE_ToKM)].index
                KSL_row_location = raw[(raw['Km'] >= KSL_FromKM) & (raw['Km'] <= KSL_ToKM)].index
                WRL_row_location = raw[(raw['Km'] >= WRL_FromKM) & (raw['Km'] <= WRL_ToKM)].index
            else:
                MOL_row_location = raw[(raw['Km'] <= MOL_FromKM) & (raw['Km'] >= MOL_ToKM)].index
                SCL_row_location = raw[(raw['Km'] <= SCL_FromKM) & (raw['Km'] >= SCL_ToKM)].index
                ETSE_row_location = raw[(raw['Km'] <= ETSE_FromKM) & (raw['Km'] >= ETSE_ToKM)].index
                KSL_row_location = raw[(raw['Km'] <= KSL_FromKM) & (raw['Km'] >= KSL_ToKM)].index
                WRL_row_location = raw[(raw['Km'] <= WRL_FromKM) & (raw['Km'] >= WRL_ToKM)].index

            raw.loc[MOL_row_location, 'Section'] = 'MOL'
            raw.loc[SCL_row_location, 'Section'] = 'SCL'
            raw.loc[ETSE_row_location, 'Section'] = 'ETSE'
            raw.loc[KSL_row_location, 'Section'] = 'KSL'
            raw.loc[WRL_row_location, 'Section'] = 'WRL'
        elif (section not in ['LMC', 'RAC', 'LOW']):
            SCL_FromKM = exception_boundary.loc[exception_boundary['Class'] == 'SCL', 'FromKM'].values[0]
            SCL_ToKM = exception_boundary.loc[exception_boundary['Class'] == 'SCL', 'ToKM'].values[0]
            EAL_FromKM = exception_boundary.loc[exception_boundary['Class'] == 'EAL', 'FromKM'].values[0]
            EAL_ToKM = exception_boundary.loc[exception_boundary['Class'] == 'EAL', 'ToKM'].values[0]
            if track == 'UP':
                SCL_row_location = raw[(raw['Km'] >= SCL_FromKM) & (raw['Km'] <= SCL_ToKM)].index
                EAL_row_location = raw[(raw['Km'] >= EAL_FromKM) & (raw['Km'] <= EAL_ToKM)].index
            else:
                SCL_row_location = raw[(raw['Km'] <= SCL_FromKM) & (raw['Km'] >= SCL_ToKM)].index
                EAL_row_location = raw[(raw['Km'] <= EAL_FromKM) & (raw['Km'] >= EAL_ToKM)].index

            raw.loc[SCL_row_location, 'Section'] = 'SCL'
            raw.loc[EAL_row_location, 'Section'] = 'EAL'

        # ---------- removes unreasonable CW height data ----------
        WH_min = raw.groupby('Km')[['height1', 'height2', 'height3', 'height4']].min().reset_index()
        WH_min.loc[(WH_min['height1'] < 3500) | (WH_min['height2'] < 3500) |
                (WH_min['height3'] < 3500) | (WH_min['height4'] < 3500), 'error'] = WH_min['Km']
        WH_error = WH_min['error'].dropna().to_list()
        WH = raw[['Km', 'height1', 'height2', 'height3', 'height4', 'Section']]
        WH_cleaned = WH[~WH.Km.isin(WH_error)]
        # ---------- removes unreasonable CW height data ----------

        # ---------- preprocess the data before generating exception ----------
        stagger_left = raw.groupby('Km')[['stagger1', 'stagger2', 'stagger3', 'stagger4', 'Section']]\
            .max()\
            .reset_index()
        stagger_right = raw.groupby('Km')[['stagger1', 'stagger2', 'stagger3', 'stagger4', 'Section']]\
            .min()\
            .reset_index()
        wear_min = raw.groupby('Km')[['wear1', 'wear2', 'wear3', 'wear4', 'Section']]\
            .min()\
            .reset_index()
        WH_cleaned_max = WH_cleaned.groupby('Km')[['height1', 'height2', 'height3', 'height4', 'Section']]\
            .max()\
            .reset_index()
        WH_cleaned_min = WH_cleaned.groupby('Km')[['height1', 'height2', 'height3', 'height4', 'Section']]\
            .min()\
            .reset_index()
        # ---------- preprocess the data before generating exception ----------

        # ---------- identify track type (tangent/curve) ----------
        stagger_left = stagger_left \
            .assign(key=1) \
            .merge(track_type.assign(key=1), on='key') \
            .query('`Km`.between(`startKM`, `endKM`)', engine='python') \
            .drop(columns=['startKM', 'endKM', 'key']) \
            .reset_index(drop=True)
        
        stagger_right = stagger_right \
            .assign(key=1) \
            .merge(track_type.assign(key=1), on='key') \
            .query('`Km`.between(`startKM`, `endKM`)', engine='python') \
            .drop(columns=['startKM', 'endKM', 'key']) \
            .reset_index(drop=True)

        wear_min = wear_min \
            .assign(key=1) \
            .merge(track_type.assign(key=1), on='key') \
            .query('`Km`.between(`startKM`, `endKM`)', engine='python') \
            .drop(columns=['startKM', 'endKM', 'key']) \
            .reset_index(drop=True)

        WH_cleaned_max = WH_cleaned_max \
            .assign(key=1) \
            .merge(track_type.assign(key=1), on='key') \
            .query('`Km`.between(`startKM`, `endKM`)', engine='python') \
            .drop(columns=['startKM', 'endKM', 'key']) \
            .reset_index(drop=True)

        WH_cleaned_min = WH_cleaned_min \
            .assign(key=1) \
            .merge(track_type.assign(key=1), on='key') \
            .query('`Km`.between(`startKM`, `endKM`)', engine='python') \
            .drop(columns=['startKM', 'endKM', 'key']) \
            .reset_index(drop=True)
        # ---------- identify track type (tangent/curve) ----------

        # ---------- find extreme value amongst 4 channels ----------
        WH_cleaned_min['maxValue'] = WH_cleaned_min[['height1', 'height2', 'height3', 'height4']].min(axis=1)
        WH_cleaned_max['maxValue'] = WH_cleaned_max[['height1', 'height2', 'height3', 'height4']].max(axis=1)

        wear_min['maxValue'] = wear_min[['wear1', 'wear2', 'wear3', 'wear4']].min(axis=1)

        stagger_left['maxValue'] = stagger_left[['stagger1', 'stagger2', 'stagger3', 'stagger4']].max(axis=1)
        stagger_right['maxValue'] = stagger_right[['stagger1', 'stagger2', 'stagger3', 'stagger4']].min(axis=1)
        # ---------- find extreme value amongst 4 channels ----------

        return raw, WH_cleaned_max, WH_cleaned_min, wear_min, stagger_left, stagger_right
    except Exception:
        print(traceback.format_exc())
        print('This error occured when processing the .datac file.')
        input('Press Enter to exit')
        exit()


def min_aggregate_group(g):
    """
    A helper function for applying chain length.

    Takes the row with the smallest maxValue and uses it in the new created row.

    Args:
        g (DataFrame): The DataFrame that contains all the rows for a specific group to be chained.

        Returns:
        (pandas.Series): The row with the min aggregrated values.
    """
    startM = g['startKm'].min()
    endM = g['endKm'].max()
    maxValue = g['maxValue'].min()
    max_row = g.loc[g['maxValue'].idxmin()]
    length = endM - startM
      
    return pd.Series({
        'startKm': startM,
        'endKm': endM,
        'length': length,
        'exception type': max_row['exception type'],
        'maxValue': maxValue,
        'maxLocation': max_row['maxLocation'],
        'track type': max_row['track type'],
        'Section': max_row['Section']
        })


def max_aggregate_group(g):
    """
    A helper function for applying chain length.

    Takes the row with the largest maxValue and uses it in the new created row.

    Args:
        g (DataFrame): The DataFrame that contains all the rows for a specific group to be chained.

    Returns:
        pandas Series: The row with the max aggregrated values.
    """
    startM = g['startKm'].min()
    endM = g['endKm'].max()
    maxValue = g['maxValue'].max()
    max_row = g.loc[g['maxValue'].idxmax()]
    length = endM - startM
   
    return pd.Series({
        'startKm': startM,
        'endKm': endM,
        'length': length,
        'exception type': max_row['exception type'],
        'maxValue': maxValue,
        'maxLocation': max_row['maxLocation'],
        'track type': max_row['track type'],
        'Section': max_row['Section']
    })


def km_to_m(exception_table):
    """
    Converts the unit of the location from km to meter.

    Args:
        exception_table (DataFrame): The exception table.

    Returns:
        exception_table (DataFrame): The exception table with converted unit.
    """
    exception_table = exception_table.rename({'startKm': 'startM', 'endKm': 'endM'}, axis=1)
    exception_table['startM'] = exception_table['startM'] * 1000
    exception_table['endM'] = exception_table['endM'] * 1000
    exception_table['maxLocation'] = exception_table['maxLocation'] * 1000
    exception_table['length'] = exception_table['length'] * 1000
    return exception_table


def sort_table(exception_table):
    """
    Sorts the table by the id of each exception location.
    Gets the integer in the back of the id and sorts the table by the integer.

    Args:
        exception_table (DataFrame): The exception table to be sorted.

    Returns:
        exceptino_table (DataFrame): The sorted exception table.
    """
    exception_table['id_number'] = exception_table['id'].str.extract(r'(\d+)$').astype(int)
    exception_table = exception_table.sort_values(by='id_number').drop(columns='id_number')
    return exception_table


def find_low_height_exception(WH_cleaned_min, threshold, date, line, section, track):
    """
    Finds all the exception of low height and outputs a table of it.

    Gets low height threshold values from metadata first,
    then distinguishes the corresponding threshold value to be used on the each data point.
    Compares the maxValue of each point and assigns the corresponding level of danger to it.

    Args:
        WH_cleaned_min (DataFrame): The minimum low height exception values out of the four channels.
        threshold (DataFrame): The threshold of the line.
        date (str): The date of the data collected.
        line (str): The name of the line.
        section (str): The section of the line.
        track (str): The track of the line.

    Returns:
        low_height_exception (DataFrame): The details of the found low height exception.
    """
    # Loads the corresponding low height exception boundary values
    if line == 'TML':
        WRL_low_height_L1_max = threshold.loc[(threshold['Class'] == 'WRL') &
                                              (threshold['Exc Type'] == 'Low Height L1')]['max'].values.item()
        # WRL_low_height_L2_min = threshold.loc[(threshold['Class'] == 'WRL') &
        #                                       (threshold['Exc Type'] == 'Low Height L2')]['min'].values.item()
        WRL_low_height_L2_max = threshold.loc[(threshold['Class'] == 'WRL') &
                                              (threshold['Exc Type'] == 'Low Height L2')]['max'].values.item()
        KSL_low_height_L1_max = threshold.loc[(threshold['Class'] == 'KSL') &
                                              (threshold['Exc Type'] == 'Low Height L1')]['max'].values.item()
        # KSL_low_height_L2_min = threshold.loc[(threshold['Class'] == 'KSL') &
        #                                       (threshold['Exc Type'] == 'Low Height L2')]['min'].values.item()
        KSL_low_height_L2_max = threshold.loc[(threshold['Class'] == 'KSL') &
                                              (threshold['Exc Type'] == 'Low Height L2')]['max'].values.item()
        ETSE_low_height_L1_max = threshold.loc[(threshold['Class'] == 'ETSE') &
                                               (threshold['Exc Type'] == 'Low Height L1')]['max'].values.item()
        # ETSE_low_height_L2_min = threshold.loc[(threshold['Class'] == 'ETSE') &
        #                                        (threshold['Exc Type'] == 'Low Height L2')]['min'].values.item()
        ETSE_low_height_L2_max = threshold.loc[(threshold['Class'] == 'ETSE') &
                                               (threshold['Exc Type'] == 'Low Height L2')]['max'].values.item()
        MOL_low_height_L1_max = threshold.loc[(threshold['Class'] == 'MOL') &
                                              (threshold['Exc Type'] == 'Low Height L1')]['max'].values.item()
        # MOL_low_height_L2_min = threshold.loc[(threshold['Class'] == 'MOL') &
        #                                       (threshold['Exc Type'] == 'Low Height L2')]['min'].values.item()
        MOL_low_height_L2_max = threshold.loc[(threshold['Class'] == 'MOL') &
                                              (threshold['Exc Type'] == 'Low Height L2')]['max'].values.item()
        SCL_low_height_L1_max = threshold.loc[(threshold['Class'] == 'SCL') &
                                              (threshold['Exc Type'] == 'Low Height L1')]['max'].values.item()
        # SCL_low_height_L2_min = threshold.loc[(threshold['Class'] == 'SCL') &
        #                                       (threshold['Exc Type'] == 'Low Height L2')]['min'].values.item()
        SCL_low_height_L2_max = threshold.loc[(threshold['Class'] == 'SCL') &
                                              (threshold['Exc Type'] == 'Low Height L2')]['max'].values.item()
    else:
        SCL_low_height_L1_max = threshold.loc[(threshold['Class'] == 'SCL') &
                                              (threshold['Exc Type'] == 'Low Height L1')]['max'].values.item()
        # SCL_low_height_L2_min = np.nan
        SCL_low_height_L2_max = threshold.loc[(threshold['Class'] == 'SCL') &
                                              (threshold['Exc Type'] == 'Low Height L2')]['max'].values.item()
        if section in ['LMC']:
            class_low_height_L1_max = threshold.loc[(threshold['Class'] == 'LMC') &
                                                    (threshold['Exc Type'] == 'Low Height L1')]['max'].values.item()
            # class_low_height_L2_min = threshold.loc[(threshold['Class'] == 'LMC') &
            #                                         (threshold['Exc Type'] == 'Low Height L2')]['min'].values.item()
            class_low_height_L2_max = threshold.loc[(threshold['Class'] == 'LMC') &
                                                    (threshold['Exc Type'] == 'Low Height L2')]['max'].values.item()
        else:
            class_low_height_L1_max = threshold.loc[(threshold['Class'] == 'both') &
                                                    (threshold['Exc Type'] == 'Low Height L1')]['max'].values.item()
            # class_low_height_L2_min = threshold.loc[(threshold['Class'] == 'both') &
            #                                         (threshold['Exc Type'] == 'Low Height L2')]['min'].values.item()
            class_low_height_L2_max = threshold.loc[(threshold['Class'] == 'both') &
                                                    (threshold['Exc Type'] == 'Low Height L2')]['max'].values.item()

    # ---------- low height exception ----------
    WH_cleaned_min['Km_roundup'] = WH_cleaned_min['Km'].round(3)

    # Assigns all exceptions to L2 for now
    WH_cleaned_min['L2_max'] = ''

    if (line == 'TML'):
        WRL_L2_max_row_location = WH_cleaned_min[WH_cleaned_min['Section'] == 'WRL'].index
        WH_cleaned_min.loc[WRL_L2_max_row_location, 'L2_max'] = WRL_low_height_L2_max
        KSL_L2_max_row_location = WH_cleaned_min[WH_cleaned_min['Section'] == 'KSL'].index
        WH_cleaned_min.loc[KSL_L2_max_row_location, 'L2_max'] = KSL_low_height_L2_max
        ETSE_L2_max_row_location = WH_cleaned_min[WH_cleaned_min['Section'] == 'ETSE'].index
        WH_cleaned_min.loc[ETSE_L2_max_row_location, 'L2_max'] = ETSE_low_height_L2_max
        MOL_L2_max_row_location = WH_cleaned_min[WH_cleaned_min['Section'] == 'MOL'].index
        WH_cleaned_min.loc[MOL_L2_max_row_location, 'L2_max'] = MOL_low_height_L2_max
        SCL_L2_max_row_location = WH_cleaned_min[WH_cleaned_min['Section'] == 'SCL'].index
        WH_cleaned_min.loc[SCL_L2_max_row_location, 'L2_max'] = SCL_low_height_L2_max
    else:
        if (section in ['LMC', 'RAC', 'LOW']):
            WH_cleaned_min['L2_max'] = class_low_height_L2_max
        else:
            WH_cleaned_min['L2_max'] = class_low_height_L2_max
            SCL_L2_max_row_location = WH_cleaned_min[WH_cleaned_min['Section'] == 'SCL'].index
            WH_cleaned_min.loc[SCL_L2_max_row_location, 'L2_max'] = SCL_low_height_L2_max

    WH_cleaned_min['L2'] = (WH_cleaned_min.maxValue <= WH_cleaned_min.L2_max)
    WH_cleaned_min['L2_id'] = (WH_cleaned_min.L2 != WH_cleaned_min.L2.shift()).cumsum()
    WH_cleaned_min['L2_count'] = WH_cleaned_min.groupby(['L2', 'L2_id']).cumcount(ascending=False) + 1
    WH_cleaned_min.loc[~WH_cleaned_min['L2'], 'L2_count'] = 0

    if WH_cleaned_min['L2'].any():
        low_height_exception_full = WH_cleaned_min[(WH_cleaned_min['L2_count'] != 0)]

        low_height_exception_full.loc[low_height_exception_full['L2'], 'exception type'] = 'Low Height'

        low_height_exception = (low_height_exception_full
                                .groupby(['exception type', 'L2_id'])
                                .agg({'Km': ['min', 'max'], 'maxValue': ['min', 'max'], 'Km_roundup': ['min', 'max']}))
        low_height_exception.columns = low_height_exception.columns.map('_'.join)
        low_height_exception = (low_height_exception
                                .assign(key=1)
                                    .merge(low_height_exception_full.assign(key=1), on='key')
                                        .query('`maxValue_min` == `maxValue` & `Km_roundup`.between(`Km_roundup_min`, `Km_roundup_max`)', engine='python')
                                            .drop(columns=['maxValue_max', 'key', 'maxValue', 'L2_id', 'L2_count',
                                                           'height1', 'height2', 'height3', 'height4', 'L2'])
                                                            .rename({'Km': 'maxLocation', 'Km_roundup_min': 'startKm',
                                                                     'Km_roundup_max': 'endKm', 'maxValue_min': 'maxValue'}, axis=1)
                                                                        .reset_index()
                                                                            .drop('index', axis=1))
        low_height_exception['length'] = low_height_exception['endKm'] - low_height_exception['startKm']
        low_height_exception = (low_height_exception[['exception type', 'startKm', 'endKm', 'length',
                                                      'maxValue', 'maxLocation', 'track type', 'Section']]
                                                        .sort_values('startKm')
                                                            .drop_duplicates(subset=['exception type', 'startKm', 'endKm'])
                                                                .reset_index()
                                                                    .drop('index', axis=1))

        # apply chain length
        low_height_exception['group'] = (low_height_exception['startKm'] - low_height_exception['endKm'].shift().fillna(0) >= chain_length).cumsum()
        low_height_exception = low_height_exception.groupby('group').apply(min_aggregate_group).reset_index(drop=True)
        # ------ defining alarm level depending on maxValue only ------
        # Assigns L1 exception
        WH_cleaned_min['L1_max'] = ''

        if (line == 'TML'):
            WRL_L1_max_row_location = low_height_exception[low_height_exception['Section'] == 'WRL'].index
            low_height_exception.loc[WRL_L1_max_row_location, 'L1_max'] = WRL_low_height_L1_max
            KSL_L1_max_row_location = low_height_exception[low_height_exception['Section'] == 'KSL'].index
            low_height_exception.loc[KSL_L1_max_row_location, 'L1_max'] = KSL_low_height_L1_max
            ETSE_L1_max_row_location = low_height_exception[low_height_exception['Section'] == 'ETSE'].index
            low_height_exception.loc[ETSE_L1_max_row_location, 'L1_max'] = ETSE_low_height_L1_max
            MOL_L1_max_row_location = low_height_exception[low_height_exception['Section'] == 'MOL'].index
            low_height_exception.loc[MOL_L1_max_row_location, 'L1_max'] = MOL_low_height_L1_max
            SCL_L1_max_row_location = low_height_exception[low_height_exception['Section'] == 'SCL'].index
            low_height_exception.loc[SCL_L1_max_row_location, 'L1_max'] = SCL_low_height_L1_max
        else:
            if (section in ['LOW', 'RAC', 'LMC']):
                low_height_exception['L1_max'] = class_low_height_L1_max
            else:
                low_height_exception['L1_max'] = class_low_height_L1_max
                SCL_L1_max_row_location = low_height_exception[low_height_exception['Section'] == 'SCL'].index
                low_height_exception.loc[SCL_L1_max_row_location, 'L1_max'] = SCL_low_height_L1_max

        low_height_exception.loc[(low_height_exception.maxValue <= low_height_exception.L1_max), 'level'] = 'L1'

        low_height_exception['level'] = low_height_exception['level'].replace('na', np.nan)
        low_height_exception['level'] = low_height_exception['level'].fillna('L2')
        low_height_exception = low_height_exception[['exception type', 'level', 'startKm', 'endKm',
                                                     'length', 'maxValue', 'maxLocation', 'track type']].reset_index()
        low_height_exception['id'] = (date + '_' + line + '_' + section + '_' + track + '_' + 'LH' +
                                      low_height_exception['index'].astype(str))
        low_height_exception = low_height_exception[
            ['id', 'exception type', 'level', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type']]
        # ------ defining alarm level depending on maxValue only ------
    else:
        low_height_exception = pd.DataFrame(columns=['id', 'exception type', 'level', 'startKm', 'endKm', 'length',
                                                     'maxValue', 'maxLocation', 'track type'])
    # ---------- low height exception ----------

    return low_height_exception


def find_high_height_exception(WH_cleaned_max, threshold, date, line, section, track):
    """
    Finds all the exception of high height and outputs a table of it.

    Gets high height threshold values from metadata first,
    then distinguishes the corresponding threshold value to be used on the each data point.
    Compares the maxValue of each point and assigns the corresponding level of danger to it.

    Args:
        WH_cleaned_min (DataFrame): The maximum high height exception values out of the four channels.
        threshold (DataFrame): The threshold of the line.
        date (str): The date of the data collected.
        line (str): The name of the line.
        section (str): The section of the line.
        track (str): The track of the line.

    Returns:
        high_height_exception (DataFrame): The DataFrame with details of the found high height exception.
    """
    # ---------- high height exception ----------
    high_height_L1_min = threshold.loc[threshold['Exc Type'] == 'High Height L1']['min'].values.item()
    high_height_L2_min = threshold.loc[threshold['Exc Type'] == 'High Height L2']['min'].values.item()
    # high_height_L2_max = threshold.loc[threshold['Exc Type'] == 'High Height L2']['max'].values.item()

    WH_cleaned_max['Km_roundup'] = WH_cleaned_max['Km'].round(3)
    WH_cleaned_max['L2'] = (WH_cleaned_max.maxValue >= high_height_L2_min)
    WH_cleaned_max['L2_id'] = (WH_cleaned_max.L2 != WH_cleaned_max.L2.shift()).cumsum()
    WH_cleaned_max['L2_count'] = WH_cleaned_max.groupby(['L2', 'L2_id']).cumcount(ascending=False) + 1
    WH_cleaned_max.loc[~WH_cleaned_max['L2'], 'L2_count'] = 0

    if WH_cleaned_max['L2'].any():
        high_height_exception_full = WH_cleaned_max[(WH_cleaned_max['L2_count'] != 0)]
        high_height_exception_full.loc[high_height_exception_full['L2'], 'exception type'] = 'High Height'

        high_height_exception = high_height_exception_full\
            .groupby(['exception type', 'L2_id'])\
            .agg({'Km': ['min', 'max'], 'maxValue': ['min', 'max'], 'Km_roundup': ['min', 'max']})
        high_height_exception.columns = high_height_exception.columns.map('_'.join)
        high_height_exception = high_height_exception\
            .assign(key=1)\
                .merge(high_height_exception_full.assign(key=1), on='key')\
                    .query('`maxValue_max` == `maxValue` & `Km_roundup`.between(`Km_roundup_min`, `Km_roundup_max`)', engine='python')\
                        .drop(columns=['maxValue_min', 'key', 'maxValue', 'L2_id', 'L2_count', 'height1',
                                       'height2', 'height3', 'height4', 'L2'])\
                                        .rename({'Km': 'maxLocation', 'Km_roundup_min': 'startKm', 'Km_roundup_max': 'endKm',
                                                 'maxValue_max': 'maxValue'}, axis=1)\
                                                    .reset_index()\
                                                        .drop('index', axis=1)

        high_height_exception['length'] = high_height_exception['endKm'] - high_height_exception['startKm']
        high_height_exception = (high_height_exception[['exception type', 'startKm', 'endKm', 'length',
                                                        'maxValue', 'maxLocation', 'track type', 'Section']]
                                                            .sort_values('startKm')
                                                                .drop_duplicates(subset=['exception type', 'startKm', 'endKm'])
                                                                    .reset_index()
                                                                        .drop('index', axis=1))

        # apply chain length
        high_height_exception['group'] = (high_height_exception['startKm'] - high_height_exception['endKm'].shift().fillna(0) >= chain_length).cumsum()
        high_height_exception = high_height_exception.groupby('group').apply(max_aggregate_group).reset_index(drop=True)

        # ------ defining alarm level depending on maxValue only ------
        high_height_exception.loc[(high_height_exception['maxValue'] >= high_height_L1_min), 'level'] = 'L1'

        high_height_exception['level'] = high_height_exception['level'].replace('na', np.nan)
        high_height_exception['level'] = high_height_exception['level'].fillna('L2')
        high_height_exception = high_height_exception[['exception type', 'level', 'startKm', 'endKm',
                                                       'length', 'maxValue', 'maxLocation', 'track type']].reset_index()
        high_height_exception['id'] = (date + '_' + line + '_' + section + '_' + track + '_' + 'HH' +
                                       high_height_exception['index'].astype(str))
        high_height_exception = high_height_exception[
            ['id', 'exception type', 'level', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type']]
        # ------ defining alarm level depending on maxValue only ------
    else:
        high_height_exception = pd.DataFrame(columns=['id', 'exception type', 'level', 'startKm', 'endKm',
                                                      'length', 'maxValue', 'maxLocation', 'track type'])
    # ---------- high height exception ----------
    
    return high_height_exception


def find_wire_wear_exception(wear_min, threshold, date, line, section, track):
    """
    Finds all the exception of wire wear and output a table of it.

    Gets wire wear threshold values from metadata first,
    then distinguishes the corresponding threshold value to be used on the each data point.
    Compares the maxValue of each point and assigns the corresponding level of danger to it.

    Args:
        wear_min (DataFrame): The minimum wire wear exception values out of the four channels.
        threshold (DataFrame): The threshold of the line.
        date (str): The date of the data collected.
        line (str): The name of the line.
        section (str): The section of the line.
        track (str): The track of the line.

    Returns:
        wear_exception (DataFrame): The details of the found wire wear exception.
    """
    # Loads line related threshold from metadata
    wear_L1_max = threshold.loc[(threshold['Exc Type'] == 'Wire Wear L1') &
                                (threshold['Class'] == 'both')]['max'].values[0]
    # wear_L2_min = threshold.loc[threshold['Exc Type'] == 'Wire Wear L2']['min'].values.item()
    wear_L2_max = threshold.loc[(threshold['Exc Type'] == 'Wire Wear L2') &
                                (threshold['Class'] == 'both')]['max'].values[0]

    SCL_wear_L1_max = threshold.loc[(threshold['Exc Type'] == 'Wire Wear L1') &
                                (threshold['Class'] == 'SCL')]['max'].values[0]
    SCL_wear_L2_max = threshold.loc[(threshold['Exc Type'] == 'Wire Wear L2') &
                                (threshold['Class'] == 'SCL')]['max'].values[0]

    # Assigns all exceptions to L2 for now
    wear_min['L2_max'] = ''
    if (line == 'EAL'):
        if (section in ['LMC', 'RAC', 'LOW']):
            wear_min['L2_max'] = wear_L2_max
        else:
            wear_min['L2_max'] = wear_L2_max
            SCL_L2_max_row_location = wear_min[wear_min['Section'] == 'SCL'].index
            wear_min.loc[SCL_L2_max_row_location, 'L2_max'] = SCL_wear_L2_max
    else:   # TML
        wear_min['L2_max'] = wear_L2_max
        SCL_L2_max_row_location = wear_min[wear_min['Section'] == 'SCL'].index
        wear_min.loc[SCL_L2_max_row_location, 'L2_max'] = SCL_wear_L2_max

    wear_min['Km_roundup'] = wear_min['Km'].round(3)
    wear_min['L2'] = (wear_min.maxValue <= wear_min.L2_max)
    wear_min['L2_id'] = (wear_min.L2 != wear_min.L2.shift()).cumsum()
    wear_min['L2_count'] = wear_min.groupby(['L2', 'L2_id']).cumcount(ascending=False) + 1
    wear_min.loc[~wear_min['L2'], 'L2_count'] = 0
    if wear_min['L2'].any():
        wear_exception_full = wear_min[(wear_min['L2_count'] != 0)]
        wear_exception_full.loc[wear_exception_full['L2'], 'exception type'] = 'Wire Wear'
        wear_exception = wear_exception_full\
            .groupby(['exception type', 'L2_id'])\
            .agg({'Km': ['min', 'max'], 'maxValue': ['min', 'max'], 'Km_roundup': ['min', 'max']})
        wear_exception.columns = wear_exception.columns.map('_'.join)
        wear_exception = (wear_exception\
                          .assign(key=1)\
                            .merge(wear_exception_full.assign(key=1), on='key')\
                                .query('`maxValue_min` == `maxValue` & `Km_roundup`.between(`Km_roundup_min`, `Km_roundup_max`)', engine='python')\
                                    .drop(columns=['maxValue_max', 'key', 'maxValue', 'L2_id', 'L2_count',
                                                   'wear1', 'wear2', 'wear3', 'wear4', 'L2'])\
                                                    .rename({'Km': 'maxLocation', 'Km_roundup_min': 'startKm', 'Km_roundup_max': 'endKm',
                                                             'maxValue_min': 'maxValue'}, axis=1)\
                                                                .reset_index()\
                                                                    .drop('index', axis=1))
        wear_exception['length'] = wear_exception['endKm'] - wear_exception['startKm']
        wear_exception = (wear_exception[['exception type', 'startKm', 'endKm', 'length',
                                          'maxValue', 'maxLocation', 'track type', 'Section']]\
                                            .sort_values('startKm')\
                                                .drop_duplicates(subset=['exception type', 'startKm', 'endKm'])\
                                                    .reset_index()\
                                                        .drop('index', axis=1))
        
        # apply chain length
        wear_exception['group'] = (wear_exception['startKm'] - wear_exception['endKm'].shift().fillna(0) >= chain_length).cumsum()
        wear_exception = wear_exception.groupby('group').apply(min_aggregate_group).reset_index(drop=True)

        # ------ defining alarm level depending on maxValue only ------
        # Assigns L1 exception
        wear_exception['L1_max'] = ''
        if (line == 'EAL'):
            if (section in ['LMC', 'RAC', 'LOW']):
                wear_exception['L1_max'] = wear_L1_max
            else:
                wear_exception['L1_max'] = wear_L1_max
                SCL_L1_max_row_location = wear_exception[wear_exception['Section'] == 'SCL'].index
                wear_exception.loc[SCL_L1_max_row_location, 'L1_max'] = SCL_wear_L1_max
        else:   # TML
            wear_exception['L1_max'] = wear_L1_max
            SCL_L1_max_row_location = wear_exception[wear_exception['Section'] == 'SCL'].index
            wear_exception.loc[SCL_L1_max_row_location, 'L1_max'] = SCL_wear_L1_max

        wear_exception.loc[(wear_exception['maxValue'] <= wear_exception['L1_max']), 'level'] = 'L1'
        wear_exception['level'] = wear_exception['level'].replace('na', np.nan)
        wear_exception['level'] = wear_exception['level'].fillna('L2')
        wear_exception = wear_exception[['exception type', 'level', 'startKm', 'endKm', 'length',
                                         'maxValue', 'maxLocation', 'track type']].reset_index()
        wear_exception['id'] = (date + '_' + line + '_' + section + '_' + track + '_' + 'W' +
                                wear_exception['index'].astype(str))
        wear_exception = wear_exception[['id', 'exception type', 'level', 'startKm',
                                         'endKm', 'length', 'maxValue', 'maxLocation', 'track type']]
        # ------ defining alarm level depending on maxValue only ------
    else:
        wear_exception = pd.DataFrame(columns=['id', 'exception type', 'level', 'startKm', 'endKm',
                                               'length', 'maxValue', 'maxLocation', 'track type'])
    
    return wear_exception


def find_stagger_exception(stagger_left, stagger_right, threshold, date, line, section, track):
    """
    Finds all the exception of both left and right stagger
    and outputs two tables for left and right stagger respectively.

    Gets stagger threshold values from metadata first,
    then distinguishes the corresponding threshold value to be used on the each data point.
    Compares the maxValue of each point and assigns the corresponding level of danger to it.

    Args:
        stagger_left (DataFrame): The maximum left stagger exception values out of the four channels.
        stagger_right (DataFrame): The minimum (negative) right stagger exception valus out of the four channels.
        threshold (DataFrame): The threshold of the line.
        date (str): The date of the data collected.
        line (str): The name of the line.
        section (str): The section of the line.
        track (str): The track of the line.

    Returns:
        stagger_left_exception (DataFrame): The details of the found left stagger exception.
        stagger_right_exception (DataFrame): The details of the found right stagger exception.
    """
    # Loads line related threshold from metadata
    if line == 'TML':
        KSL_tangent_stagger_L1_min = threshold.loc[(threshold['Class'] == 'KSL') &
                                                   (threshold['Track Type'] == 'Tangent') &
                                                   (threshold['Exc Type'] == 'Stagger L1')]['min'].values.item()
        KSL_tangent_stagger_L2_min = threshold.loc[(threshold['Class'] == 'KSL') &
                                                   (threshold['Track Type'] == 'Tangent') &
                                                   (threshold['Exc Type'] == 'Stagger L2')]['min'].values.item()
        KSL_tangent_stagger_L2_max = threshold.loc[(threshold['Class'] == 'KSL') &
                                                   (threshold['Track Type'] == 'Tangent') &
                                                   (threshold['Exc Type'] == 'Stagger L2')]['max'].values.item()
        KSL_tangent_stagger_L3_min = threshold.loc[(threshold['Class'] == 'KSL') &
                                                   (threshold['Track Type'] == 'Tangent') &
                                                   (threshold['Exc Type'] == 'Stagger L3')]['min'].values.item()

        # KSL_tangent_stagger_L3_max = threshold.loc[(threshold['Class'] == 'KSL') &
        #                                             (threshold['Track Type'] == 'Tangent') &
        #                                             (threshold['Exc Type'] == 'Stagger L3')]['max'].values.item()

        KSL_curve_stagger_L1_min = threshold.loc[(threshold['Class'] == 'KSL') &
                                                 (threshold['Track Type'] == 'Curve') &
                                                 (threshold['Exc Type'] == 'Stagger L1')]['min'].values.item()
        KSL_curve_stagger_L3_min = threshold.loc[(threshold['Class'] == 'KSL') &
                                                 (threshold['Track Type'] == 'Curve') &
                                                 (threshold['Exc Type'] == 'Stagger L3')]['min'].values.item()
        # KSL_curve_stagger_L3_max = threshold.loc[(threshold['Class'] == 'KSL') &
        #                                             (threshold['Track Type'] == 'Curve') &
        #                                             (threshold['Exc Type'] == 'Stagger L3')]['max'].values.item()
        SCL_tangent_stagger_L1_min = threshold.loc[(threshold['Class'] == 'SCL') &
                                                   (threshold['Track Type'] == 'Tangent') &
                                                   (threshold['Exc Type'] == 'Stagger L1')]['min'].values.item()
        SCL_tangent_stagger_L3_min = threshold.loc[(threshold['Class'] == 'SCL') &
                                                   (threshold['Track Type'] == 'Tangent') &
                                                   (threshold['Exc Type'] == 'Stagger L3')]['min'].values.item()
        # SCL_tangent_stagger_L3_max = threshold.loc[(threshold['Class'] == 'SCL') &
        #                                             (threshold['Track Type'] == 'Tangent') &
        #                                             (threshold['Exc Type'] == 'Stagger L3')]['max'].values.item()
            
        SCL_curve_stagger_L1_min = threshold.loc[(threshold['Class'] == 'SCL') &
                                                 (threshold['Track Type'] == 'Curve') &
                                                 (threshold['Exc Type'] == 'Stagger L1')]['min'].values.item()
        SCL_curve_stagger_L3_min = threshold.loc[(threshold['Class'] == 'SCL') &
                                                 (threshold['Track Type'] == 'Curve') &
                                                 (threshold['Exc Type'] == 'Stagger L3')]['min'].values.item()
        # SCL_curve_stagger_L3_max = threshold.loc[(threshold['Class'] == 'SCL') &
        #                                             (threshold['Track Type'] == 'Curve') &
        #                                             (threshold['Exc Type'] == 'Stagger L3')]['max'].values.item()
    else:
        SCL_tangent_stagger_L1_min = threshold.loc[(threshold['Class'] == 'SCL') &
                                                   (threshold['Track Type'] == 'Tangent') &
                                                   (threshold['Exc Type'] == 'Stagger L1')]['min'].values.item()
        SCL_tangent_stagger_L3_min = threshold.loc[(threshold['Class'] == 'SCL') &
                                                   (threshold['Track Type'] == 'Tangent') &
                                                   (threshold['Exc Type'] == 'Stagger L3')]['min'].values.item()
        # SCL_tangent_stagger_L3_max = np.nan
        SCL_curve_stagger_L1_min = threshold.loc[(threshold['Class'] == 'SCL') &
                                                   (threshold['Track Type'] == 'Curve') &
                                                   (threshold['Exc Type'] == 'Stagger L1')]['min'].values.item()
        SCL_curve_stagger_L3_min = threshold.loc[(threshold['Class'] == 'SCL') &
                                                   (threshold['Track Type'] == 'Curve') &
                                                   (threshold['Exc Type'] == 'Stagger L3')]['min'].values.item()
        # SCL_curve_stagger_L3_max = np.nan

    tangent_stagger_L1_min = threshold.loc[(threshold['Class'] == 'both') &
                                           (threshold['Track Type'] == 'Tangent') &
                                           (threshold['Exc Type'] == 'Stagger L1')]['min'].values.item()
    tangent_stagger_L2_min = threshold.loc[(threshold['Class'] == 'both') &
                                           (threshold['Track Type'] == 'Tangent') &
                                           (threshold['Exc Type'] == 'Stagger L2')]['min'].values.item()
    tangent_stagger_L2_max = threshold.loc[(threshold['Class'] == 'both') &
                                           (threshold['Track Type'] == 'Tangent') &
                                           (threshold['Exc Type'] == 'Stagger L2')]['max'].values.item()
    tangent_stagger_L3_min = threshold.loc[(threshold['Class'] == 'both') &
                                           (threshold['Track Type'] == 'Tangent') &
                                           (threshold['Exc Type'] == 'Stagger L3')]['min'].values.item()
    
    # tangent_stagger_L3_max = threshold.loc[(threshold['Class'] == 'both') &
    #                                        (threshold['Track Type'] == 'Tangent') &
    #                                        (threshold['Exc Type'] == 'Stagger L3')]['max'].values.item()

    curve_stagger_L1_min = threshold.loc[(threshold['Class'] == 'both') &
                                         (threshold['Track Type'] == 'Curve') &
                                         (threshold['Exc Type'] == 'Stagger L1')]['min'].values.item()
    curve_stagger_L2_min = threshold.loc[(threshold['Class'] == 'both') &
                                         (threshold['Track Type'] == 'Curve') &
                                         (threshold['Exc Type'] == 'Stagger L2')]['min'].values.item()
    curve_stagger_L2_max = threshold.loc[(threshold['Class'] == 'both') &
                                         (threshold['Track Type'] == 'Curve') &
                                         (threshold['Exc Type'] == 'Stagger L2')]['max'].values.item()
    curve_stagger_L3_min = threshold.loc[(threshold['Class'] == 'both') &
                                         (threshold['Track Type'] == 'Curve') &
                                         (threshold['Exc Type'] == 'Stagger L3')]['min'].values.item()
    # curve_stagger_L3_max = threshold.loc[(threshold['Class'] == 'both') &
    #                                      (threshold['Track Type'] == 'Curve') &
    #                                      (threshold['Exc Type'] == 'Stagger L3')]['max'].values.item()

    # ---------- Left stagger exception ----------
    stagger_left['Km_roundup'] = stagger_left['Km'].round(3)

    # Labels all found exceptions L3 for now
    stagger_left['curve_L3_min'] = float('inf')
    stagger_left['tangent_L3_min'] = float('inf')
    if (line == 'TML'):
        other_curve_L3_min_row_location = stagger_left[(stagger_left['track type'] == 'Curve')].index
        stagger_left.loc[other_curve_L3_min_row_location, 'curve_L3_min'] = curve_stagger_L3_min
        KSL_curve_L3_min_row_location = stagger_left[(stagger_left['track type'] == 'Curve') &
                                                     (stagger_left['Section'] == 'KSL')].index
        stagger_left.loc[KSL_curve_L3_min_row_location, 'curve_L3_min'] = KSL_curve_stagger_L3_min
        SCL_curve_L3_min_row_location = stagger_left[(stagger_left['track type'] == 'Curve') &
                                                     (stagger_left['Section'] == 'SCL')].index
        stagger_left.loc[SCL_curve_L3_min_row_location, 'curve_L3_min'] = SCL_curve_stagger_L3_min

        other_tangent_L3_min_row_location = stagger_left[(stagger_left['track type'] == 'Tangent')].index
        stagger_left.loc[other_tangent_L3_min_row_location, 'tangent_L3_min'] = tangent_stagger_L3_min
        KSL_tangent_L3_min_row_location = stagger_left[(stagger_left['track type'] == 'Tangent') &
                                                       (stagger_left['Section'] == 'KSL')].index
        stagger_left.loc[KSL_tangent_L3_min_row_location, 'tangent_L3_min'] = KSL_tangent_stagger_L3_min
        SCL_tangent_L3_min_row_location = stagger_left[(stagger_left['track type'] == 'Tangent') &
                                                       (stagger_left['Section'] == 'SCL')].index
        stagger_left.loc[SCL_tangent_L3_min_row_location, 'tangent_L3_min'] = SCL_tangent_stagger_L3_min
    else:
        if (section in ['LOW', 'RAC', 'LMC']):
            other_curve_L3_min_row_location = stagger_left[(stagger_left['track type'] == 'Curve')].index
            stagger_left.loc[other_curve_L3_min_row_location, 'curve_L3_min'] = curve_stagger_L3_min
            other_tangent_L3_min_row_location = stagger_left[(stagger_left['track type'] == 'Tangent')].index
            stagger_left.loc[other_tangent_L3_min_row_location, 'tangent_L3_min'] = tangent_stagger_L3_min
        else:
            other_curve_L3_min_row_location = stagger_left[(stagger_left['track type'] == 'Curve')].index
            stagger_left.loc[other_curve_L3_min_row_location, 'curve_L3_min'] = curve_stagger_L3_min
            SCL_curve_L3_min_row_location = stagger_left[(stagger_left['track type'] == 'Curve') &
                                                         (stagger_left['Section'] == 'SCL')].index
            stagger_left.loc[SCL_curve_L3_min_row_location, 'curve_L3_min'] = SCL_curve_stagger_L3_min

            other_tangent_L3_min_row_location = stagger_left[(stagger_left['track type'] == 'Tangent')].index
            stagger_left.loc[other_tangent_L3_min_row_location, 'tangent_L3_min'] = tangent_stagger_L3_min
            SCL_tangent_L3_min_row_location = stagger_left[(stagger_left['track type'] == 'Tangent') &
                                                           (stagger_left['Section'] == 'SCL')].index
            stagger_left.loc[SCL_tangent_L3_min_row_location, 'tangent_L3_min'] = SCL_tangent_stagger_L3_min

    stagger_left['curve_L3'] = (stagger_left.maxValue >= stagger_left.curve_L3_min)
    stagger_left['tangent_L3'] = (stagger_left.maxValue >= stagger_left.tangent_L3_min)
    stagger_left['curve_L3_id'] = (stagger_left.curve_L3 != stagger_left.curve_L3.shift()).cumsum()
    stagger_left['tangent_L3_id'] = (stagger_left.tangent_L3 != stagger_left.tangent_L3.shift()).cumsum()
    stagger_left['curve_L3_count'] = stagger_left.groupby(['curve_L3', 'curve_L3_id']).cumcount(ascending=False) + 1
    stagger_left['tangent_L3_count'] = stagger_left.groupby(['tangent_L3', 'tangent_L3_id']).cumcount(ascending=False) + 1
    stagger_left.loc[~stagger_left['curve_L3'], 'curve_L3_count'] = 0
    stagger_left.loc[~stagger_left['tangent_L3'], 'tangent_L3_count'] = 0
    if stagger_left['curve_L3'].any() or stagger_left['tangent_L3'].any():

        stagger_left_exception_full = stagger_left[(stagger_left['curve_L3_count'] != 0) |
                                                   (stagger_left['tangent_L3_count'] != 0)]

        stagger_left_exception_full.loc[stagger_left_exception_full['curve_L3'], 'exception type'] = 'Stagger'
        stagger_left_exception_full.loc[stagger_left_exception_full['tangent_L3'], 'exception type'] = 'Stagger'

        stagger_left_exception = (stagger_left_exception_full\
                                  .groupby(['exception type', 'curve_L3_id', 'tangent_L3_id'])\
                                    .agg({'Km': ['min', 'max'], 'maxValue': ['min', 'max'], 'Km_roundup': ['min', 'max']}))
        stagger_left_exception.columns = stagger_left_exception.columns.map('_'.join)
        stagger_left_exception = (stagger_left_exception\
                                  .assign(key=1)\
                                    .merge(stagger_left_exception_full.assign(key=1), on='key')\
                                        .query('`maxValue_max` == `maxValue` & `Km_roundup`.between(`Km_roundup_min`, `Km_roundup_max`)', engine='python')\
                                            .drop(columns=['maxValue_min', 'key', 'maxValue', 'stagger1', 'stagger2', 'stagger3', 'stagger4',
                                                           'curve_L3', 'tangent_L3', 'curve_L3_id', 'tangent_L3_id', 'curve_L3_count', 'tangent_L3_count'])\
                                                            .rename({'Km': 'maxLocation', 'Km_roundup_min': 'startKm',
                                                                     'Km_roundup_max': 'endKm', 'maxValue_max': 'maxValue'}, axis=1)\
                                                                        .reset_index()\
                                                                            .drop('index', axis=1))

        stagger_left_exception['length'] = stagger_left_exception['endKm'] - stagger_left_exception['startKm']
        stagger_left_exception = stagger_left_exception[['exception type', 'startKm', 'endKm', 'length', 'maxValue',
                                                         'maxLocation', 'track type', 'Section']]\
                                                            .sort_values('startKm') \
                                                                .drop_duplicates(subset=['exception type', 'startKm', 'endKm'])\
                                                                    .reset_index()\
                                                                        .drop('index', axis=1)

        # apply chain length
        stagger_left_exception['group'] = (stagger_left_exception['startKm'] - stagger_left_exception['endKm'].shift().fillna(0) >= chain_length).cumsum()
        stagger_left_exception = stagger_left_exception.groupby('group').apply(max_aggregate_group).reset_index(drop=True)

        # ------ defining alarm level depending on maxValue only ------
        # Finds L2 left stagger exception
        stagger_left_exception['curve_L2_min'] = 0
        stagger_left_exception['curve_L2_max'] = 0
        stagger_left_exception['tangent_L2_min'] = 0
        stagger_left_exception['tangent_L2_max'] = 0
        if (line == 'TML'):
            other_curve_L2_row_location = stagger_left_exception[(stagger_left_exception['track type'] == 'Curve') &
                                                                 (~stagger_left_exception['Section'].isin(['KSL', 'SCL']))].index
            stagger_left_exception.loc[other_curve_L2_row_location, 'curve_L2_min'] = curve_stagger_L2_min
            stagger_left_exception.loc[other_curve_L2_row_location, 'curve_L2_max'] = curve_stagger_L2_max

            other_tangent_L2_row_location = stagger_left_exception[(stagger_left_exception['track type'] == 'Tangent') &
                                                                   (~stagger_left_exception['Section'].isin(['KSL', 'SCL']))].index
            stagger_left_exception.loc[other_tangent_L2_row_location, 'tangent_L2_min'] = tangent_stagger_L2_min
            stagger_left_exception.loc[other_tangent_L2_row_location, 'tangent_L2_max'] = tangent_stagger_L2_max
        
            KSL_tangent_L2_row_location = stagger_left_exception[(stagger_left_exception['track type'] == 'Tangent') &
                                                                 (stagger_left_exception['Section'] == 'KSL')].index
            stagger_left_exception.loc[KSL_tangent_L2_row_location, 'tangent_L2_min'] = KSL_tangent_stagger_L2_min
            stagger_left_exception.loc[KSL_tangent_L2_row_location, 'tangent_L2_max'] = KSL_tangent_stagger_L2_max
        else:
            if (section in ['LOW', 'RAC', 'LMC']):
                other_curve_L2_row_location = stagger_left_exception[(stagger_left_exception['track type'] == 'Curve')].index
                stagger_left_exception.loc[other_curve_L2_row_location, 'curve_L2_min'] = curve_stagger_L2_min
                stagger_left_exception.loc[other_curve_L2_row_location, 'curve_L2_max'] = curve_stagger_L2_max

                other_tangent_L2_row_location = stagger_left_exception[(stagger_left_exception['track type'] == 'Tangent')].index
                stagger_left_exception.loc[other_tangent_L2_row_location, 'tangent_L2_min'] = tangent_stagger_L2_min
                stagger_left_exception.loc[other_tangent_L2_row_location, 'tangent_L2_max'] = tangent_stagger_L2_max
            else:
                other_curve_L2_row_location = stagger_left_exception[(stagger_left_exception['Section'] != 'SCL') &
                                                                     (stagger_left_exception['track type'] == 'Curve')].index
                stagger_left_exception.loc[other_curve_L2_row_location, 'curve_L2_min'] = curve_stagger_L2_min
                stagger_left_exception.loc[other_curve_L2_row_location, 'curve_L2_max'] = curve_stagger_L2_max

                other_tangent_L2_row_location = stagger_left_exception[(stagger_left_exception['Section'] != 'SCL') &
                                                                       (stagger_left_exception['track type'] == 'Tangent')].index
                stagger_left_exception.loc[other_tangent_L2_row_location, 'tangent_L2_min'] = tangent_stagger_L2_min
                stagger_left_exception.loc[other_tangent_L2_row_location, 'tangent_L2_max'] = tangent_stagger_L2_max
        
        stagger_left_exception.loc[((stagger_left_exception.curve_L2_min <= stagger_left_exception.maxValue) &
                                    (stagger_left_exception.maxValue < stagger_left_exception.curve_L2_max)), 'level'] = 'L2'
        stagger_left_exception.loc[((stagger_left_exception.tangent_L2_min <= stagger_left_exception.maxValue) &
                                    (stagger_left_exception.maxValue < stagger_left_exception.tangent_L2_max)), 'level'] = 'L2'
        # Finds L1 left stagger exception
        stagger_left_exception['curve_L1'] = float('inf')
        stagger_left_exception['tangent_L1'] = float('inf')

        if (line == 'TML'):
            other_curve_L1_row_location = stagger_left_exception[(stagger_left_exception['track type'] == 'Curve')].index
            stagger_left_exception.loc[other_curve_L1_row_location, 'curve_L1'] = curve_stagger_L1_min
            KSL_curve_L1_row_location = stagger_left_exception[(stagger_left_exception['track type'] == 'Curve') &
                                                               (stagger_left_exception['Section'] == 'KSL')].index
            stagger_left_exception.loc[KSL_curve_L1_row_location, 'curve_L1'] = KSL_curve_stagger_L1_min
            SCL_curve_L1_row_location = stagger_left_exception[(stagger_left_exception['track type'] == 'Curve') &
                                                               (stagger_left_exception['Section'] == 'SCL')].index
            stagger_left_exception.loc[SCL_curve_L1_row_location, 'curve_L1'] = SCL_curve_stagger_L1_min

            other_tangent_L1_row_location = stagger_left_exception[(stagger_left_exception['track type'] == 'Tangent')].index
            stagger_left_exception.loc[other_tangent_L1_row_location, 'tangent_L1'] = tangent_stagger_L1_min
            KSL_tangent_L1_row_location = stagger_left_exception[(stagger_left_exception['track type'] == 'Tangent') &
                                                                 (stagger_left_exception['Section'] == 'KSL')].index
            stagger_left_exception.loc[KSL_tangent_L1_row_location, 'tangent_L1'] = KSL_tangent_stagger_L1_min
            SCL_tangent_L1_row_location = stagger_left_exception[(stagger_left_exception['track type'] == 'Tangent') &
                                                                 (stagger_left_exception['Section'] == 'SCL')].index
            stagger_left_exception.loc[SCL_tangent_L1_row_location, 'tangent_L1'] = SCL_tangent_stagger_L1_min
        else:
            other_curve_L1_row_location = stagger_left_exception[(stagger_left_exception['track type'] == 'Curve')].index
            stagger_left_exception.loc[other_curve_L1_row_location, 'curve_L1'] = curve_stagger_L1_min
            other_tangent_L1_row_location = stagger_left_exception[(stagger_left_exception['track type'] == 'Tangent')].index
            stagger_left_exception.loc[other_tangent_L1_row_location, 'tangent_L1'] = tangent_stagger_L1_min

            if (section not in ['LOW', 'RAC', 'LMC']):
                SCL_curve_L1_row_location = stagger_left_exception[(stagger_left_exception['track type'] == 'Curve') &
                                                                   (stagger_left_exception['Section'] == 'SCL')].index
                stagger_left_exception.loc[SCL_curve_L1_row_location, 'curve_L1'] = SCL_curve_stagger_L1_min

                
                SCL_tangent_L1_row_location = stagger_left_exception[(stagger_left_exception['track type'] == 'Tangent') &
                                                                     (stagger_left_exception['Section'] == 'SCL')].index
                stagger_left_exception.loc[SCL_tangent_L1_row_location, 'tangent_L1'] = SCL_tangent_stagger_L1_min

        stagger_left_exception.loc[(stagger_left_exception['maxValue'] >= stagger_left_exception['curve_L1']), 'level'] = 'L1' 
        stagger_left_exception.loc[(stagger_left_exception['maxValue'] >= stagger_left_exception['tangent_L1']), 'level'] = 'L1'

        stagger_left_exception['level'] = stagger_left_exception['level'].replace('na', np.nan)
        stagger_left_exception['level'] = stagger_left_exception['level'].fillna('L3')
        stagger_left_exception = stagger_left_exception[['exception type', 'level', 'startKm', 'endKm','length',
                                                         'maxValue', 'maxLocation', 'track type']].reset_index()
        stagger_left_exception['id'] = (date + '_' + line + '_' + section + '_' + track + '_' + 'SL' +
                                        stagger_left_exception['index'].astype(str))
        stagger_left_exception = stagger_left_exception[['id', 'exception type', 'level', 'startKm', 'endKm',
                                                         'length', 'maxValue', 'maxLocation', 'track type']]
        # ------ defining alarm level depending on maxValue only ------
    else:
        stagger_left_exception = pd.DataFrame(columns=['id', 'exception type', 'level', 'startKm', 'endKm',
                                                       'length', 'maxValue', 'maxLocation', 'track type'])
        # ---------- Left stagger exception ----------

    # ---------- Right stagger exception ----------
    stagger_right['Km_roundup'] = stagger_right['Km'].round(3)

    # Labels all found exceptions L3 for now
    stagger_right['curve_L3_min'] = float('inf')
    stagger_right['tangent_L3_min'] = float('inf')
    
    if (line == 'TML'):
        other_curve_L3_min_row_location = stagger_right[(stagger_right['track type'] == 'Curve')].index
        stagger_right.loc[other_curve_L3_min_row_location, 'curve_L3_min'] = curve_stagger_L3_min
        KSL_curve_L3_min_row_location = stagger_right[(stagger_right['track type'] == 'Curve') &
                                                      (stagger_right['Section'] == 'KSL')].index
        stagger_right.loc[KSL_curve_L3_min_row_location, 'curve_L3_min'] = KSL_curve_stagger_L3_min
        SCL_curve_L3_min_row_location = stagger_right[(stagger_right['track type'] == 'Curve') &
                                                      (stagger_right['Section'] == 'SCL')].index
        stagger_right.loc[SCL_curve_L3_min_row_location, 'curve_L3_min'] = SCL_curve_stagger_L3_min

        other_tangentL3_min_row_location = stagger_right[(stagger_right['track type'] == 'Tangent')].index
        stagger_right.loc[other_tangentL3_min_row_location, 'tangent_L3_min'] = tangent_stagger_L3_min
        KSL_tangent_L3_min_row_location = stagger_right[(stagger_right['track type'] == 'Tangent') &
                                                        (stagger_right['Section'] == 'KSL')].index
        stagger_right.loc[KSL_tangent_L3_min_row_location, 'tangent_L3_min'] = KSL_tangent_stagger_L3_min
        SCL_tangent_L3_min_row_location = stagger_right[(stagger_right['track type'] == 'Tangent') &
                                                        (stagger_right['Section'] == 'SCL')].index
        stagger_right.loc[SCL_tangent_L3_min_row_location, 'tangent_L3_min'] = SCL_tangent_stagger_L3_min
    else:
        if (section in ['LOW', 'RAC', 'LMC']):
            other_curve_L3_min_row_location = stagger_right[(stagger_right['track type'] == 'Curve')].index
            stagger_right.loc[other_curve_L3_min_row_location, 'curve_L3_min'] = curve_stagger_L3_min
            other_tangentL3_min_row_location = stagger_right[(stagger_right['track type'] == 'Tangent')].index
            stagger_right.loc[other_tangentL3_min_row_location, 'tangent_L3_min'] = tangent_stagger_L3_min
        else:
            other_curve_L3_min_row_location = stagger_right[(stagger_right['track type'] == 'Curve')].index
            stagger_right.loc[other_curve_L3_min_row_location, 'curve_L3_min'] = curve_stagger_L3_min
            SCL_curve_L3_min_row_location = stagger_right[(stagger_right['track type'] == 'Curve') &
                                                          (stagger_right['Section'] == 'SCL')].index
            stagger_right.loc[SCL_curve_L3_min_row_location, 'curve_L3_min'] = SCL_curve_stagger_L3_min

            other_tangentL3_min_row_location = stagger_right[(stagger_right['track type'] == 'Tangent')].index
            stagger_right.loc[other_tangentL3_min_row_location, 'tangent_L3_min'] = tangent_stagger_L3_min
            SCL_tangent_L3_min_row_location = stagger_right[(stagger_right['track type'] == 'Tangent') &
                                                            (stagger_right['Section'] == 'SCL')].index
            stagger_right.loc[SCL_tangent_L3_min_row_location, 'tangent_L3_min'] = SCL_tangent_stagger_L3_min

        
    stagger_right['curve_L3'] = (stagger_right.maxValue <= -stagger_right.curve_L3_min)
    stagger_right['tangent_L3'] = (stagger_right.maxValue <= -stagger_right.tangent_L3_min)
    stagger_right['curve_L3_id'] = (stagger_right.curve_L3 != stagger_right.curve_L3.shift()).cumsum()
    stagger_right['tangent_L3_id'] = (stagger_right.tangent_L3 != stagger_right.tangent_L3.shift()).cumsum()
    stagger_right['curve_L3_count'] = stagger_right.groupby(['curve_L3', 'curve_L3_id']).cumcount(ascending=False) + 1
    stagger_right['tangent_L3_count'] = stagger_right.groupby(['tangent_L3', 'tangent_L3_id']).cumcount(ascending=False) + 1
    stagger_right.loc[~stagger_right['curve_L3'], 'curve_L3_count'] = 0
    stagger_right.loc[~stagger_right['tangent_L3'], 'tangent_L3_count'] = 0
    if stagger_right['curve_L3'].any() or stagger_right['tangent_L3'].any():

        stagger_right_exception_full = stagger_right[(stagger_right['curve_L3_count'] != 0) |
                                                     (stagger_right['tangent_L3_count'] != 0)]

        stagger_right_exception_full.loc[stagger_right_exception_full['curve_L3'], 'exception type'] = 'Stagger'
        stagger_right_exception_full.loc[stagger_right_exception_full['tangent_L3'], 'exception type'] = 'Stagger'

        stagger_right_exception = stagger_right_exception_full\
            .groupby(['exception type', 'curve_L3_id', 'tangent_L3_id'])\
            .agg({'Km': ['min', 'max'], 'maxValue': ['min', 'max'], 'Km_roundup': ['min', 'max']})
        stagger_right_exception.columns = stagger_right_exception.columns.map('_'.join)
        stagger_right_exception = (stagger_right_exception.assign(key=1)\
                                   .merge(stagger_right_exception_full.assign(key=1), on='key')\
                                    .query('`maxValue_min` == `maxValue` & `Km_roundup`.between(`Km_roundup_min`, `Km_roundup_max`)', engine='python')\
                                        .drop(columns=['maxValue_max', 'key', 'maxValue', 'stagger1', 'stagger2', 'stagger3', 'stagger4',
                                                       'curve_L3', 'tangent_L3', 'curve_L3_id', 'tangent_L3_id', 'curve_L3_count', 'tangent_L3_count'])\
                                                        .rename({'Km': 'maxLocation', 'Km_roundup_min': 'startKm',
                                                                 'Km_roundup_max': 'endKm','maxValue_min': 'maxValue'}, axis=1)\
                                                                    .reset_index()\
                                                                        .drop('index', axis=1))
        stagger_right_exception['length'] = stagger_right_exception['endKm'] - stagger_right_exception['startKm']
        stagger_right_exception = stagger_right_exception[['exception type', 'startKm', 'endKm', 'length',
                                                           'maxValue', 'maxLocation', 'track type', 'Section']]\
                                                            .sort_values('startKm')\
                                                                .drop_duplicates(subset=['exception type', 'startKm', 'endKm'])\
                                                                    .reset_index()\
                                                                        .drop('index', axis=1)

        # Applies chain length
        stagger_right_exception['group'] = (stagger_right_exception['startKm'] - stagger_right_exception['endKm'].shift().fillna(0) >= chain_length).cumsum()
        stagger_right_exception = (stagger_right_exception.groupby('group').apply(min_aggregate_group).reset_index(drop=True))

        # ------ defining alarm level depends on maxValue only ------
        # Finds L1 left stagger exception
        stagger_right_exception['curve_L2_min'] = 0
        stagger_right_exception['curve_L2_max'] = 0
        stagger_right_exception['tangent_L2_min'] = 0
        stagger_right_exception['tangent_L2_max'] = 0

        if (line == 'TML'):
            other_curve_L2_row_location = stagger_right_exception[(stagger_right_exception['track type'] == 'Curve') &
                                                                  (~stagger_right_exception['Section'].isin(['KSL', 'SCL']))].index
            stagger_right_exception.loc[other_curve_L2_row_location, 'curve_L2_min'] = curve_stagger_L2_min
            stagger_right_exception.loc[other_curve_L2_row_location, 'curve_L2_max'] = curve_stagger_L2_max

            other_tangent_L2_row_location = stagger_right_exception[(stagger_right_exception['track type'] == 'Tangent') &
                                                                    (~stagger_right_exception['Section'].isin(['KSL', 'SCL']))].index
            stagger_right_exception.loc[other_tangent_L2_row_location, 'tangent_L2_min'] = tangent_stagger_L2_min
            stagger_right_exception.loc[other_tangent_L2_row_location, 'tangent_L2_max'] = tangent_stagger_L2_max

            KSL_tangent_L2_row_location = stagger_right_exception[(stagger_right_exception['track type'] == 'Tangent') &
                                                                  (stagger_right_exception['Section'] == 'KSL')].index
            stagger_right_exception.loc[KSL_tangent_L2_row_location, 'tangent_L2_min'] = KSL_tangent_stagger_L2_min
            stagger_right_exception.loc[KSL_tangent_L2_row_location, 'tangent_L2_max'] = KSL_tangent_stagger_L2_max
        else:
            if (section in ['LOW', 'RAC', 'LMC']):
                other_curve_L2_row_location = stagger_right_exception[(stagger_right_exception['track type'] == 'Curve')].index
                stagger_right_exception.loc[other_curve_L2_row_location, 'curve_L2_min'] = curve_stagger_L2_min
                stagger_right_exception.loc[other_curve_L2_row_location, 'curve_L2_max'] = curve_stagger_L2_max

                other_tangent_L2_row_location = stagger_right_exception[(stagger_right_exception['track type'] == 'Tangent')].index
                stagger_right_exception.loc[other_tangent_L2_row_location, 'tangent_L2_min'] = tangent_stagger_L2_min
                stagger_right_exception.loc[other_tangent_L2_row_location, 'tangent_L2_max'] = tangent_stagger_L2_max
            else:
                other_curve_L2_row_location = stagger_right_exception[(stagger_right_exception['Section'] != 'SCL') &
                                                                      (stagger_right_exception['track type'] == 'Curve')].index
                stagger_right_exception.loc[other_curve_L2_row_location, 'curve_L2_min'] = curve_stagger_L2_min
                stagger_right_exception.loc[other_curve_L2_row_location, 'curve_L2_max'] = curve_stagger_L2_max

                other_tangent_L2_row_location = stagger_right_exception[(stagger_right_exception['Section'] != 'SCL') &
                                                                        (stagger_right_exception['track type'] == 'Tangent')].index
                stagger_right_exception.loc[other_tangent_L2_row_location, 'tangent_L2_min'] = tangent_stagger_L2_min
                stagger_right_exception.loc[other_tangent_L2_row_location, 'tangent_L2_max'] = tangent_stagger_L2_max

        stagger_right_exception.loc[((-stagger_right_exception.curve_L2_min >= stagger_right_exception.maxValue) & 
                                     (stagger_right_exception.maxValue > -stagger_right_exception.curve_L2_max)), 'level'] = 'L2'
        stagger_right_exception.loc[((-stagger_right_exception.tangent_L2_min >= stagger_right_exception.maxValue) &
                                     (stagger_right_exception.maxValue > -stagger_right_exception.tangent_L2_max)), 'level'] = 'L2'

        # Finds L1 left stagger exception
        stagger_right_exception['curve_L1'] = float('inf')
        stagger_right_exception['tangent_L1'] = float('inf')

        if (line == 'TML'):
            other_curve_L1_row_location = stagger_right_exception[(stagger_right_exception['track type'] == 'Curve')].index
            stagger_right_exception.loc[other_curve_L1_row_location, 'curve_L1'] = curve_stagger_L1_min
            KSL_curve_L1_row_location = stagger_right_exception[(stagger_right_exception['track type'] == 'Curve') &
                                                                (stagger_right_exception['Section'] == 'KSL')].index
            stagger_right_exception.loc[KSL_curve_L1_row_location, 'curve_L1'] = KSL_curve_stagger_L1_min
            SCL_curve_L1_row_location = stagger_right_exception[(stagger_right_exception['track type'] == 'Curve') &
                                                                (stagger_right_exception['Section'] == 'SCL')].index
            stagger_right_exception.loc[SCL_curve_L1_row_location, 'curve_L1'] = SCL_curve_stagger_L1_min

            other_tangent_L1_row_location = stagger_right_exception[(stagger_right_exception['track type'] == 'Tangent')].index
            stagger_right_exception.loc[other_tangent_L1_row_location, 'tangent_L1'] = tangent_stagger_L1_min
            KSL_tangent_L1_row_location = stagger_right_exception[(stagger_right_exception['track type'] == 'Tangent') &
                                                                (stagger_right_exception['Section'] == 'KSL')].index
            stagger_right_exception.loc[KSL_tangent_L1_row_location, 'tangent_L1'] = KSL_tangent_stagger_L1_min
            SCL_tangent_L1_row_location = stagger_right_exception[(stagger_right_exception['track type'] == 'Tangent') &
                                                                (stagger_right_exception['Section'] == 'SCL')].index
            stagger_right_exception.loc[SCL_tangent_L1_row_location, 'tangent_L1'] = SCL_tangent_stagger_L1_min
        else:
            other_curve_L1_row_location = stagger_right_exception[(stagger_right_exception['track type'] == 'Curve')].index
            stagger_right_exception.loc[other_curve_L1_row_location, 'curve_L1'] = curve_stagger_L1_min
            other_tangent_L1_row_location = stagger_right_exception[(stagger_right_exception['track type'] == 'Tangent')].index
            stagger_right_exception.loc[other_tangent_L1_row_location, 'tangent_L1'] = tangent_stagger_L1_min

            if (section not in ['LOW', 'RAC', 'LMC']):
                SCL_curve_L1_row_location = stagger_right_exception[(stagger_right_exception['track type'] == 'Curve') &
                                                                    (stagger_right_exception['Section'] == 'SCL')].index
                stagger_right_exception.loc[SCL_curve_L1_row_location, 'curve_L1'] = SCL_curve_stagger_L1_min
                SCL_tangent_L1_row_location = stagger_right_exception[(stagger_right_exception['track type'] == 'Tangent') &
                                                                      (stagger_right_exception['Section'] == 'SCL')].index
                stagger_right_exception.loc[SCL_tangent_L1_row_location, 'tangent_L1'] = SCL_tangent_stagger_L1_min


        stagger_right_exception.loc[(stagger_right_exception['maxValue'] <= -stagger_right_exception['curve_L1']), 'level'] = 'L1'
        stagger_right_exception.loc[(stagger_right_exception['maxValue'] <= -stagger_right_exception['tangent_L1']), 'level'] = 'L1'

        stagger_right_exception['level'] = stagger_right_exception['level'].replace('na', np.nan)
        stagger_right_exception['level'] = stagger_right_exception['level'].fillna('L3')
        stagger_right_exception = stagger_right_exception[['exception type', 'level', 'startKm', 'endKm', 'length',
                                                           'maxValue', 'maxLocation', 'track type']].reset_index()
        stagger_right_exception['id'] = date + '_' + line + '_' + section + '_' + track + '_' + 'SR' + stagger_right_exception['index'].astype(str)
        stagger_right_exception = stagger_right_exception[['id', 'exception type', 'level', 'startKm', 'endKm',
                                                           'length', 'maxValue', 'maxLocation', 'track type']]
        # ------ defining alarm level depends on maxValue only ------
    else:
        stagger_right_exception = pd.DataFrame(columns=['id', 'exception type', 'level', 'startKm', 'endKm',
                                                        'length', 'maxValue', 'maxLocation', 'track type'])
        # ---------- Right stagger exception ----------

    return stagger_left_exception, stagger_right_exception


def indicate_overlap_landmark(overlap, landmark, exception_table):
    """
    Writes the overlap and landmark columns in the input exception table.
    For each exception id, assigns the corresponding overlap and landmark values.

    Args:
        overlap (DataFrame): The overlap and tension length of the line.
        landmark (DataFrame): The landmark of the line.
        exception_table (DataFrame): The exception table to be edited.

    Returns:
        (DataFrame): The exception table with added overlap and landmark columns.
    """
    overlap_id = exception_table\
        .assign(key=1).merge(overlap.assign(key=1), on='key')\
        .query('`maxLocation`.between(`FromKM`, `ToKM`)', engine='python')\
        .reset_index(drop=True)[['id', 'Overlap', 'Tension Length']]

    exception_table = pd.merge(exception_table, overlap_id, on='id', how='outer')

    landmark_id = (exception_table\
                   .assign(key=1)\
                    .merge(landmark.assign(key=1), on='key')\
                        .query('`maxLocation`.between(`FromKM`, `ToKM`)', engine='python')\
                            .reset_index(drop=True)[['id', 'Landmark']])

    exception_table = pd.merge(exception_table, landmark_id, on='id', how='outer')
    exception_table['Landmark'] = exception_table['Landmark'].fillna(value='typical')
    return exception_table[['id', 'exception type', 'level', 'startKm', 'endKm', 'length', 'maxValue',
                            'maxLocation', 'track type', 'Overlap', 'Tension Length', 'Landmark']].drop_duplicates(subset=['id'], keep='first').reset_index(drop=True)


def main():
    line, section, track, raw = file_input()
    date = output_date(raw)
    track_type, overlap, landmark, threshold = load_metadata(line, section, track)

    # ----------- for debugging ----------
    # section = 'LMC'
    # line = 'EAL'
    # track = 'DN'
    # date = '20220610'
    # track_type = pd.read_excel('C:/Users/issac/Coding Projects/TOV_ExceptionReportGenerator/640/TOV RAW/metadata/EAL metadata.xlsx', sheet_name='DN track type')
    # threshold = pd.read_excel('C:/Users/issac/Coding Projects/TOV_ExceptionReportGenerator/640/TOV RAW/metadata/EAL metadata.xlsx', sheet_name='threshold')
    # overlap = pd.read_excel('C:/Users/issac/Coding Projects/TOV_ExceptionReportGenerator/640/TOV RAW/metadata/EAL metadata.xlsx', sheet_name='DN Tension Length')
    # raw = pd.read_csv('C:/Users/issac/Coding Projects/TOV_ExceptionReportGenerator/640/TOV RAW/2206/EAL_D1_2206.datac', sep=';').apply(lambda x: x.str.strip() if x.dtype == "object" else x)
    # ----------- for debugging ----------

    # Loads processed data into corresponding DataFrames
    raw, WH_cleaned_max, WH_cleaned_min, wear_min, stagger_left, stagger_right = \
        data_process(raw, track_type, line, track, section)

    # Loads exception tables
    low_height_exception = find_low_height_exception(WH_cleaned_min, threshold, date, line, section, track)
    high_height_exception = find_high_height_exception(WH_cleaned_max, threshold, date, line, section, track)
    wear_exception = find_wire_wear_exception(wear_min, threshold, date, line, section, track)
    stagger_left_exception, stagger_right_exception = find_stagger_exception(stagger_left, stagger_right, threshold, date, line, section, track)

    # ---------- indicate exceptions within overlap  and Landmark section ----------
    high_height_exception = indicate_overlap_landmark(overlap, landmark, high_height_exception)
    low_height_exception = indicate_overlap_landmark(overlap, landmark, low_height_exception)
    stagger_left_exception = indicate_overlap_landmark(overlap, landmark, stagger_left_exception)
    stagger_right_exception = indicate_overlap_landmark(overlap, landmark, stagger_right_exception)
    wear_exception = indicate_overlap_landmark(overlap, landmark, wear_exception)
    # ---------- indicate exceptions within overlap section ----------

    # ---------- Change Km to M ----------
    wear_exception = km_to_m(wear_exception)
    low_height_exception = km_to_m(low_height_exception)
    high_height_exception = km_to_m(high_height_exception)
    stagger_left_exception = km_to_m(stagger_left_exception)
    stagger_right_exception = km_to_m(stagger_right_exception)
    # ---------- Change Km to M ----------

    # ---------- sort the table by id ----------
    wear_exception = sort_table(wear_exception)
    low_height_exception = sort_table(low_height_exception)
    high_height_exception = sort_table(high_height_exception)
    stagger_left_exception = sort_table(stagger_left_exception)
    stagger_right_exception = sort_table(stagger_right_exception)
    # ---------- sort the table by id ----------

    # ---------- output results ----------
    print('Saving as Excel at', datetime.now())
    print('Done!')
    input('Press Enter to select save location')
    directory = filedialog.askdirectory()

    with pd.ExcelWriter(directory + '/' + date + '_' + line + '_' + track + '_' + 
                        section.replace('-', '_') + '_' + 'Exception Report.xlsx') as writer:
        wear_exception.to_excel(writer, sheet_name='wear exception', index=False)
        low_height_exception.to_excel(writer, sheet_name='low height exception', index=False)
        high_height_exception.to_excel(writer, sheet_name='high height exception', index=False)
        stagger_left_exception.to_excel(writer, sheet_name='stagger left exception', index=False)
        stagger_right_exception.to_excel(writer, sheet_name='stagger right exception', index=False)
    
    print('Saved at ' + directory)
    # ---------- output results ----------

if __name__ == "__main__":
    main()
