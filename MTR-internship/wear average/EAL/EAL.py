import pandas as pd
import math
import pandas as pd
from tkinter import filedialog
from datetime import datetime
import re

def read_data(raw_data_path, name):
    """
    Reads the selected .xlsx file. Combines the four wire columns into one column and returns the dataframe.

    Args:
        raw_data_path (str): The path to the .xlsx file.
        name (str): The name of the sheet to be returned.

    Returns:
        (DataFrame): The modified reqeusted sheet.
    """
    data = pd.read_excel(raw_data_path, name)
    data.columns = data.iloc[1]
    data = data.iloc[2:].reset_index().drop(columns='index').dropna(how='all')
    data = data.drop(data[data['LINE'].isin(['Date', 'LINE'])].index).reset_index(drop=True)
    return data[['TRACK', 'CHAINAGE', 'RWH1mm', 'RWH2mm', 'RWH3mm', 'RWH4mm']]

def get_individual_report(data, lookup_table):
    """
    Calculates the mean, standard deviation, and percentage of wear of the given data except EAL down track tension length X36.
    Stores the results in a DataFrame and returns it.

    Args:
        data (DataFrame): The given input data.
        lookup_table (DataFrame): The Tension Length lookup table for respective section and track of the line.

    Returns:
        (DataFrame): A DataFrame that contains the calculated mean, standard deviation, and percentage of wear at each tension length.
    """
    report = pd.DataFrame()
    for index, row in lookup_table.iterrows():
        if row['Tension Length'] == 'X36':
            break
        else:
            location = data[(data['CHAINAGE'] >= row['Overlap FromM']) & (data['CHAINAGE'] <= row['Overlap ToM'])].index
            frames = [data.loc[location, 'RWH1mm'], data.loc[location, 'RWH2mm'], data.loc[location, 'RWH3mm'], data.loc[location, 'RWH4mm']]
            new_df = pd.concat(frames)
            mean = new_df.mean()
            sd = new_df.std()
            if mean != 0:
                wear_percentage = ((((math.acos((mean-6.6)/6.6))*(6.6*6.6))-((6.6*(math.sin(math.acos((mean-6.6)/6.6))))*(mean-6.6)))/120)*100
            else:
                wear_percentage = 0
            temp = pd.DataFrame({
                'EAL': [row['Tension Length']], 
                'Mean of Remaining': [mean], 
                'Mean-SD of Remaining': [mean - sd], 
                'Mean-2SD of Remaining': [mean - 2 * sd], 
                'Mean-3SD of Remaining': [mean - 3 * sd], 
                '% of wear': [wear_percentage]
            })
            report = pd.concat([report, temp])
    return report.reset_index().drop(columns=['index'], axis=1)

def main():
    raw_data_path = filedialog.askopenfilename()
    print("raw:", raw_data_path)
    print(type(raw_data_path))
    print("Reading Files...")

    # ---------- Reads input data ----------
    EAL_up = read_data(raw_data_path, 'up')
    EAL_down = read_data(raw_data_path, 'down')
    RAC_up = read_data(raw_data_path, 'RAC up')
    RAC_down = read_data(raw_data_path, 'RAC down')
    LOW_s1 = read_data(raw_data_path, 'LOW S1')
    # ---------- Reads input data ----------

    pd.set_option('display.max_columns', 500)
    pd.set_option('display.width', 1000)

    # ---------- Loads lookup tables ----------
    EAL_lookup_up = pd.read_excel('./EAL/EAL.xlsx', sheet_name='EAL')
    EAL_lookup_up = EAL_lookup_up[['eal_up_from', 'eal_up_to', 'eal_up_tl']]\
        .rename({'eal_up_from': 'Overlap FromM', 'eal_up_to': 'Overlap ToM', 'eal_up_tl': 'Tension Length'}, axis=1).dropna(how='all')
    
    EAL_lookup_down = pd.read_excel('./EAL/EAL.xlsx', sheet_name='EAL')
    EAL_lookup_down = EAL_lookup_down[['eal_dn_from', 'eal_dn_to', 'eal_dn_tl']]\
        .rename({'eal_dn_from': 'Overlap FromM', 'eal_dn_to': 'Overlap ToM', 'eal_dn_tl': 'Tension Length'}, axis=1).dropna(how='all')
    
    RAC_lookup_up = pd.read_excel('./EAL/EAL.xlsx', sheet_name='RAC')
    RAC_lookup_up = RAC_lookup_up[['rac_up_from', 'rac_up_to', 'rac_up_tl']]\
        .rename({'rac_up_from': 'Overlap FromM', 'rac_up_to': 'Overlap ToM', 'rac_up_tl': 'Tension Length'}, axis=1).dropna(how='all')
    
    RAC_lookup_down = pd.read_excel('./EAL/EAL.xlsx', sheet_name='RAC')
    RAC_lookup_down = RAC_lookup_down[['rac_dn_from', 'rac_dn_to', 'rac_dn_tl']]\
        .rename({'rac_dn_from': 'Overlap FromM', 'rac_dn_to': 'Overlap ToM', 'rac_dn_tl': 'Tension Length'}, axis=1).dropna(how='all')
    
    LOW_s1_lookup = pd.read_excel('./EAL/EAL.xlsx', sheet_name='LOW S1')\
        .rename({'lows1_from': 'Overlap FromM', 'lows1_to': 'Overlap ToM', 'lows1_tl': 'Tension Length'}, axis=1)
    # ---------- Loads lookup tables ----------

    # ---------- Gets reports for each section and track ----------
    EAL_up_report = get_individual_report(EAL_up, EAL_lookup_up)
    EAL_down_report = get_individual_report(EAL_down, EAL_lookup_down)

    # EAL down track tension length X36
    X36_location = EAL_down[((EAL_down['CHAINAGE'] >= EAL_lookup_down.iloc[-2]['Overlap FromM']) & (EAL_down['CHAINAGE'] <= EAL_lookup_down.iloc[-2]['Overlap ToM'])) |
                            ((EAL_down['CHAINAGE'] >= EAL_lookup_down.iloc[-1]['Overlap FromM']) & (EAL_down['CHAINAGE'] <= EAL_lookup_down.iloc[-1]['Overlap ToM']))].index
    X36_frames = [EAL_down.loc[X36_location, 'RWH1mm'], EAL_down.loc[X36_location, 'RWH2mm'], EAL_down.loc[X36_location, 'RWH3mm'], EAL_down.loc[X36_location, 'RWH4mm']]
    X36_df = pd.concat(X36_frames)
    X36_mean = X36_df.mean()
    X36_sd = X36_df.std()
    if X36_mean != 0:
        X36_wear_percentage = ((((math.acos((X36_mean-6.6)/6.6))*(6.6*6.6))-((6.6*(math.sin(math.acos((X36_mean-6.6)/6.6))))*(X36_mean-6.6)))/120)*100
    else:
        X36_wear_percentage = 0
    X36_temp = pd.DataFrame({
        'EAL': ['X36'], 
        'Mean of Remaining': [X36_mean], 
        'Mean-SD of Remaining': [X36_mean - X36_sd], 
        'Mean-2SD of Remaining': [X36_mean - 2 * X36_sd], 
        'Mean-3SD of Remaining': [X36_mean - 3 * X36_sd], 
        '% of wear': [X36_wear_percentage]
        })
    
    # Combines up and down track into one report
    EAL_down_report = pd.concat([EAL_down_report, X36_temp])

    # ---------- Sorts the EAL table by the tension length ----------
    # H01 --> 01 --> T3 --> X36
    EAL_report = pd.concat([EAL_up_report, EAL_down_report], ignore_index=True)
    eal = EAL_report[EAL_report['EAL'].str[0] == 'H'].sort_values(by='EAL')

    eal_num = EAL_report[(EAL_report['EAL'].str[0] != 'H') &
                         (EAL_report['EAL'].str[0] != 'T') &
                         (EAL_report['EAL'].str[0] != 'X')]
    index = eal_num[eal_num['EAL'].astype(str).str.len() == 1].index
    eal_num['id'] = ''
    eal_num['id'] = 'A' + eal_num['EAL'].astype(str)
    eal_num.loc[index, 'id'] = 'A0' + eal_num.loc[index, 'EAL'].astype(str)
    
    eal_num = eal_num.sort_values(by='id').drop(columns='id')
    EAL_report = pd.concat([eal, eal_num, EAL_report[(EAL_report['EAL'] == 'T3') |
                                                     (EAL_report['EAL'] == 'X36')]], ignore_index=True)
    # ---------- Sorts the EAL table by the tension length ----------

    # ---------- Sorts the RAC table by the tension length ----------
    RAC_up_report = get_individual_report(RAC_up, RAC_lookup_up)
    RAC_down_report = get_individual_report(RAC_down, RAC_lookup_down)

    RAC_report = pd.concat([RAC_up_report, RAC_down_report], ignore_index=True)
    RAC_report['id'] = RAC_report['EAL'].str[1::].astype(int)
    RAC_report = RAC_report.sort_values(by='id').drop(columns='id')
    # ---------- Sorts the RAC table by the tension length ----------


    LOW_s1_report = get_individual_report(LOW_s1, LOW_s1_lookup)
    # ---------- Gets reports for each section and track ----------

    # ---------- Saving output .xlsx files ----------
    date_pattern = re.search(r'\d{6}', raw_data_path)
    date = date_pattern.group(0)

    print('Saving as Excel at', datetime.now())
    print('Done!')
    input('Please select save location')
    directory = filedialog.askdirectory()
    with pd.ExcelWriter(directory + '/' + date + ' EAL Report.xlsx') as writer:
            EAL_report.to_excel(writer, sheet_name='EAL', index=False)
            RAC_report.to_excel(writer, sheet_name='RAC', index=False)
            LOW_s1_report.to_excel(writer, sheet_name='LOW S1', index=False)

    print("Done!")
    print("Saved at " + directory)

if __name__ == "__main__":
    main()
